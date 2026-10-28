import { readFile, mkdtemp, readdir, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { parsePresentation } from '../src/domain/schema.js';
import { threeReasonsPattern } from '../src/patterns/three-reasons/index.js';
import { checkSelectable, selectPattern } from '../src/selector/selector.js';
import { decideSlidePattern } from '../src/selector/selection-policy.js';
import { generate, generateFromFile } from '../src/pipeline.js';
import { readPptx, extractTexts } from './helpers.js';

const fixture = new URL('../examples/relationships/sequence-claim-evidence.invalid.json', import.meta.url);
const guidance = /sequence.*claim_evidence.*process.*timeline.*ページ/;
async function sample() { return parsePresentation(JSON.parse(await readFile(fixture, 'utf8'))); }

describe('sequence must not become evidence cards', () => {
  it('rejects three-reasons through the shared gate and direct bindings for every sequence relation', async () => {
    const spec = await sample();
    for (const logicalRelation of ['claim_evidence','process','timeline','cause_effect'] as const) {
      const slide = {...spec.slides[0]!, logicalRelation};
      expect(checkSelectable(threeReasonsPattern, slide).ok).toBe(false);
      expect(threeReasonsPattern.match(slide).reason).toMatch(/sequence/);
      expect(() => threeReasonsPattern.buildBindings(slide, 'standard')).toThrow(/sequence/);
    }
  });
  it('stops incompatible selection without changing the input, including manual and AI policy modes', async () => {
    const spec=await sample(), slide=spec.slides[0]!;
    const original=structuredClone(spec);
    expect(()=>selectPattern(slide)).toThrow(guidance);
    for (const mode of ['manual','assist','auto'] as const) {
      await expect(decideSlidePattern(slide,{mode,overrides:{[slide.id]:{patternId:'message-only'}}})).rejects.toThrow(guidance);
    }
    expect(spec).toEqual(original);
  });
  it('fails file generation before writing any output and names the offending slide', async () => {
    const dir=await mkdtemp(join(tmpdir(),'sdg-sequence-'));
    try {
      await expect(generateFromFile(fixture.pathname,{outDir:dir,includeTitleSlide:false})).rejects.toThrow(/sequence-claim-evidence/);
      expect(await readdir(dir)).toEqual([]);
    } finally { await rm(dir,{recursive:true,force:true}); }
  });
  it.each(['process','timeline'] as const)('renders an author-corrected %s and retains every evidence statement in notes', async logicalRelation => {
    const spec=await sample(); spec.slides[0]!.logicalRelation=logicalRelation;
    const original=structuredClone(spec);
    const result=await generate(spec,{includeTitleSlide:false,themeId:'portable'});
    expect(result.plan.slides[0]!.patternId).toBe('process');
    const pptx=await readPptx(result.pptxBuffer);
    const textElements=extractTexts(pptx.slideXmls[0]!);
    const texts=textElements.join('\n');
    const content=spec.slides[0]!.content!;
    if(content.kind!=='sequence') throw new Error('fixture must have sequence');
    let position=-1;
    for(const step of content.steps) {expect(texts).toContain(step.description);expect(textElements.indexOf(step.label)).toBeGreaterThan(position);position=textElements.indexOf(step.label);}
    const notes=await pptx.zip.file('ppt/notesSlides/notesSlide1.xml')!.async('string');
    for(const evidence of spec.slides[0]!.evidence) expect(extractTexts(notes).join('\n')).toContain(evidence.statement);
    expect(spec).toEqual(original);
  });
});
