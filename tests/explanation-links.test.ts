import { readFile } from 'node:fs/promises';
import { describe, it, expect } from 'vitest';
import { parsePresentation } from '../src/domain/schema.js';
import { contentItemEntries } from '../src/domain/content.js';
import { buildVisualPlan, parseVisualPlan, serializeVisualPlan } from '../src/plan/visual-plan.js';
import { resolveExplanationLinks } from '../src/plan/explanation-links.js';
import { renderVisualPlan, toPptxBuffer } from '../src/render/renderer.js';
import { readPptx, extractTexts } from './helpers.js';

async function sample() {
 return JSON.parse(await readFile(new URL('../examples/relationships/six-relationships.json', import.meta.url), 'utf8'));
}
describe('explicit object explanation references', () => {
 it('rejects missing and duplicate references, including slides without content', async () => {
  const spec = await sample();
  spec.slides[0].explanation = {readerQuestion:'何を示すか',sections:[{heading:'対象',body:'説明',nodeRefs:['missing']}]};
  expect(() => parsePresentation(spec)).toThrow(/未定義/);
  delete spec.slides[0].content;
  expect(() => parsePresentation(spec)).toThrow(/未定義/);
  spec.slides[0].explanation.sections[0].nodeRefs=['missing','missing'];
  expect(() => parsePresentation(spec)).toThrow(/重複/);
 });
 it('matches visible object labels and section headings in all six relationships without mutating input', async () => {
  const spec = await sample();
  for (const s of spec.slides) {
   const entries=contentItemEntries(s.content);
   entries.forEach(({item},i)=>{item.id=`node-${i}`; item.label=`要素${i+1}`; delete item.description;});
   s.explanation={readerQuestion:'対象を確認する',sections:[{heading:'要素1',body:'この対象の具体的な説明。',nodeRefs:['node-0']},{heading:'注意',body:'これは全体に共通の条件。'}]};
  }
  const parsed=parsePresentation(spec); const before=JSON.stringify(parsed);
  const plan=buildVisualPlan(parsed); expect(JSON.stringify(parsed)).toBe(before);
  const result=renderVisualPlan(parseVisualPlan(JSON.parse(serializeVisualPlan(plan))),{includeTitleSlide:false});
  expect(result.issues).toEqual([]);
  const xmls=(await readPptx(await toPptxBuffer(result.pptx))).slideXmls;
  for (const xml of xmls) {
   const texts=extractTexts(xml).join('\n');
   expect(texts.match(/\[1\] 要素1/g)?.length).toBeGreaterThanOrEqual(2);
   expect(texts).toContain('全体｜注意');
  }
  expect(xmls[2]).not.toContain('STEP 1');
  expect(xmls[5]).toContain('cycle-step-4-to-1');
 });
 it('does not infer links when none are authored', async()=>{
  const s=(await sample()).slides[0];
  s.explanation={readerQuestion:'問い',sections:[{heading:'概要',body:'全体の説明'}]};
  expect(resolveExplanationLinks(s)).toBe(s);
 });
});
