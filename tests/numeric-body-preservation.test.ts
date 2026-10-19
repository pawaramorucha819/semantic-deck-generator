import { describe, expect, it } from 'vitest';
import { bigNumberPattern } from '../src/patterns/big-number/index.js';
import { beforeAfterPattern } from '../src/patterns/before-after/index.js';
import { chartWithMessagePattern } from '../src/patterns/chart-with-message/index.js';
import { checkSelectable, selectPattern } from '../src/selector/selector.js';
import { buildVisualPlan } from '../src/plan/visual-plan.js';
import { renderVisualPlan, toPptxBuffer } from '../src/render/renderer.js';
import { extractTexts, metricEvidence, pointsContent, readPptx, sequenceContent, slide } from './helpers.js';

const comparison = { id: 'c', type: 'comparison' as const, statement: 'Synthetic test values', data: { unit: '件', items: [{label:'A',value:1},{label:'B',value:2},{label:'C',value:3}] } };

describe('authored bodies with numeric evidence', () => {
  const points = pointsContent({id:'quality',label:'品質',description:'検索の正しさ',evidenceRefs:['m']}, {id:'time',label:'時間',description:'確認の所要時間'});
  const sequence = sequenceContent({id:'receive',label:'受付',description:'対象を登録'}, {id:'check',label:'確認',description:'内容を確認'});
  it.each([points, sequence])('rejects lossy numeric patterns for $kind through the shared gate and direct bindings', content => {
    for (const [pattern, logicalRelation, evidence] of [
      [bigNumberPattern, 'claim_evidence', [metricEvidence('m','仮の値',42)]],
      [chartWithMessagePattern, 'timeline', [comparison]],
      [beforeAfterPattern, 'before_after', [metricEvidence('m','前',42),metricEvidence('n','後',21)]],
    ] as const) {
      const input = slide({content, logicalRelation, evidence:[...evidence]});
      expect(checkSelectable(pattern, input).ok).toBe(false);
      expect(() => pattern.buildBindings(input, "standard")).toThrow();
    }
  });
  it('keeps every point and ordered step in the generated slide, and retains evidence provenance', async () => {
    const slides = [slide({id:'points',content:points,evidence:[metricEvidence('m','仮の値',42)]}),slide({id:'timeline',logicalRelation:'timeline',content:sequence,evidence:[comparison]})];
    expect(slides.map(s=>selectPattern(s).patternId)).toEqual(['three-reasons','process']);
    const plan=buildVisualPlan({title:'Regression',slides});
    const result=renderVisualPlan(plan,{includeTitleSlide:false});
    const xmls=(await readPptx(await toPptxBuffer(result.pptx))).slideXmls;
    for (const [i, items] of [[0,points.kind==='points'?points.items:[]],[1,sequence.kind==='sequence'?sequence.steps:[]]] as const) {
      const text=extractTexts(xmls[i]!).join('\n');
      for (const item of items) { expect(text).toContain(item.label); expect(text).toContain(item.description); }
      expect(plan.slides[i]!.provenance?.evidence).toEqual(slides[i]!.evidence.map(({id,type,statement}) => ({id,type,statement})));
    }
  });
  it('retains evidence-only numeric selection', () => {
    expect(selectPattern(slide({evidence:[metricEvidence('m','値',42)]})).patternId).toBe('big-number');
    expect(selectPattern(slide({logicalRelation:'timeline',evidence:[comparison]})).patternId).toBe('chart-with-message');
  });
});
