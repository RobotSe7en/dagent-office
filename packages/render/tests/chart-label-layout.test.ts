import { describe, expect, it } from 'vitest';
import { chartLegendPlan } from '../src/chart-label-layout.js';

describe('measured chart legend packing',()=>{
  it('wraps horizontal entries without splitting names or adding a trailing gap',()=>{
    const plan=chartLegendPlan([{width:40,height:10},{width:40,height:12},{width:20,height:8}],90,true);
    expect(plan).toEqual({width:88,height:22,entries:[{x:0,y:0},{x:48,y:0},{x:0,y:14}]});
    expect(Object.isFrozen(plan.entries)).toBe(true);
  });
  it('flows side legends into measured columns instead of clipping later entries',()=>{
    expect(chartLegendPlan([{width:40,height:10},{width:20,height:12},{width:30,height:8}],28,false)).toEqual({width:78,height:26,entries:[{x:0,y:0},{x:0,y:14},{x:48,y:0}]});
  });
  it('rejects nonfinite geometry and keeps empty entries bounded',()=>{
    expect(chartLegendPlan([],20,true)).toEqual({width:0,height:0,entries:[]});
    for(const limit of [0,-1,NaN,Infinity])expect(()=>chartLegendPlan([],limit,true)).toThrowError(expect.objectContaining({code:'ResourceLimitExceeded'}));
    expect(()=>chartLegendPlan([{width:Infinity,height:1}],30,true)).toThrowError(expect.objectContaining({code:'ResourceLimitExceeded'}));
  });
});
