import {describe,expect,it} from 'vitest';
import {clipChartValueSegment,chartStrokeOutset} from '../src/chart-clipping.js';

describe('bounded chart value-domain clipping',()=>{
  it('intersects original segments instead of clamping source point heights',()=>{
    const forward=clipChartValueSegment({x:0,value:-10},{x:30,value:20},0,10,0,1)!;
    expect(forward[0]!.x).toBeCloseTo(10,12);expect(forward[1]!.x).toBeCloseTo(20,12);expect(forward.map(point=>point.fraction)).toEqual([0,1]);
    const reverse=clipChartValueSegment({x:30,value:20},{x:0,value:-10},0,10,0,1)!;
    expect(reverse[0]!.x).toBeCloseTo(20,12);expect(reverse[1]!.x).toBeCloseTo(10,12);expect(reverse.map(point=>point.fraction)).toEqual([1,0]);
  });
  it('retains stroke outsets on subnormal axes without rounding expanded value endpoints',()=>{
    const line=clipChartValueSegment({x:0,value:0},{x:100,value:Number.MIN_VALUE*4},0,Number.MIN_VALUE*2,-.1,1.1)!;
    expect(line[0]).toEqual({x:0,fraction:0});expect(line[1]!.x).toBeCloseTo(55,12);expect(line[1]!.fraction).toBe(1.1);
  });
  it('bounds finite extremes and omits segments wholly beyond the same clip plane',()=>{
    const max=Number.MAX_VALUE,line=clipChartValueSegment({x:0,value:-max},{x:100,value:max},0,1,0,1)!;
    expect(line).toEqual([{x:50,fraction:0},{x:50,fraction:1}]);
    for(const sign of [-1,1])expect(clipChartValueSegment({x:0,value:sign*max},{x:100,value:sign*max/2},0,1,0,1)).toBeUndefined();
    expect(clipChartValueSegment({x:0,value:3},{x:100,value:3},0,1,0,1)).toBeUndefined();
  });
  it('measures square caps and miter joins while keeping inactive strokes unpainted',()=>{
    expect(chartStrokeOutset({stroke:'red','stroke-width':2,'stroke-linejoin':'miter','stroke-miterlimit':8})).toBe(8);
    expect(chartStrokeOutset({stroke:'red','stroke-width':2,'stroke-linecap':'square'})).toBe(Math.SQRT2);
    expect(chartStrokeOutset({stroke:'none','stroke-width':Number.MAX_VALUE})).toBe(0);
    expect(()=>chartStrokeOutset({stroke:'red','stroke-width':Number.MAX_VALUE})).toThrowError(expect.objectContaining({code:'ResourceLimitExceeded'}));
    expect(()=>clipChartValueSegment({x:0,value:Infinity},{x:1,value:0},0,1,0,1)).toThrowError(expect.objectContaining({code:'InvalidOperation'}));
  });
});
