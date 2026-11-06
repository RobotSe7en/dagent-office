import type { WorkbookSnapshot } from '@doffi/core';
import { createWorkbookSurface } from '../../../packages/render/src/workbook/surface.js';
import { WorkbookContentSources } from '../../../packages/render/src/workbook/content-source.js';
import { createWorkbookChartView } from '../../../packages/render/src/workbook/chart-view.js';
import type { BasicChartData } from '../../../packages/render/src/chart.js';

export async function measureChartClipping(snapshot:WorkbookSnapshot){
  const container=document.createElement('div');document.body.append(container);
  const css='.clip-solid{--office-fill-kind:solid;--office-fill-color:#ff00ff;--office-line-kind:solid;--office-line-color:#ff00ff;--office-line-width:2;--office-marker-fill-kind:solid;--office-marker-fill-color:#ff00ff;--office-marker-line-kind:none;--office-invert-negative:false}.clip-gradient,.clip-reverse,.clip-diagonal{--office-fill-kind:linear;--office-fill-angle:0;--office-fill-stop-count:2;--office-fill-stop-0-position:0;--office-fill-stop-0-color:#ff0000;--office-fill-stop-1-position:1;--office-fill-stop-1-color:#0000ff;--office-line-kind:none}.clip-reverse{--office-fill-angle:180}.clip-diagonal{--office-fill-angle:45}';
  const surface=await createWorkbookSurface(container,{...snapshot,styles:[...snapshot.styles,{path:'chart-clipping.css',css}]},{widthPt:520,heightPt:340}),sources=new WorkbookContentSources(surface.document,snapshot),charts=snapshot.metadata.workbook.charts!,projection=sources.project(surface.source,charts.map(chart=>chart.nodeId)),before=surface.source.innerHTML;
  const max=Number.MAX_VALUE;
  const cases=[
    {kind:'line',values:[-max,.5,max]}, {kind:'line',values:[max,.5,-max],reverse:true},
    {kind:'line',values:[-max,null,max],blank:'gap'}, {kind:'line',values:[-max,null,max],blank:'span'},
    {kind:'line',values:[-max,null,max],blank:'zero',error:true}, {kind:'line',values:[-max,null,max],blank:'zero'},
    {kind:'column',values:[-max,.5,max]}, {kind:'column',values:[-max,.5,max],reverse:true},
    {kind:'column',values:[-max,.5,max],gradient:'clip-gradient'}, {kind:'column',values:[-max,.5,max],gradient:'clip-reverse'},
    {kind:'line',values:[-.005,.5,1.005]},
  ] as const;
  try{
    const results=[];
    for(const [index,item]of cases.entries()){
      const original=charts.find(chart=>chart.kind===item.kind)!,source=projection.get(original.nodeId)!;
      const appearance={...original,series:[original.series[0]!],minimum:0,maximum:1,majorUnit:.25,titleVisible:false,showCategoryAxis:false,showValueAxis:false,showLabels:false,legend:'none' as const,gridCssClass:'',categoryGridCssClass:'',reverseValues:'reverse'in item&&item.reverse,blankAs:'blank'in item?item.blank:'gap' as const,
        seriesStyles:[original.series[0]!].map(series=>({seriesId:series.seriesId,cssClass:'gradient'in item?item.gradient:'clip-solid',marker:'circle' as const,markerSize:7})),pointStyles:[]};
      const data:BasicChartData={categories:['P0','P1','P2'],series:[{seriesId:original.series[0]!.seriesId,name:'Values',values:item.values,...('error'in item?{missing:[undefined,'gap' as const,undefined]}:{})}]};
      const result=createWorkbookChartView(source,appearance,data,480,300,'clipping-'+index,{formatNumber:String});surface.view.replaceChildren(result.element);
      const svg=result.element.querySelector('svg')!,frame=svg.querySelector('rect')!,plot={x:frame.x.baseVal.value,y:frame.y.baseVal.value,width:frame.width.baseVal.value,height:frame.height.baseVal.value};
      const walker=surface.document.createTreeWalker(result.element,0xffff_ffff);let actual=1;while(walker.nextNode())actual++;
      const clone=svg.cloneNode(true)as SVGSVGElement;clone.setAttribute('width','480');clone.setAttribute('height','300');
      const url=URL.createObjectURL(new Blob([new XMLSerializer().serializeToString(clone)],{type:'image/svg+xml'})),image=new Image();let painted=0,escaped=0,gradientPixels:number[][]=[];
      try{
        image.src=url;await image.decode();const canvas=document.createElement('canvas');canvas.width=480;canvas.height=300;const context=canvas.getContext('2d')!;context.drawImage(image,0,0);const pixels=context.getImageData(0,0,480,300).data;
        for(let y=0;y<300;y++)for(let x=0;x<480;x++){const at=(y*480+x)*4;if(pixels[at]!>180&&pixels[at+1]!<50&&pixels[at+2]!>180&&pixels[at+3]!>180){painted++;if(x<plot.x-1||x>plot.x+plot.width+1||y<plot.y-1||y>plot.y+plot.height+1)escaped++;}}
        if('gradient'in item){const bar=svg.querySelectorAll<SVGRectElement>('[data-sdk-chart-point=column]')[2]!;gradientPixels=[.25,.75].map(fraction=>[...context.getImageData(Math.floor(bar.x.baseVal.value+bar.width.baseVal.value*fraction),Math.floor(plot.y+plot.height/2),1,1).data]);}
      }finally{image.removeAttribute('src');URL.revokeObjectURL(url);}
      const path=svg.querySelector('[data-sdk-chart-line]')?.getAttribute('d')??'',coordinates=[...path.matchAll(/[ML]([-+.\dEe]+)\s+([-+.\dEe]+)/g)].map(point=>({x:Number(point[1]),y:Number(point[2])}));
      const stable=surface.view.innerHTML,errors=[];
      if(index===0||index===8)for(const failure of ['nodes','diagonal-gradient']){
        const rejected=failure==='nodes'?appearance:{...appearance,kind:'column'as const,seriesStyles:appearance.seriesStyles.map(series=>({...series,cssClass:'clip-diagonal'}))};
        try{createWorkbookChartView(source,rejected,data,480,300,'clip-failure',{formatNumber:String,...(failure==='nodes'?{maxNodes:result.peakNodes-1}:{})});errors.push('accepted');}catch(error){errors.push((error as{code:string}).code);}
      }
      results.push({index,kind:item.kind,plot,coordinates,path,painted,escaped,gradientPixels,markers:svg.querySelectorAll('[data-sdk-chart-marker-role=point]').length,nodes:result.nodes,actual,invalid:/NaN|Infinity/.test(svg.outerHTML),errors,preserved:stable===surface.view.innerHTML});
    }
    return{results,unchanged:before===surface.source.innerHTML,cleaned:!surface.source.querySelector('[data-office-view=workbook-chart-label-measurement]')};
  }finally{projection.dispose();sources.dispose();surface.dispose();container.remove();}
}
