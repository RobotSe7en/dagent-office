import { readFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { describe,expect,it } from 'vitest';
import { openDocument } from '@dagent-office/core';
import { importXlsx } from '@dagent-office/office';

describe('independent native chart label sources',()=>{
  it.each(['','-producer'])('preserves%s actual flags and current signed values through package reopen',async suffix=>{
    const bytes=readFileSync('packages/office/tests/fixtures/xlsx-chart-labels'+suffix+'.xlsx');
    if(!suffix)expect(execFileSync('python3',['-c',"import importlib.util,sys;s=importlib.util.spec_from_file_location('labels','scripts/check-xlsx-chart-labels.py');m=importlib.util.module_from_spec(s);s.loader.exec_module(m);sys.stdout.buffer.write(m.generate(m.inputs()))"])).toEqual(bytes);
    const model=openDocument(await importXlsx(bytes,{documentId:'native-labels'}),{kind:'workbook'}),snapshot=model.snapshot(),charts=snapshot.metadata.workbook.charts!;
    expect(charts).toHaveLength(12);
    if(!suffix)expect(charts.slice(0,6).map(chart=>chart.showLabelsOverMaximum)).toEqual([undefined,false,true,undefined,false,true]);
    for(const [index,sheet]of snapshot.metadata.workbook.sheets.entries())expect(model.getCell(sheet.sheetId,39,1)?.value).toEqual({type:'number',value:index<10?-2:2});
    expect(openDocument(model.serialize(),{kind:'workbook'}).snapshot()).toEqual(snapshot);
    // Calc's malformed child ordering is recorded by the official XSD audit,
    // separately from this semantic import/preservation check.
    const report=JSON.parse(readFileSync('docs/evidence/xlsx-chart-labels.json','utf8'));
    expect(report.schemas.source.invalidChartParts).toEqual([]);
    expect(report.schemas.producer.invalidChartParts.length).toBeGreaterThan(0);
  });
});
