"""Independent basic chart label limits, negative fills and native PDF evidence.

Source options and actual Calc resaves are distinct. PDF text/paint observations
do not imply Excel/WPS machine verification or whole-page fidelity.
"""
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile
from io import BytesIO
from hashlib import sha256
from copy import deepcopy
import argparse
import importlib.util
import json
import os
import re
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location('chart_labels_colors', ROOT / 'scripts/check-xlsx-chart-colors.py')
colors = importlib.util.module_from_spec(spec); spec.loader.exec_module(colors)
C, A, S = colors.C, colors.A, '{' + colors.native.fixture.S + '}'
REPORT = ROOT / 'docs/evidence/xlsx-chart-labels.json'
FIXTURES = ROOT / 'packages/office/tests/fixtures'


def audit(data, path, standards, producer=False):
    """Keep the producer's known invalid order visible; audit the rest too.

    Only a schema verification copy is reordered after the original has failed.
    The native bytes, SDK input and invalid-part classification never change.
    """
    from lxml import etree as E
    spec=importlib.util.spec_from_file_location('labels_schema_audit',ROOT/'scripts/check-xlsx-output.py')
    checker=importlib.util.module_from_spec(spec);spec.loader.exec_module(checker)
    issues=[]; schema=standards[C[1:-1]]
    class ChartObserver:
        def assertValid(self,root):
            try:schema.assertValid(root)
            except E.DocumentInvalid:
                assert producer
                reference=root.find('.//'+C+'cat/'+C+'strRef/'+C+'f').text
                assert reference.split('!')[0].strip("'") in ['Case'+str(index) for index in range(1,len(inputs())+1)]
                copy=deepcopy(root); changed=False
                for labels in copy.findall('.//'+C+'dLbls'):
                    points=labels.findall(C+'dLbl')
                    if points and labels[0].tag==C+'numFmt':
                        for point in points:labels.remove(point)
                        for index,point in enumerate(points):labels.insert(index,point)
                        changed=True
                assert changed
                schema.assertValid(copy)
                issues.append(dict(reference=reference,issue='numFmt precedes point dLbl elements; original fails official XSD, remaining structure passes the verification projection'))
    selected=dict(standards)
    if producer:selected[C[1:-1]]=ChartObserver()
    path.write_bytes(data);roots,_=checker.audit(path,selected)
    return dict(xmlParts=len(roots),invalidChartParts=sorted(issues,key=lambda issue:issue['reference']))


def pdf_observations(pdf, work, prefix):
    pages=ET.fromstring(colors.native.shared.run(['pdftotext','-bbox',str(pdf),'-'])).findall('.//'+colors.HTML+'page')
    assert len(pages)==len(inputs()); records=[]
    for index,page in enumerate(pages,1):
        svg=work/f'{prefix}-{index}.svg';colors.native.shared.run(['pdftocairo','-f',str(index),'-l',str(index),'-svg',str(pdf),str(svg)])
        paints=set()
        for node in ET.fromstring(svg.read_bytes()).iter():
            match=re.fullmatch(r'rgb\(([^)]+)\)',node.get('fill',''))
            if not match:continue
            rgb=[round(float(value.strip().rstrip('%'))*(2.55 if '%' in value else 1)) for value in match[1].split(',')]
            assert len(rgb)==3 and all(0<=value<=255 for value in rgb)
            paints.add(''.join(f'{value:02x}' for value in rgb))
        records.append(dict(words=[dict(text=node.text,boundsPt=[round(float(node.get(name)),3) for name in ['xMin','yMin','xMax','yMax']]) for node in page.findall(colors.HTML+'word')],fillColors=sorted(paints)))
    return records


def inputs():
    return ([dict(kind=kind, over=over, invert=False, fill='series') for kind in ['line', 'column'] for over in [None, False, True]]
            + [dict(kind='column', over=False, invert=invert, fill=fill) for invert in [False, True] for fill in [None, 'point']]
            + [dict(kind='pie', over=None, invert=False, fill='series', leader=leader) for leader in [False, True]])


def generate(cases):
    data = colors.generate([dict(kind=case['kind'], vary=False, series=1, **({'fill': case['fill']} if case['fill'] else {})) for case in cases])
    with ZipFile(BytesIO(data)) as archive: parts = {name: archive.read(name) for name in archive.namelist()}
    for index, case in enumerate(cases, 1):
        path = f'xl/charts/chart{index}.xml'; root = ET.fromstring(parts[path]); chart = root.find(C + 'chart'); series = root.find('.//' + C + 'ser')
        graph = chart.find('.//' + C + {'line':'lineChart','column':'barChart','pie':'pieChart'}[case['kind']])
        for parent in [graph,series]:
            for old in parent.findall(C+'dLbls'): parent.remove(old)
        labels = ET.Element(C + 'dLbls')
        if case['kind']=='pie': series.insert(next(i for i,node in enumerate(series) if node.tag==C+'cat'),labels)
        else: graph.insert(list(graph).index(series)+1, labels)
        ET.SubElement(labels, C + 'dLblPos', {'val':'outEnd' if case['kind'] != 'line' else 't'})
        for name in ['showLegendKey','showVal','showCatName','showSerName','showPercent','showBubbleSize']: ET.SubElement(labels,C+name,{'val':'1' if name=='showVal' else '0'})
        if case.get('leader') is not None: ET.SubElement(labels, C + 'showLeaderLines', {'val':str(int(case['leader']))})
        if case['over'] is not None:
            for old in chart.findall(C + 'showDLblsOverMax'): chart.remove(old)
            ET.SubElement(chart, C + 'showDLblsOverMax', {'val':str(int(case['over']))})
        else:
            for old in chart.findall(C + 'showDLblsOverMax'): chart.remove(old)
        values = [-2, 1, 6] if case['kind'] != 'pie' else [2, 1, 6]
        for cache in root.findall('.//' + C + 'val/' + C + 'numRef/' + C + 'numCache'):
            for point, value in zip(cache.findall(C+'pt'), values): point.find(C+'v').text = str(value)
        if case['kind'] != 'pie':
            axis = root.find('.//' + C + 'valAx'); scaling = axis.find(C + 'scaling')
            for old in list(scaling):
                if old.tag in [C+'min', C+'max']: scaling.remove(old)
            ET.SubElement(scaling,C+'max',{'val':'4' if index<=6 else '8'}); ET.SubElement(scaling,C+'min',{'val':'0' if index<=6 else '-4'})
            for old in series.findall(C+'invertIfNegative'): series.remove(old)
            if case['kind']=='column': series.insert(next(i for i,node in enumerate(series) if node.tag in [C+'dPt',C+'dLbls',C+'cat']),ET.Element(C+'invertIfNegative',{'val':str(int(case['invert']))}))
        parts[path] = ET.tostring(root,encoding='unicode')
        path=f'xl/worksheets/sheet{index}.xml'; sheet=ET.fromstring(parts[path])
        for row,value in zip(range(40,43),values): sheet.find(S+f'sheetData/{S}row/{S}c[@r="B{row}"]/{S}v').text=str(value)
        parts[path]=ET.tostring(sheet,encoding='unicode')
    return colors.native.fixture.zipped(parts)


def main():
    parser=argparse.ArgumentParser(description=__doc__); parser.add_argument('--record',action='store_true'); args=parser.parse_args(); native=colors.native
    assert sha256((ROOT/'.cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz').read_bytes()).hexdigest()==native.ARCHIVE_HASH
    os.environ['PATH']=str(ROOT/'.cache/acceptance-office/bin')+os.pathsep+os.environ['PATH']; os.environ['LD_LIBRARY_PATH']=str(ROOT/'.cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu')
    assert native.shared.run(['libreoffice','--version'])==native.BUILD
    font=next(font for font in json.loads((ROOT/'examples/xlsx/reference.json').read_text())['environment']['fonts'] if font['file']=='LiberationSans-Regular.ttf')
    family,filename=native.shared.run(['fc-match','-f','%{family}|%{file}','Liberation Sans:style=Regular']).split('|')
    assert 'Liberation Sans' in family and sha256(Path(filename).read_bytes()).hexdigest()==font['sha256']
    cases=inputs();data=generate(cases)
    with TemporaryDirectory(dir=ROOT/'.cache') as temporary:
        work=Path(temporary); source=work/'labels.xlsx';source.write_bytes(data);out=work/'out';out.mkdir()
        for kind in ['pdf','xlsx:Calc MS Excel 2007 XML']: native.shared.run(['libreoffice','-env:UserInstallation='+(work/'profile').as_uri(),'--headless','--convert-to',kind,'--outdir',str(out),str(source)])
        producer=(out/source.name).read_bytes(); reopened=work/'reopened';reopened.mkdir()
        native.shared.run(['libreoffice','-env:UserInstallation='+(work/'profile-reopen').as_uri(),'--headless','--convert-to','pdf','--outdir',str(reopened),str(out/source.name)])
        observations=[dict(**case,sourcePdf=source_pdf,producerPdf=producer_pdf) for case,source_pdf,producer_pdf in zip(cases,pdf_observations(out/'labels.pdf',work,'source'),pdf_observations(reopened/'labels.pdf',work,'producer'))]
        with ZipFile(BytesIO(producer)) as archive:
            for index,record in enumerate(observations,1):
                chart=ET.fromstring(archive.read(f'xl/charts/chart{index}.xml'))
                record['producerFlags']={name:[node.attrib for node in chart.findall('.//'+C+name)] for name in ['showDLblsOverMax','invertIfNegative','showLeaderLines']}
        spec=importlib.util.spec_from_file_location('labels_schemas',ROOT/'scripts/check-xlsx-output.py');checker=importlib.util.module_from_spec(spec);spec.loader.exec_module(checker)
        with checker.schemas() as standards:
            schema_report=dict(source=audit(data,work/'audit-source.xlsx',standards),producer=audit(producer,work/'audit-producer.xlsx',standards,True))
        report=dict(scope=__doc__,office=native.BUILD,archiveSha256=native.ARCHIVE_HASH,fontManifest=[font],sourceSha256=sha256(data).hexdigest(),cases=observations,schemas=schema_report,
                    limits=['Calc ignores column label-limit flags, rewrites invertIfNegative to false and can display the negative value by magnitude. SDK keeps source signs and separately follows the chosen white-negative-fill policy.',
                            'The recorded actual producer parts fail official XSD child order. Their bytes are retained and never counted as schema-conformant; a verification-only projection checks the remaining schema and package relationships.',
                            'PDF text and color observations are separate from SDK layout and do not prove whole-chart pixel equivalence or Excel/WPS behavior.'])
        if args.record:
            (FIXTURES/'xlsx-chart-labels.xlsx').write_bytes(data);(FIXTURES/'xlsx-chart-labels-producer.xlsx').write_bytes(producer);REPORT.write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n')
        else:
            assert (FIXTURES/'xlsx-chart-labels.xlsx').read_bytes()==data;assert json.loads(REPORT.read_text())==report
    print(f'Verified {len(cases)} independent chart label and negative-fill cases')


if __name__=='__main__':main()
