from pathlib import Path
import json
import zipfile
import xml.etree.ElementTree as E

ROOT=Path(__file__).resolve().parent.parent
OUT=ROOT/'.source-test/ui-regression'
NS={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

def canonical(element):
    return (element.tag,tuple(sorted(element.attrib.items())),element.text,tuple(canonical(c) for c in element))

def snapshot(path):
    with zipfile.ZipFile(path) as z:
        style_root=E.fromstring(z.read('xl/styles.xml'))
        tables={name:list(style_root.find('s:'+name,NS)) for name in ('fonts','fills','borders','cellStyleXfs','cellXfs')}
        numbers={x.get('numFmtId'):x.get('formatCode') for x in style_root.findall('s:numFmts/s:numFmt',NS)}
        def xf(element):
            attrs=dict(element.attrib)
            # Excel rewrites redundant apply* flags and deduplicates XF records
            # on save. Compare the resolved font/fill/border/protection instead.
            for key in list(attrs):
                if key.startswith('apply'):attrs.pop(key)
            for field,table in [('fontId','fonts'),('fillId','fills'),('borderId','borders')]:
                if field in attrs:attrs[field]=canonical(tables[table][int(attrs[field])])
            if 'numFmtId' in attrs:attrs['numFmtId']=numbers.get(attrs['numFmtId'],attrs['numFmtId'])
            if 'xfId' in attrs:attrs['xfId']=xf(tables['cellStyleXfs'][int(attrs['xfId'])])
            return (tuple(sorted(attrs.items())),tuple(canonical(c) for c in element))
        formats=[xf(x) for x in tables['cellXfs']]
        def styled_attrs(element):
            attrs=dict(element.attrib)
            for field in ('s','style'):
                if field in attrs:attrs[field]=formats[int(attrs[field])]
            attrs.pop('spans',None)
            return attrs
        strings=[''.join(t.text or '' for t in x.findall('.//s:t',NS)) for x in E.fromstring(z.read('xl/sharedStrings.xml'))]
        sheets={}
        for name in z.namelist():
            if name.startswith('xl/worksheets/sheet') and name.endswith('.xml'):
                root=E.fromstring(z.read(name))
                cells={}
                for c in root.findall('.//s:sheetData/s:row/s:c',NS):
                    v=c.find('s:v',NS)
                    f=c.find('s:f',NS)
                    value=v.text if v is not None else None
                    if c.get('t')=='s' and value is not None:value=strings[int(value)]
                    if c.get('t')=='inlineStr':value=''.join(c.find('s:is',NS).itertext())
                    cells[c.get('r')]=(value, f.text if f is not None else None,formats[int(c.get('s','0'))])
                sheets[name]={'cells':cells,
                  'default_style':formats[0],
                  'default_width':root.find('s:sheetFormatPr',NS).get('defaultColWidth','9'),
                  'rows':{r.get('r'):styled_attrs(r) for r in root.findall('.//s:sheetData/s:row',NS)},
                  'columns':[styled_attrs(c) for c in root.findall('.//s:cols/s:col',NS)],
                  'merges':[c.attrib for c in root.findall('.//s:mergeCell',NS)],
                  'validations':[E.tostring(c).decode() for c in root.findall('.//s:dataValidation',NS)]}
        return sheets

before=snapshot(ROOT/'workbook/근무일수계산.xlsm')
after=snapshot(OUT/'근무일수계산.xlsm')
allowed={f'{c}{r}' for c in 'EF' for r in (9,11,13,15,17,19,21)}
report={'unexpected_cell_changes':{},'unexpected_row_changes':{},'other_layout_preserved':True}
for name,old in before.items():
    new=after[name]
    changed={a for a in old['cells'].keys()|new['cells'].keys() if old['cells'].get(a,(None,None,old['default_style']))!=new['cells'].get(a,(None,None,new['default_style']))}
    if name=='xl/worksheets/sheet1.xml':changed-=allowed
    if changed:report['unexpected_cell_changes'][name]=sorted(changed)
    changed_rows={r for r in old['rows'].keys()|new['rows'].keys() if old['rows'].get(r)!=new['rows'].get(r)}
    if name=='xl/worksheets/sheet1.xml':changed_rows.discard('21')
    if changed_rows:report['unexpected_row_changes'][name]=sorted(changed_rows)
    def columns(sheet):
        defaults={'width':sheet['default_width'],'style':sheet['default_style']}
        result={i:defaults for i in range(1,16385)}
        for spec in sheet['columns']:
            attrs={k:v for k,v in spec.items() if k not in ('min','max','customWidth')}
            attrs.setdefault('style',sheet['default_style'])
            attrs.setdefault('width',sheet['default_width'])
            for i in range(int(spec['min']),int(spec['max'])+1):result[i]=attrs
        return result
    column_a,column_b=columns(old),columns(new)
    if column_a!=column_b:
        i=next(i for i in column_a.keys()|column_b.keys() if column_a.get(i)!=column_b.get(i))
        print('column difference',name,i,{k:v for k,v in column_a.get(i,{}).items() if column_b.get(i,{}).get(k)!=v}, {k:v for k,v in column_b.get(i,{}).items() if column_a.get(i,{}).get(k)!=v})
    for k in ('merges','validations'):
        if old[k]!=new[k]:print('layout difference',name,k,old[k],new[k])
    report['other_layout_preserved'] &= column_a==column_b and all(old[k]==new[k] for k in ('merges','validations'))
report['passed']=not report['unexpected_cell_changes'] and not report['unexpected_row_changes'] and report['other_layout_preserved']
(OUT/'scope-report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({k:v if not isinstance(v,dict) else {n:len(a) for n,a in v.items()} for k,v in report.items()},ensure_ascii=False,indent=2))
assert report['passed']
