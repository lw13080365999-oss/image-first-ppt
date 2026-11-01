#!/usr/bin/env python3
"""Bind explicitly authored OMML to named formula placeholders. No LaTeX guessing."""
import argparse, copy, json, pathlib, posixpath, zipfile, xml.etree.ElementTree as ET
from urllib.parse import unquote, urlsplit

NS={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p':'http://schemas.openxmlformats.org/presentationml/2006/main',
    'm':'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'a14':'http://schemas.microsoft.com/office/drawing/2010/main',
    'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
for k,v in NS.items():ET.register_namespace(k,v)
def q(k,n):return '{'+NS[k]+'}'+n
def safe(root,rel):
    if not isinstance(rel,str) or pathlib.Path(rel).is_absolute():raise ValueError('Expected project-relative path')
    p=(root/rel).resolve()
    if not p.is_relative_to(root):raise ValueError('Path escapes project')
    return p
def ordered_slide_parts(archive, expected_count):
    """Resolve presentation order, including package-absolute relationship targets."""
    presentation=ET.fromstring(archive.read('ppt/presentation.xml'))
    rels=ET.fromstring(archive.read('ppt/_rels/presentation.xml.rels'))
    relationships={r.get('Id'):r for r in rels}
    rid='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
    pages=presentation.findall('p:sldIdLst/p:sldId',NS)
    if len(pages)!=expected_count:raise ValueError('Presentation page count differs from project')
    parts=[]
    for page in pages:
        relationship=relationships.get(page.get(rid))
        if relationship is None or relationship.get('TargetMode','').lower()=='external':
            raise ValueError('Invalid slide relationship')
        target=relationship.get('Target','')
        parsed=urlsplit(target)
        if not target or parsed.scheme or parsed.netloc or '\\' in target:
            raise ValueError('Invalid slide relationship target')
        target=unquote(parsed.path)
        part=posixpath.normpath(target.lstrip('/') if target.startswith('/') else posixpath.join('ppt',target))
        if not part.startswith('ppt/slides/') or part not in archive.namelist():
            raise ValueError('Invalid slide relationship target: '+target)
        parts.append(part)
    return parts
def main():
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('input');p.add_argument('output');p.add_argument('--project',required=True)
    a=p.parse_args();root=pathlib.Path(a.project).resolve();out=pathlib.Path(a.output).resolve()
    if out.exists():raise ValueError('Refusing to overwrite output')
    deck=json.loads((root/'deck.json').read_text());updates={};count=0
    with zipfile.ZipFile(a.input) as z:
        parts=ordered_slide_parts(z,len(deck['slides']))
        for i,spec in enumerate(deck['slides'],1):
            mapping=json.loads(safe(root,'slides/'+spec['id']+'/objects.json').read_text())
            if mapping.get('slide_id')!=spec['id']:raise ValueError('Object map slide_id differs from project')
            formulas=[o for o in mapping['objects'] if o['kind']=='formula']
            if not formulas:continue
            part=parts[i-1]
            if not part or not part.startswith('ppt/slides/'):raise ValueError('Invalid slide relationship')
            tree=ET.fromstring(z.read(part))
            for obj in formulas:
                name='ifppt:%s:%s'%(spec['id'],obj['id'])
                matches=[s for s in tree.findall('.//p:sp',NS) if (s.find('p:nvSpPr/p:cNvPr',NS) is not None and s.find('p:nvSpPr/p:cNvPr',NS).get('name')==name)]
                if len(matches)!=1:raise ValueError('Expected one named formula placeholder: '+name)
                raw=safe(root,obj['omml_path']).read_bytes()
                if b'<!DOCTYPE' in raw.upper() or b'<!ENTITY' in raw.upper():raise ValueError('DTD/entities not accepted')
                math=ET.fromstring(raw)
                if math.tag not in (q('m','oMath'),q('m','oMathPara')):raise ValueError('Root must be m:oMath or m:oMathPara')
                allowed={NS['m'],NS['w']}
                if any(not e.tag.startswith('{') or e.tag[1:].split('}')[0] not in allowed for e in math.iter()):raise ValueError('Unexpected namespace in OMML')
                if not ''.join(e.text or '' for e in math.iter(q('m','t'))).strip():raise ValueError('Empty equation')
                math_font=obj.get('font_family') or deck.get('style',{}).get('math_font_family') or 'STIX Two Math'
                for run in math.iter(q('m','r')):
                    props=run.find(q('w','rPr'))
                    if props is None:
                        props=ET.Element(q('w','rPr'))
                        # OMML CT_R orders m:rPr before w:rPr, then run content.
                        math_props=run.find(q('m','rPr'))
                        index=list(run).index(math_props)+1 if math_props is not None else 0
                        run.insert(index,props)
                    if props.find(q('w','rFonts')) is None:ET.SubElement(props,q('w','rFonts'),{q('w','ascii'):math_font,q('w','hAnsi'):math_font})
                    if props.find(q('w','sz')) is None:ET.SubElement(props,q('w','sz'),{q('w','val'):str(round(obj.get('font_size_pt',24)*2))})
                tx=matches[0].find('p:txBody',NS)
                if tx is None:raise ValueError('Formula placeholder has no text body')
                original_para=tx.find('a:p',NS)
                paragraph_props=original_para.find('a:pPr',NS) if original_para is not None else None
                for para in list(tx):
                    if para.tag==q('a','p'):tx.remove(para)
                para=ET.SubElement(tx,q('a','p'))
                if paragraph_props is not None:para.append(copy.deepcopy(paragraph_props))
                wrapper=ET.SubElement(para,q('a14','m'));wrapper.append(copy.deepcopy(math))
                ET.SubElement(para,q('a','endParaRPr'),{'lang':'en-US'})
                count+=1
            updates[part]=ET.tostring(tree,encoding='utf-8',xml_declaration=True)
        out.parent.mkdir(parents=True,exist_ok=True)
        with zipfile.ZipFile(out,'w',compression=zipfile.ZIP_DEFLATED) as dest:
            for info in z.infolist():dest.writestr(info,updates.get(info.filename,z.read(info.filename)))
    print(json.dumps({'output':str(out),'native_formulas_embedded':count,'semantic_and_visual_review_required':True}))
if __name__=='__main__':
    try:main()
    except (ValueError,OSError,KeyError,zipfile.BadZipFile,ET.ParseError) as e:raise SystemExit(str(e))
