#!/usr/bin/env python3
"""Render an exported PPTX and prepare hash-bound visual review evidence."""
import argparse, hashlib, json, os, pathlib, re, shutil, subprocess, tempfile

def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def run(args,env=None):
    r=subprocess.run(args,capture_output=True,text=True,timeout=150,env=env)
    if r.returncode:raise ValueError('Render command failed: '+r.stderr[-2000:])
    return r.stdout.strip()
def sheet(paths,out,labels,columns=3,width=480):
    from PIL import Image,ImageOps,ImageDraw
    height=round(width*9/16);rows=(len(paths)+columns-1)//columns
    canvas=Image.new('RGB',(columns*width,rows*(height+30)), '#E8E8E8');draw=ImageDraw.Draw(canvas)
    for i,(p,label) in enumerate(zip(paths,labels)):
        im=Image.open(p).convert('RGB');im=ImageOps.contain(im,(width-12,height-12))
        x=(i%columns)*width;y=(i//columns)*(height+30)
        canvas.paste(im,(x+(width-im.width)//2,y+(height-im.height)//2))
        draw.text((x+8,y+height+5),label,fill='#202020')
    canvas.save(out)
def main():
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('pptx');p.add_argument('--project',required=True);p.add_argument('--out',required=True)
    p.add_argument('--soffice');p.add_argument('--pdftoppm');p.add_argument('--size',type=int,default=1800)
    p.add_argument('--font-dir',action='append',default=[],help='Directory of existing fonts loaded into the temporary Office profile; repeat as needed (nonrecursive)')
    p.add_argument('--font-file',action='append',default=[],help='Existing TTF/OTF/TTC/OTC font loaded into the temporary Office profile; repeat as needed')
    a=p.parse_args();pptx=pathlib.Path(a.pptx).resolve();root=pathlib.Path(a.project).resolve();out=pathlib.Path(a.out).resolve()
    if not pptx.is_file():raise ValueError('PPTX missing')
    if out.exists() and any(out.iterdir()):raise ValueError('Use a new empty review directory for this revision')
    bins=pathlib.Path(os.environ['PPT_BIN_DIR']) if os.environ.get('PPT_BIN_DIR') else None
    office=a.soffice or (str(bins/'soffice') if bins and (bins/'soffice').is_file() else shutil.which('soffice'))
    if not office:
        mac=pathlib.Path('/Applications/LibreOffice.app/Contents/MacOS/soffice')
        if mac.is_file():office=str(mac)
    raster=a.pdftoppm or (str(bins/'pdftoppm') if bins and (bins/'pdftoppm').is_file() else shutil.which('pdftoppm'))
    if not office or not raster:raise ValueError('Need LibreOffice soffice and Poppler pdftoppm; rendering has not been performed')
    font_dirs=[str(pathlib.Path(d).resolve()) for d in a.font_dir]
    if any(not pathlib.Path(d).is_dir() for d in font_dirs):raise ValueError('Font directory does not exist')
    extensions={'.ttf','.otf','.ttc','.otc'}
    font_files=[pathlib.Path(f).resolve() for f in a.font_file]
    for directory in font_dirs:
        font_files.extend(f.resolve() for f in pathlib.Path(directory).iterdir() if f.is_file() and f.suffix.lower() in extensions)
    font_files=sorted(set(font_files))
    if any(not f.is_file() or f.suffix.lower() not in extensions for f in font_files):raise ValueError('Expected existing TTF/OTF/TTC/OTC font file')
    if font_dirs and not font_files:raise ValueError('No font files found in supplied directories')
    if a.size<600 or a.size>6000:raise ValueError('--size must be 600..6000')
    try:from PIL import Image
    except ImportError:raise ValueError('Pillow required for contact/comparison sheets; use the available document runtime')
    out.mkdir(parents=True,exist_ok=True)
    deck=json.loads((root/'deck.json').read_text())
    with tempfile.TemporaryDirectory(prefix='image-first-ppt-lo-') as temp:
        profile_path=pathlib.Path(temp)/'profile'
        if font_files:
            # The bundled headless renderer may only see application/profile fonts.
            # A private profile loads existing fonts without changing the runtime.
            profile_fonts=profile_path/'user'/'fonts';profile_fonts.mkdir(parents=True)
            for i,font_file in enumerate(font_files):shutil.copyfile(font_file,profile_fonts/(str(i)+'-'+font_file.name))
        run([office,'-env:UserInstallation='+profile_path.as_uri(),'--headless','--convert-to','pdf','--outdir',str(out),str(pptx)])
    pdf=out/(pptx.stem+'.pdf')
    if not pdf.is_file() or not pdf.stat().st_size:raise ValueError('LibreOffice did not produce PDF')
    run([raster,'-png','-scale-to',str(a.size),str(pdf),str(out/'page')])
    pages=sorted(out.glob('page-*.png'),key=lambda x:int(re.search(r'-(\d+)\.png$',x.name).group(1)))
    if len(pages)!=len(deck['slides']):raise ValueError('Rendered page count differs from deck.json')
    packet={'schema_version':1,'pptx':str(pptx),'pptx_sha256':sha(pptx),'deck_sha256':sha(root/'deck.json'),
            'renderer':'LibreOffice + Poppler','renderer_version':run([office,'--version']),
            'additional_font_dirs':font_dirs,
            'additional_fonts':[{'path':str(f),'sha256':sha(f)} for f in font_files],
            'visual_review_status':'pending','limitations':['Rendered by LibreOffice, not proof of PowerPoint/WPS compatibility','Images must be inspected; preparing this packet is not visual approval'],
            'slides':[]}
    for s,page in zip(deck['slides'],pages):
        row={'slide_id':s['id'],'render_path':str(page),'render_sha256':sha(page),'visual_review_status':'pending'}
        mp=root/'slides'/s['id']/'objects.json'
        if mp.is_file():
            row['objects_sha256']=sha(mp);mapping=json.loads(mp.read_text());design=mapping.get('design',{})
            rel=design.get('path')
            if rel:
                target=(root/rel).resolve()
                if not target.is_relative_to(root) or not target.is_file():raise ValueError('Invalid design path')
                row.update(design_path=str(target),design_sha256=sha(target))
                if design.get('sha256') and row['design_sha256']!=design['sha256']:raise ValueError('Stale design mapping for '+s['id'])
                comparison=out/(s['id']+'-comparison.png')
                try:sheet([target,page],comparison,['DESIGN REFERENCE','FINAL PPTX RENDER'],columns=2,width=900)
                except (OSError,ValueError) as e:raise ValueError('Cannot compare design image for '+s['id']+': '+str(e))
                row['comparison_path']=str(comparison)
        packet['slides'].append(row)
    contact=out/'contact-sheet.png';sheet(pages,contact,[s['id'] for s in deck['slides']])
    packet['contact_sheet']=str(contact)
    review=out/'review.json';review.write_text(json.dumps(packet,ensure_ascii=False,indent=2))
    print(json.dumps({'review_packet':str(review),'contact_sheet':str(contact),'pages':len(pages),'visual_review_status':'pending'},ensure_ascii=False))
if __name__=='__main__':
    try:main()
    except (ValueError,OSError,KeyError,subprocess.TimeoutExpired) as e:raise SystemExit(str(e))
