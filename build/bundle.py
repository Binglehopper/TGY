import base64
tmpl = open('template.html','rb').read()
parse_src = open('parse.py').read(); build_src = open('build.py').read()
parse_body = parse_src.split('if __name__ ==')[0].replace('def parse(path):','def parse_workbook(path):')
build_body = build_src.split('if __name__ ==')[0].replace('import parse\n','')
build_body = build_body.replace('''    here = os.path.dirname(os.path.abspath(__file__))
    with open(os.path.join(here, "template.html")) as f:
        html = f.read()''','''    html = base64.b64decode(TEMPLATE_B64).decode()''')
build_body = build_body.replace('''    with open(os.path.join(here, "template.html")) as f:
        tmpl = f.read()''','''    tmpl = base64.b64decode(TEMPLATE_B64).decode()''')
build_body = build_body.replace('payload = parse.parse(workbook)','payload = parse_workbook(workbook)')
open('taglyz_builder.py','w').write(f'''#!/usr/bin/env python3
"""
TAGLYZ portfolio report - self-contained builder.

  pip install openpyxl cryptography
  python3 taglyz_builder.py <workbook.xlsx> "<source file name>" "<passphrase>" <outdir>

Writes <outdir>/index.html: the full report with its data encrypted under the
passphrase. Upload that file to github.com/Binglehopper/TGY to publish it.

Generated file - bundles parse.py, build.py and template.html so the whole
pipeline travels as one artifact. Regenerate with bundle.py after editing those.
"""
{parse_body}
{build_body}

TEMPLATE_B64 = "{base64.b64encode(tmpl).decode()}"

if __name__ == "__main__":
    main()
''')
print('bundled', len(open('taglyz_builder.py').read()), 'bytes')
