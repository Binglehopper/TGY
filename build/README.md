# Generator

`taglyz_builder.py` is the whole pipeline in one file — it bundles `parse.py`,
`build.py` and `template.html`. The monthly refresh job fetches *this* file from
the repo's raw URL, so keep it here and keep the repo public.

```bash
pip install openpyxl cryptography
python3 taglyz_builder.py <workbook.xlsx> "<source file name>" "<passphrase>" <outdir>
```

The three source files beside it are what it was generated from; edit those and
regenerate if you need to change the report.
