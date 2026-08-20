# Generator

`taglyz_builder.py` is the whole pipeline in one file — it bundles `parse.py`,
`build.py` and `template.html`. The monthly refresh job fetches *this* file from
the repo's raw URL, so keep it here and keep the repo public.

```bash
pip install openpyxl cryptography
python3 taglyz_builder.py <workbook.xlsx> "<source file name>" "<passphrase>" <outdir>
```

The files beside it are the sources it was built from. Edit those, then run
`python3 bundle.py` to regenerate `taglyz_builder.py`.
