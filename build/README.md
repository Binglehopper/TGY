# Generator

`taglyz_builder.py` is the whole pipeline in one file — it bundles `parse.py`,
`build.py` and `template.html`. The refresh jobs fetch *this* file from the repo's
raw URL, so keep it here and keep the repo public.

**This is the file that decides what the dashboard can do.** A refresh job builds
`index.html` by running whatever version of this file is in the repo. If it is out
of date, a refresh will quietly roll the dashboard back to an older feature set,
even though the numbers are current. After any change to the report itself, upload
this folder before running a refresh.

```bash
pip install openpyxl cryptography
python3 taglyz_builder.py <workbook.xlsx> "<source file name>" "<passphrase>" <outdir>
```

The files beside it are the sources it was built from. Edit those, then run
`python3 bundle.py` to regenerate `taglyz_builder.py`.
