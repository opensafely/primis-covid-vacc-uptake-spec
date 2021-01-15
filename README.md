# PRIMIS Covid-19 Vaccine Uptake Spec

This repo contains:

* `input/v.x.y.docx`, successive versions of "SARS-CoV2 (COVID-19) Vaccine Uptake Reporting Specification" documents from PRIMIS
* `extact_data.py`, a script to extract data from the "Clinical data extraction criteria" tables
* `output/v.x.y.json`, the data from these tables as JSON
* `output/latest.json`, the latest version the above, allowing diffs between versions

When there's a new version of the spec:

* copy it to `input/v.x.y.docx`
* run `python extract_data input/v.x.y.docx output/v.x.y.json`
* run `cp output/v.x.y.json output/latest.json`
* commit and push the changes
