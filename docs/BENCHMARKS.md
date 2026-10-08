<!-- Generated from benchmark results; edit the generator, not this file. -->

# Benchmarks

Measured on **undoc 0.16.0** (the published package) on 2026-10-08. Scores come from each benchmark's own evaluator; nothing here is re-implemented.

## opendataloader-bench, DOCX edition

There is no public DOCX-to-Markdown dataset with ground truth. This edition typesets the ground truth of [opendataloader-bench](https://github.com/opendataloader-project/opendataloader-bench) (200 real documents, Apache-2.0) as Word documents -- Markdown headings become Heading styles, HTML tables become Word tables with merged cells -- and scores the output with that benchmark's own evaluator. It measures whether the structure a document *declares* survives extraction, not layout inference, so it is an easier test than the PDF benchmark and the scores are not comparable to it.

NID = reading-order text similarity, TEDS = table structure, MHS = heading structure. Each is averaged over the documents it applies to (last column): TEDS only over documents whose ground truth has a table, MHS only over those with headings. A document's overall score is the mean of the metrics that apply to it, and Overall is that score averaged over all documents -- so it is not the mean of the three rows below it.

| Metric | Score | Documents scored |
|---|---|---|
| Overall | **0.993** | 200 |
| NID | 0.991 | 200 |
| TEDS | 0.970 | 42 |
| MHS | 0.996 | 107 |

Reproduce (the corpus builder is [`benchmarks/build_odl_docx.py`](https://github.com/iyulab/undoc/blob/main/benchmarks/build_odl_docx.py)):

```bash
# From a clone of this repository, with Python 3.13 or newer (what the evaluator declares)
git clone https://github.com/opendataloader-project/opendataloader-bench && git -C opendataloader-bench checkout 7af1d8f4d0c09f51ea1a5c6ba5f66e993286d109
pip install undoc==0.16.0 python-docx apted rapidfuzz beautifulsoup4 lxml
python benchmarks/build_odl_docx.py opendataloader-bench odl-docx
python benchmarks/build_odl_docx.py --convert odl-docx
# the evaluator resolves relative directories against its own checkout, so pass absolute ones
python opendataloader-bench/src/evaluator.py --ground-truth-dir "$PWD/odl-docx/ground-truth/markdown" --prediction-root "$PWD/odl-docx/prediction" --engine undoc
# scores: odl-docx/prediction/undoc/evaluation.json, under metrics.score
```

<!-- benchmark-data {"benchmarks": ["odl"], "odl": {"mhs": 0.996, "nid": 0.991, "overall": 0.993, "teds": 0.97}, "version": "0.16.0"} -->
