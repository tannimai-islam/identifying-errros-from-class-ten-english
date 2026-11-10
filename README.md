# Error Analysis of Students' Written English (Bangladesh)
​Educational materials and student corpora at the secondary level (Class 10) often reveal recurring patterns of interference, structural divergence, and pedagogical gaps. This project provides a structured approach to examining these linguistic deviations through systematic error analysis (EA).
​By combining corpus linguistic methods with error categorization, this repository serves as a toolkit for researchers, educators, and NLP developers interested in second language acquisition (SLA) and automated grammatical error correction (GEC) for South Asian English contexts.

## Setup
pip install pandas matplotlib pyspellchecker python-docx

## Run
python error_analysis.py Sample_data.docx

Input: a .docx with essays labelled "Sample 1", "Sample 2", ... (duplicates are removed automatically).

## Outputs (./output)
- flagged_errors.csv : every flagged item with sample number, error type, matched text, full sentence
- summary.csv        : count, % of all errors, errors per 100 words, number of essays affected
- per_essay.csv      : errors per essay
- errors_chart.png   : bar chart of error shares

## Method
1. Load essays, remove duplicate samples, split into sentences.
2. Detect errors: regex rules (agreement, prepositions, articles, tense, plurals, word form,
   apostrophes, punctuation, capitalisation, run-ons) + dictionary spell-check.
3. Count and rank error types (share of errors and errors per 100 words).
4. Manually verify flagged_errors.csv (precision check) before reporting final ratios.

## Limitations
- Rule-based detection finds surface errors; it misses missing articles/words, wrong word choice,
  and meaning-level errors.
- Some flags may be false positives; manual verification is needed.
- Accepted British spellings (centre, favour) are not counted as errors.
