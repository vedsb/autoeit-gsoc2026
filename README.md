# AutoEIT — Automated EIT Scoring (GSoC 2026 Test II)

## Task
Implement a reproducible script applying the Ortega (2000) meaning-based rubric to Spanish EIT
sentence transcriptions, outputting sentence-level scores (0–4) for each learner utterance.

## Approach

A **hybrid NLP + rule-based pipeline** using spaCy (Spanish) that:

1. **Preprocesses** each transcription → extracts the best final response (strips `[pause]`, `[gibberish]`, `xxx`, false starts like `dis- disminuido`)
2. **Score 0**: no real words produced (silence / entirely garbled)
3. **Score 4**: normalized exact match OR synonym-only difference (`muy`, `y`/`pero`)
4. **NLP via spaCy `es_core_news_sm`**: lemmatise both stimulus and transcription, compute *idea-unit overlap ratio* (matched content words / total stimulus content words)
5. **Score 1**: ≤1 real word, only function words, or ≤1 content word matched
6. **Score 2**: idea-unit ratio 0.40–0.75 (partial meaning)
7. **Score 3**: idea-unit ratio ≥ 0.75 (full meaning preserved; per protocol: "when in doubt → score 2")

## Files

| File | Purpose |
|------|---------|
| `eit_scorer.py` | Standalone scoring script |
| `AutoEIT_Scoring.ipynb` | Jupyter notebook — full methodology, results, charts |
| `generate_outputs.py` | Generates all outputs (Excel + chart) in one command |
| `AutoEIT_Scored_Output.xlsx` | **Primary output** — colour-coded scorated xlsx |
| `score_distributions.png` | Score distribution bar charts per participant |

Input data (not included — place in this directory):
```
AutoEIT Test Files/
  Sample Audio Files and Transcriptions/
    AutoEIT Sample Transcriptions for Scoring.xlsx
```

## Setup & Run

```bash
# 1. Install dependencies
pip install openpyxl spacy
python -m spacy download es_core_news_sm

# 2. Run the scorer
python eit_scorer.py

# 3. Or run everything (scorer + chart)
python generate_outputs.py

# 4. Open the Jupyter notebook
jupyter notebook AutoEIT_Scoring.ipynb
```

## Results Summary

| Participant | Mean Score | Notes |
|-------------|-----------|-------|
| 38001-1A    | ~2.48     | Stronger L2 profile — mostly scores 3–4 |
| 38002-2A    | ~1.83     | Weaker — many garbled/minimal responses |
| 38004-2A    | ~1.97     | Moderate — partial repetitions |
| 38006-2A    | ~1.80     | Weaker — frequent unintelligible segments |

Scores validated against 8 rubric examples from the Ortega (2000) protocol — all 8 matched expected scores.

## Evaluation Strategy
- Validated against rubric-provided examples (8/8 correct)
- Future: compare against human rater scores using Cohen's kappa
- Future: improve with FastText ES word embeddings for semantic similarity (detects meaning opposition, key for score 2 boundary)

## Dependencies
- Python 3.8+
- `openpyxl` — Excel I/O
- `spacy` + `es_core_news_sm` — Spanish NLP (lemmatisation, POS tagging)
- `matplotlib`, `pandas` — charts and data analysis (notebook only)
