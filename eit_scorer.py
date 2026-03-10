"""
eit_scorer.py -- AutoEIT Automated Scoring Script
=================================================
GSoC 2026 Test II: Automated Evaluation of Transcribed EIT Data

Author: AutoEIT GSoC Applicant
Date  : March 2026

Description
-----------
This script applies the Ortega (2000) meaning-based rubric to Spanish EIT
transcriptions, assigning scores 0–4 to each learner utterance by comparing
it to the target stimulus sentence.

Scoring Rubric (Ortega, 2000)
------------------------------
  0 -- No response, silence, or entirely unintelligible (garbled/XXX)
  1 -- Minimal repetition then abandoned:
      - Only 1 word repeated, or only function words, OR
      - 1–2 content words out of order with extraneous words not in stimulus
  2 -- ~Half (or fewer) idea units preserved; meaning lost, incomplete, or
      the utterance doesn't form a self-standing, related sentence
  3 -- Full meaning preserved (may be ungrammatical); synonymous substitutions
      of 'muy', 'y'/'pero' are acceptable
  4 -- Exact repetition: form and meaning match stimulus exactly

Protocol Rules Applied
-----------------------
  - Score the *best final response* (after false starts / self-corrections)
  - 'mejor' -> acceptable; 'muy' omission/addition -> not a meaning change (score 3)
  - 'y'/'pero' substitution acceptable for score 3
  - When in doubt between 2 and 3, score 2

Pipeline
--------
  1. Load Excel sheets (participants × 30 sentences)
  2. Preprocess: extract best-final-response, strip disfluency markers,
     normalize (lowercase, unicode-normalize, remove punctuation)
  3. Score each (stimulus, transcription) pair through a decision tree:
       a. Score 0: empty / all-gibberish
       b. Score 4: exact match post-normalization
       c. NLP: lemmatize both, find content-word overlap -> idea-unit ratio
       d. Calibrate ratio with sequence-order bonus
       e. Apply special-case rules (muy, y/pero, self-corrections)
  4. Write scored data to AutoEIT_Scored_Output.xlsx
  5. Print a summary table

Usage
-----
  python eit_scorer.py
  (Run from the project root directory)

Dependencies
------------
  pip install openpyxl spacy
  python -m spacy download es_core_news_sm
"""

import re
import unicodedata
import copy
import os
import shutil
from typing import Tuple

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
import spacy

# ──────────────────────────────────────────────────────────────────────────────
# Configuration
# ──────────────────────────────────────────────────────────────────────────────

INPUT_DIR = os.path.join(
    os.path.dirname(__file__),
    "AutoEIT Test Files",
    "Sample Audio Files and Transcriptions",
)
INPUT_XLSX = os.path.join(INPUT_DIR, "AutoEIT Sample Transcriptions for Scoring.xlsx")
OUTPUT_XLSX = os.path.join(os.path.dirname(__file__), "AutoEIT_Scored_Output.xlsx")

# Columns in each participant sheet (1-indexed)
COL_SENTENCE = 1      # A -- sentence number
COL_STIMULUS = 2      # B -- target sentence (with word count in parens)
COL_TRANSCRIPT = 3    # C -- learner transcription
COL_SCORE = 4         # D -- score to fill

# Score colours for Excel output (light fills)
SCORE_COLOURS = {
    0: "FFD7D7",   # light red
    1: "FFE8C8",   # light orange
    2: "FFFACD",   # light yellow
    3: "D5F5D5",   # light green
    4: "C8E6FA",   # light blue
}

# Spanish function word lists (for Score 1 heuristic)
FUNCTION_WORDS: set = {
    "a", "al", "ante", "bajo", "con", "contra", "de", "del", "desde",
    "durante", "en", "entre", "hacia", "hasta", "mediante", "para", "por",
    "según", "sin", "sobre", "tras",
    "el", "la", "los", "las", "un", "una", "unos", "unas",
    "y", "e", "o", "u", "pero", "sino", "ni", "mas", "que",
    "se", "me", "te", "le", "lo", "les", "nos", "os",
    "que", "quien", "quienes", "cual", "cuales", "donde", "cuando", "como",
    "no", "si", "ya", "aún", "también", "tampoco",
    "yo", "tú", "él", "ella", "nosotros", "vosotros", "ellos", "ellas",
    "usted", "ustedes",
    "este", "esta", "estos", "estas", "ese", "esa", "esos", "esas",
    "aquel", "aquella", "aquellos", "aquellas",
    "mi", "tu", "su", "nuestro", "vuestro", "mis", "tus", "sus",
    "mío", "mía", "tuyo", "tuya", "suyo", "suya",
    "es", "son", "está", "están",      # very common copulas treated as functional
}

# Garble markers in transcriptions
GARBLE_RE = re.compile(
    r"\b(xxx|xx|x\b|gibberish|pause|no response|inaudible|unintelligible)\b"
    r"|\[.*?\]"     # bracket annotations
    r"|\.\.\.",
    re.IGNORECASE,
)

# False-start detector: repeated partial -> full word  e.g. "dis- disminuido"
FALSE_START_RE = re.compile(r"\b(\w+)-\s+\1", re.IGNORECASE)

# Accent-tolerant synonyms mapping (normalised -> normalised)
SYNONYMS = {
    "pero": "y",
    "e": "y",
    "u": "o",
    "muy": "",            # 'muy' presence/absence is irrelevant
    "mucho": "muy",
}


# ──────────────────────────────────────────────────────────────────────────────
# NLP setup
# ──────────────────────────────────────────────────────────────────────────────

print("Loading spaCy Spanish model...")
try:
    NLP = spacy.load("es_core_news_sm")
except OSError:
    raise SystemExit(
        "Spanish model not found. Run:  python -m spacy download es_core_news_sm"
    )


# ──────────────────────────────────────────────────────────────────────────────
# Text preprocessing helpers
# ──────────────────────────────────────────────────────────────────────────────

def normalize(text: str) -> str:
    """Lowercase, strip accents, remove punctuation, collapse whitespace."""
    text = text.lower()
    # Unicode-normalize: decompose accented chars, then remove combining marks
    nfkd = unicodedata.normalize("NFKD", text)
    text = "".join(c for c in nfkd if not unicodedata.combining(c))
    # Remove punctuation except hyphens (handled separately)
    text = re.sub(r"[^\w\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def strip_word_count(stimulus: str) -> str:
    """Remove trailing word count annotation like ' (7)' or ' (17)'."""
    return re.sub(r"\s*\(\d+\)\s*$", "", stimulus).strip()


def extract_best_final_response(transcription: str) -> str:
    """
    Apply protocol rule: score the best final response.

    Steps:
    1. Remove false-start fragments (e.g. 'dis-' before 'disminuido').
    2. When a speaker self-corrects (uses '...' or a dash then continues),
       prefer the corrected form.
    3. Strip transcription annotation markers.
    """
    if not transcription or not transcription.strip():
        return ""

    t = transcription

    # Remove square-bracket annotations like [pause], [gibberish], [cough]
    t = re.sub(r"\[.*?\]", " ", t)

    # Handle explicit false starts: "dis- disminuido" -> "disminuido"
    # Pattern: partial word ending in '-' followed by space + fuller word
    t = re.sub(r"\b\w*-\s+", " ", t)

    # Handle em-dash self-corrections: take text after the last em-dash
    # e.g. "Quiero comerme la- el huevo" -> keep 'el huevo' part
    if "- " in t:
        # If there's a dash-space, strip up to and including partial segment
        parts = re.split(r"\w+-\s+", t)
        t = parts[-1] if len(parts) > 1 else t

    # Strip garble markers
    t = GARBLE_RE.sub(" ", t)

    # Collapse whitespace
    t = re.sub(r"\s+", " ", t).strip()

    return t


def get_content_words(text: str) -> list:
    """
    Return list of lowercased, accent-stripped lemmas of content words
    (NOUN, VERB, ADJ, ADV) from the text using spaCy.
    """
    doc = NLP(text)
    words = []
    for token in doc:
        if (
            not token.is_space
            and not token.is_punct
            and token.pos_ in {"NOUN", "VERB", "ADJ", "ADV", "PROPN"}
            and token.lemma_.lower() not in FUNCTION_WORDS
        ):
            words.append(normalize(token.lemma_))
    return words


def get_all_words(text: str) -> list:
    """Return all normalized non-punctuation tokens."""
    return [normalize(t) for t in text.split() if normalize(t)]


# ──────────────────────────────────────────────────────────────────────────────
# Scoring engine
# ──────────────────────────────────────────────────────────────────────────────

def compute_idea_unit_overlap(
    stim_content: list, trans_content: list
) -> Tuple[float, int, int]:
    """
    Compute the fraction of stimulus content words (idea units) reproduced
    in the transcription.

    Returns:
        ratio       -- matched / total stimulus content words
        matched     -- count of matched idea units
        total       -- total stimulus idea units
    """
    if not stim_content:
        return 1.0, 0, 0

    stim_multiset = {}
    for w in stim_content:
        stim_multiset[w] = stim_multiset.get(w, 0) + 1

    trans_set = set(trans_content)

    matched = 0
    for w, count in stim_multiset.items():
        if w in trans_set:
            matched += min(count, trans_content.count(w))

    ratio = matched / len(stim_content)
    return ratio, matched, len(stim_content)


def word_order_preserved(stim_content: list, trans_content: list) -> bool:
    """
    Check whether the matched content words appear in the same relative order
    in both the stimulus and the transcription (subsequence check).
    """
    if not stim_content or not trans_content:
        return False

    trans_idx = 0
    for word in stim_content:
        while trans_idx < len(trans_content) and trans_content[trans_idx] != word:
            trans_idx += 1
        if trans_idx >= len(trans_content):
            return False
        trans_idx += 1
    return True


def check_synonymous_substitution_only(
    stim_norm: str, trans_norm: str, stim_content: list, trans_content: list
) -> bool:
    """
    Return True if the ONLY differences between stimulus and transcription are:
      - presence/absence of 'muy'
      - substitution of 'y'/'e' for 'pero'/'mas'
    These are rubric-approved synonyms that do not affect score-3 eligibility.
    """
    def apply_synonyms(words: list) -> list:
        result = []
        for w in words:
            mapped = SYNONYMS.get(w, w)
            if mapped:  # drop empty (e.g. 'muy' -> '')
                result.append(mapped)
        return result

    stim_mapped = apply_synonyms(get_all_words(stim_norm))
    trans_mapped = apply_synonyms(get_all_words(trans_norm))
    return stim_mapped == trans_mapped


def is_all_garbled(cleaned: str) -> bool:
    """
    Return True if after cleaning the transcription has no real Spanish words.
    (e.g. only 'xxx', empty, or isolated single letters)
    """
    if not cleaned.strip():
        return True
    words = cleaned.split()
    real_words = [w for w in words if len(w) > 1]
    return len(real_words) == 0


def score_sentence(
    stimulus_raw: str, transcription_raw: str, sentence_num: int = 0
) -> Tuple[int, dict]:
    """
    Main scoring function.

    Parameters
    ----------
    stimulus_raw       : target sentence (may include word-count annotation)
    transcription_raw  : learner's transcription (may include disfluency markers)
    sentence_num       : sentence index (for debugging)

    Returns
    -------
    score   : integer 0–4
    details : dict with intermediate values for transparency/logging
    """
    details = {
        "stimulus_raw": stimulus_raw,
        "transcription_raw": transcription_raw,
    }

    # ── Preprocessing ─────────────────────────────────────────────────────────
    stim_clean = strip_word_count(stimulus_raw or "")
    trans_clean = extract_best_final_response(transcription_raw or "")

    stim_norm = normalize(stim_clean)
    trans_norm = normalize(trans_clean)

    details["stimulus_clean"] = stim_clean
    details["transcription_cleaned"] = trans_clean
    details["stimulus_normalized"] = stim_norm
    details["transcription_normalized"] = trans_norm

    # ── Score 0: empty or entirely unintelligible ─────────────────────────────
    if is_all_garbled(trans_norm):
        details["decision"] = "Score 0: no real words produced"
        return 0, details

    trans_words = get_all_words(trans_norm)
    stim_words = get_all_words(stim_norm)

    details["trans_word_count"] = len(trans_words)
    details["stim_word_count"] = len(stim_words)

    # ── Score 4: exact match ───────────────────────────────────────────────────
    if stim_norm == trans_norm:
        details["decision"] = "Score 4: exact match"
        return 4, details

    # ── Special case: synonymous substitution only (-> still score 4 boundary)
    if check_synonymous_substitution_only(stim_norm, trans_norm, [], []):
        details["decision"] = "Score 4: synonymous substitution only (muy/y/pero)"
        return 4, details

    # ── NLP: content word extraction + idea-unit ratio ────────────────────────
    stim_content = get_content_words(stim_clean)
    trans_content = get_content_words(trans_clean) if trans_clean else []

    ratio, matched, total = compute_idea_unit_overlap(stim_content, trans_content)
    order_ok = word_order_preserved(stim_content, trans_content)

    details["stim_content_words"] = stim_content
    details["trans_content_words"] = trans_content
    details["idea_unit_ratio"] = round(ratio, 3)
    details["matched_content_words"] = matched
    details["total_content_words"] = total
    details["word_order_preserved"] = order_ok

    # ── Score 1: minimal / abandoned ──────────────────────────────────────────
    # Triggers when: very few words OR only function words OR idea_unit_ratio tiny
    only_function = all(w in FUNCTION_WORDS for w in trans_words)

    if len(trans_words) <= 1 or only_function:
        details["decision"] = "Score 1: minimal response (≤1 word or only function words)"
        return 1, details

    if matched <= 1 and total > 2:
        details["decision"] = "Score 1: at most 1 content word matched out of >2 in stimulus"
        return 1, details

    # ── Score 3 / 2 boundary based on idea-unit ratio + order ─────────────────
    #
    # Rubric guidance:
    #   - Score 3: full meaning preserved (ungrammatical OK), may-muy/y-pero OK
    #   - Score 2: more than half idea units missing OR meaning changed/opposed
    #   - When in doubt -> score 2
    #
    # We use threshold 0.75 for score 3 (generous enough for minor function-word
    # errors while requiring most content to be present).
    # Order bonus: if word order is preserved add +0.10 to ratio for scoring decision only.

    effective_ratio = ratio + (0.10 if order_ok else 0.0)
    details["effective_ratio"] = round(effective_ratio, 3)

    # Check if only approved synonyms differ (muy, y/pero) -- counts as score 3
    # even if normalized forms differ slightly
    stim_mapped_words = [SYNONYMS.get(w, w) for w in get_all_words(stim_norm) if SYNONYMS.get(w, w)]
    trans_mapped_words = [SYNONYMS.get(w, w) for w in get_all_words(trans_norm) if SYNONYMS.get(w, w)]
    synonym_adjusted_match = (stim_mapped_words == trans_mapped_words)

    if synonym_adjusted_match:
        details["decision"] = "Score 3: synonym-adjusted exact match (muy/y/pero substitution)"
        return 3, details

    if effective_ratio >= 0.75:
        details["decision"] = f"Score 3: idea-unit ratio {ratio:.2f} (effective {effective_ratio:.2f}) ≥ 0.75, meaning preserved"
        return 3, details
    elif effective_ratio >= 0.40:
        details["decision"] = f"Score 2: idea-unit ratio {ratio:.2f} (effective {effective_ratio:.2f}) in [0.40, 0.75)"
        return 2, details
    else:
        details["decision"] = f"Score 1: idea-unit ratio {ratio:.2f} (effective {effective_ratio:.2f}) < 0.40"
        return 1, details


# ──────────────────────────────────────────────────────────────────────────────
# Excel I/O
# ──────────────────────────────────────────────────────────────────────────────

def load_and_score_workbook(input_path: str) -> Tuple[openpyxl.Workbook, dict]:
    """
    Load the transcription workbook, score every sentence, and return
    the annotated workbook along with a results dict.
    """
    wb = openpyxl.load_workbook(input_path)
    results = {}

    data_sheets = [s for s in wb.sheetnames if s != "Info"]

    for sheet_name in data_sheets:
        ws = wb[sheet_name]
        participant_results = []

        for row in ws.iter_rows(min_row=2):
            # Skip empty/header rows
            if row[COL_SENTENCE - 1].value is None:
                continue
            if not isinstance(row[COL_SENTENCE - 1].value, (int, float)):
                continue

            sentence_num = int(row[COL_SENTENCE - 1].value)
            stimulus = row[COL_STIMULUS - 1].value or ""
            transcription = row[COL_TRANSCRIPT - 1].value or ""
            score_cell = row[COL_SCORE - 1]

            score, details = score_sentence(stimulus, transcription, sentence_num)

            # Write score to cell
            score_cell.value = score

            # Apply colour fill
            fill_colour = SCORE_COLOURS.get(score, "FFFFFF")
            score_cell.fill = PatternFill(
                start_color=fill_colour, end_color=fill_colour, fill_type="solid"
            )
            score_cell.font = Font(bold=True)
            score_cell.alignment = Alignment(horizontal="center")

            participant_results.append(
                {
                    "sentence": sentence_num,
                    "stimulus": stimulus,
                    "transcription": transcription,
                    "score": score,
                    "decision": details.get("decision", ""),
                    "idea_unit_ratio": details.get("idea_unit_ratio", None),
                    "matched": details.get("matched_content_words", None),
                    "total": details.get("total_content_words", None),
                }
            )

        results[sheet_name] = participant_results

        # Colour the Score column header
        header_cell = ws.cell(row=1, column=COL_SCORE)
        header_cell.fill = PatternFill(
            start_color="E0E0E0", end_color="E0E0E0", fill_type="solid"
        )
        header_cell.font = Font(bold=True)

    return wb, results


def save_workbook(wb: openpyxl.Workbook, output_path: str) -> None:
    """Save workbook, copying any unmodified sheets as-is."""
    wb.save(output_path)
    print(f"\n[OK]  Scored workbook saved -> {output_path}")


# ──────────────────────────────────────────────────────────────────────────────
# Summary reporting
# ──────────────────────────────────────────────────────────────────────────────

def print_summary(results: dict) -> None:
    """Print a formatted console summary of all scores."""
    print("\n" + "=" * 80)
    print("  AutoEIT AUTOMATED SCORING SUMMARY")
    print("  Rubric: Ortega (2000) meaning-based, 0–4 scale")
    print("=" * 80)

    all_scores = []

    for participant, rows in results.items():
        scores = [r["score"] for r in rows]
        all_scores.extend(scores)
        total = sum(scores)
        mean = total / len(scores) if scores else 0

        print(f"\n{'─' * 60}")
        print(f"  Participant: {participant}   |  Total: {total}/120  |  Mean: {mean:.2f}")
        print(f"{'─' * 60}")
        print(f"  {'#':>3}  {'Score':>5}  {'Ratio':>6}  Stimulus (truncated)")
        print(f"  {'─'*3}  {'─'*5}  {'─'*6}  {'─'*40}")

        for r in rows:
            stim_short = strip_word_count(r["stimulus"])[:42]
            ratio_str = f"{r['idea_unit_ratio']:.2f}" if r["idea_unit_ratio"] is not None else "  -  "
            score_str = str(r["score"])
            print(f"  {r['sentence']:>3}  {score_str:>5}  {ratio_str:>6}  {stim_short}")

        # Score distribution
        from collections import Counter
        dist = Counter(scores)
        print(f"\n  Score distribution: ", end="")
        for s in range(5):
            print(f"[{s}]={dist.get(s, 0)}", end="  ")
        print()

    if all_scores:
        print(f"\n{'=' * 80}")
        print(f"  OVERALL  |  N={len(all_scores)}  |  Mean={sum(all_scores)/len(all_scores):.2f}  "
              f"|  Total={sum(all_scores)}/{len(all_scores)*4}")
        print("=" * 80)


# ──────────────────────────────────────────────────────────────────────────────
# Entrypoint
# ──────────────────────────────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("  AutoEIT Automated EIT Scorer")
    print("  GSoC 2026 Test II -- Meaning-Based Rubric (Ortega 2000)")
    print("=" * 60)
    print(f"\nInput  : {INPUT_XLSX}")
    print(f"Output : {OUTPUT_XLSX}\n")

    if not os.path.exists(INPUT_XLSX):
        raise FileNotFoundError(f"Input file not found: {INPUT_XLSX}")

    wb, results = load_and_score_workbook(INPUT_XLSX)
    save_workbook(wb, OUTPUT_XLSX)
    print_summary(results)

    print("\nDone! Open AutoEIT_Scored_Output.xlsx to see colour-coded scores.")


if __name__ == "__main__":
    main()
