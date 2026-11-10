"""
Error analysis of students' written English (Bangladesh)
--------------------------------------------------------
Usage:
    python error_analysis.py data.docx

Reads essays labelled "Sample 1", "Sample 2", ... from a .docx file,
removes duplicate samples, splits into sentences, detects errors with
rule-based checks + a spell-checker, then reports which error types
have the biggest share.

Outputs (in ./output):
    flagged_errors.csv   - every flagged item (sentence, error type, matched text)  -> CHECK BY HAND
    summary.csv          - count, % of errors, errors per 100 words
    per_essay.csv        - error count per essay
    errors_chart.png     - bar chart
"""
import os
import re
import sys
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from docx import Document
from spellchecker import SpellChecker

OUT = "output"

# ---------------------------------------------------------------- loading
def load_samples(path):
    """Return {sample_number: text}, ignoring exact duplicates."""
    doc = Document(path)
    samples, current, buf = {}, None, []
    seen_texts = set()

    def flush():
        if current is None:
            return
        text = "\n".join(buf).strip()
        if text and text not in seen_texts:
            seen_texts.add(text)
            samples[current] = text

    for p in doc.paragraphs:
        t = p.text.strip()
        m = re.fullmatch(r"Sample\s*(\d+)", t, re.I)
        if m:
            flush()
            current, buf = int(m.group(1)), []
        elif t:
            buf.append(t)
    flush()
    return samples

def normalise(text):
    return (text.replace("\u2019", "'").replace("\u2018", "'")
                .replace("\u201c", '"').replace("\u201d", '"'))

def split_sentences(text):
    text = normalise(text)
    parts = re.split(r"(?<=[.!?])\s+|\n+", text)
    return [p.strip() for p in parts if len(p.strip().split()) >= 3]

# ---------------------------------------------------------------- rules
IRREGULAR_BASE = r"(go|buy|eat|see|come|take|make|write|read|give|get|do|have|play|study|start)"
PAST_MARKERS = r"\b(yesterday|last\s+(?:year|week|month|night)|ago)\b"
SUBJ_PLURAL = r"(?:they|we|you|i|people|students|many\s+\w+)"
SUBJ_SING = r"(?:he|she|it|everyone|everybody|one|someone|nobody|student|a\s+student|every\s+\w+)"

RULES = {
    "Subject-verb agreement": [
        rf"\b(?:he|she|it|everyone|everybody|a\s+student|every\s+student)\s+(?:want|need|have|do|don't|are|were|know|take|go|get|make|give|help|become)\b",
        r"\b(?:they|we|you|students|people)\s+(?:wants|needs|has|does|doesn't|is|was|knows|takes|goes|gets|makes|gives|helps)\b",
        r"\bI\s+(?:is|has|does|goes|wants|needs|am\s+going\s+to\s+\w+s)\b",
        r"\b(?:which|who)\s+(?:help|give|make|need|become)\b(?!\w)",
        r"\bthere\s+is\s+(?:many|few|so many|lots of|several)\b",
        r"\bit\s+are\b|\bthis\s+are\b|\bthese\s+is\b",
        r"\bif\s+(?:destiny|he|she|it|a student)\s+do\b",
    ],
    "Preposition": [
        r"\bdiscuss\s+about\b", r"\bdepends?\s+of\b", r"\bmarried\s+with\b",
        r"\blisten\s+the\b", r"\breturn\s+back\b", r"\badmitted\s+in\b",
        r"\bgood\s+in\b", r"\bgetting\s+admitted\s+into\b",
        r"\bbehind\s+of\b", r"\bin\s+term\s+of\b",
        r"\bany\s+of\s+this\s+\w+\b", r"\bstudy\s+again\s+between\b",
        r"\bcause\s+of\s+to\b", r"\bnecessary\s+for\s+to\b",
    ],
    "Word form / part of speech": [
        r"\bmore\s+\w+er\b", r"\bmost\s+\w+est\b",
        r"\b(?:hardly|fastly|nearly|lately)\b(?=\s+(?:study|so\s+that|that))",
        r"\bstudy\s+hardly\b", r"\bgoes\s+fastly\b",
        r"\b(?:be|are|is|was|were)\s+(?:wake|help|study|try)\b(?!\w)",
        r"\bhelp\s+to\s+helps?\b",
        r"\bmuch\s+(?:students|marks|questions|books)\b",
        r"\b(?:should|must|will|to)\s+emphasis\b",
        r"\ba\s+lots\s+of\b|\blots\s+of\s+a\b|\b(?:a|an)\s+lots\b",
    ],
    "Tense / verb form": [
        rf"{PAST_MARKERS}.*\b{IRREGULAR_BASE}\b",
        r"\bmust\s+be\s+\w+(?<!ed)(?<!en)\s+up\b",
        r"\bwill\s+be\s+help\b",
        r"\bhe\s+started\s+his\b",
        r"\b(?:to|can|will|should|must|may|could|would)\s+(?:gets|goes|makes|takes|does|has)\b",
        r"\bcan\s+\w+ed\b(?<!need)",
    ],
    "Plural / countable noun": [
        r"\b(?:every|each|this|that)\s+(?:students|questions|sentences|subjects|marks)\b",
        r"\b(?:many|few|several|two|three|four|five|these|those)\s+(?:student|question|subject|topic|candidate|mark|person|people's)\b(?!s)",
        r"\b(?:a|an)\s+(?:students|questions|marks|subjects)\b",
        r"\b(?:informations|advices|knowledges|equipments)\b",
    ],
}

SPELL_WHITELIST = {
    "hsc", "ssc", "buet", "du", "ru", "ju", "jnu", "gk", "mcq", "cq", "bangla",
    "bangladesh", "bangladeshi", "udvash", "unmesh", "allah", "varsity", "dhaka",
    "jahangirnagar", "etc", "covid", "bcs", "bengali", "bcs", "ielts", "sat",
    "shahjalal", "cuet", "kuet", "ruet", "duet", "bup", "cse", "eee", "mbbs",
    "smartphones", "smartphone", "doom", "scrolling", "scribble", "scribbled",
    "online", "wifi", "facebook", "youtube", "google", "biology", "syllabus",
    "coaching", "coachings", "admit", "admitted", "memorise", "memorize",
    "tricks", "revise", "revising", "p.m", "a.m", "prep", "mcqs", "sheets",
    "timetable", "ai", "pdf", "lecture", "lectures", "grammar", "grammatical",
    "admission", "vocabulary", "freehand", "aspirant", "aspirants",
    # accepted British / Bangladeshi English forms and loanwords
    "centre", "centres", "favour", "favourite", "labour", "realise", "realised",
    "memorising", "memorise", "minimalize", "roadmap", "tech", "vocab", "hrs",
    "adda", "organise", "organised", "colour", "behaviour", "programme",
}

MISSING_APOSTROPHE = {"thats", "youre", "whos", "dont", "cant", "wont", "doesnt",
                      "didnt", "isnt", "wasnt", "shouldnt", "couldnt", "wouldnt",
                      "im", "ive", "theyre", "weve", "arent", "cann't"}
CONTRACTION_OK = re.compile(r"^(?:\w+n't|\w+'(?:s|re|ve|ll|d|m))$", re.I)

def check_spelling(sentence, spell):
    words = re.findall(r"[A-Za-z]+(?:'[a-z]+)?", sentence)
    hits = []
    for i, w in enumerate(words):
        if CONTRACTION_OK.match(w) and w.lower() not in MISSING_APOSTROPHE:
            continue
        if w.lower() in MISSING_APOSTROPHE:
            hits.append(w)
            continue
        base = w.split("'")[0]
        lw = base.lower()
        if len(lw) < 3 or lw in SPELL_WHITELIST:
            continue
        if base.isupper():                      # acronyms: DU, HSC
            continue
        if i > 0 and base[0].isupper():         # proper nouns mid-sentence
            continue
        if lw in spell.unknown([lw]):
            hits.append(w)
    return hits

def detect(sentence, spell):
    """Return list of (error_type, matched_text)."""
    found = []
    s = sentence.strip()

    for etype, patterns in RULES.items():
        for p in patterns:
            m = re.search(p, s, re.I)
            if m:
                found.append((etype, m.group(0)))
                break   # one hit per category per sentence keeps counts stable

    # Articles: a + vowel / an + consonant (small exception list)
    for m in re.finditer(r"\b(a)\s+((?!uni|use|usu|eu|one|once|ub)[aeiou]\w*)", s, re.I):
        found.append(("Article", m.group(0)))
    for m in re.finditer(r"\b(an)\s+((?!hour|honest|honou?r|heir)[bcdfgjklmnpqrstvwxyz]\w*)", s, re.I):
        found.append(("Article", m.group(0)))
    for m in re.finditer(r"\bthe\s+(?:Student|Preparation)\b", s):
        found.append(("Article", m.group(0)))

    # Capitalisation
    if s[0].islower():
        found.append(("Capitalisation", s.split()[0]))
    for m in re.finditer(r"(?<![A-Za-z])i(?![A-Za-z'.])(?!\.)", s):
        found.append(("Capitalisation", "i"))
        break
    for m in re.finditer(r"[,;]\s+(Usually|Then|Also|Only|Use|Stay|Study|But|And|So)\b", s):
        found.append(("Capitalisation", m.group(0)))
        break

    # Punctuation / spacing / run-on
    if re.search(r"\s+[,.;:!?]", s):
        found.append(("Punctuation / spacing", re.search(r"\s+[,.;:!?]", s).group(0)))
    elif re.search(r"[,;][A-Za-z]|[a-z][.!?][A-Z]", s):
        found.append(("Punctuation / spacing", re.search(r"[,;][A-Za-z]|[a-z][.!?][A-Z]", s).group(0)))
    if not re.search(r'[.!?"\')]$', s):
        found.append(("Punctuation / spacing", "(no end punctuation)"))

    # Apostrophe errors (plural written as possessive, It's + verb)
    for m in re.finditer(r"\b(?:most|many|all|some|other|few|these|those)\s+\w+s?'s\b|\bIt's\s+(?:is|all\s+depends|depends)\b", s, re.I):
        found.append(("Apostrophe / possessive", m.group(0)))
        break

    # Run-on / very long sentence
    if len(s.split()) >= 45:
        found.append(("Run-on / sentence structure", " ".join(s.split()[:6]) + " ..."))

    # Spelling
    for w in check_spelling(s, spell):
        etype = "Apostrophe / possessive" if w.lower() in MISSING_APOSTROPHE else "Spelling"
        found.append((etype, w))

    return found

# ---------------------------------------------------------------- pipeline
def run(path):
    os.makedirs(OUT, exist_ok=True)
    samples = load_samples(path)
    spell = SpellChecker()
    spell.word_frequency.load_words(SPELL_WHITELIST)

    rows, per_essay = [], []
    total_words = total_sents = 0
    for sid, text in sorted(samples.items()):
        words = len(re.findall(r"[A-Za-z']+", text))
        sents = split_sentences(text)
        total_words += words
        total_sents += len(sents)
        n_err = 0
        for s in sents:
            for etype, match in detect(s, spell):
                rows.append({"sample": sid, "error_type": etype,
                             "matched_text": match, "sentence": s})
                n_err += 1
        per_essay.append({"sample": sid, "words": words, "sentences": len(sents),
                          "errors": n_err,
                          "errors_per_100_words": round(n_err / words * 100, 1)})

    df = pd.DataFrame(rows)
    df.to_csv(f"{OUT}/flagged_errors.csv", index=False)
    pd.DataFrame(per_essay).to_csv(f"{OUT}/per_essay.csv", index=False)

    counts = df["error_type"].value_counts()
    summary = pd.DataFrame({
        "count": counts,
        "percent_of_errors": (counts / counts.sum() * 100).round(1),
        "per_100_words": (counts / total_words * 100).round(2),
        "essays_affected": df.groupby("error_type")["sample"].nunique(),
    }).sort_values("count", ascending=False)
    summary.to_csv(f"{OUT}/summary.csv")

    print(f"Essays analysed (duplicates removed): {len(samples)}")
    print(f"Total words: {total_words}   Sentences: {total_sents}")
    print("\n=== Error types ranked by share ===")
    print(summary.to_string())
    print(f"\nTotal flagged errors: {counts.sum()}")
    plot(summary)

def plot(summary):
    s = summary["percent_of_errors"].sort_values()
    ax = s.plot(kind="barh", figsize=(8, 5), color="#3b6ea5")
    ax.set_xlabel("% of all detected errors")
    ax.set_title("Error types in students' written English")
    for i, v in enumerate(s):
        ax.text(v + 0.3, i, f"{v}%", va="center")
    plt.tight_layout()
    plt.savefig(f"{OUT}/errors_chart.png", dpi=150)

if __name__ == "__main__":
    run(sys.argv[1] if len(sys.argv) > 1 else "data.docx")
