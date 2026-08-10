# Academic Word List (AWL) — Anki deck (RU / ZH)

An Anki deck generator for the **Academic Word List** (Coxhead, 2000): all
**570 word families**, with **every related word form as its own separate
card** (analyse → *analysed, analyser, analysers, analyses, analysing,
analysis, analyst, analytic, analytical, …*). After removing forms shared
between families this comes to **3,110 unique cards**.

Each card follows this layout:

```
pull through (v)
────────────────────────────────
RU:            выкарабкаться, выжить
ZH:            渡过难关, 痊愈
Определение:   выжить или справиться с очень тяжёлой болезнью …
Пример:        The doctors were really worried … he pulled through.   🔊
Сочетаемость:  pull through surgery, pull through crisis               🔊
```

- **Front:** the word + part of speech.
- **Back:** Russian + Chinese translations, a Russian definition, an English
  example sentence, and typical English collocations. The example and the
  collocations each have a **playable audio** button.

Cards are tagged so you can study by sublist (`AWL::sublist01` … `sublist10`),
by family (`family::analyse`), and by `headword` vs `wordform`. Inside Anki
they are organised into sub-decks `Academic Word List (AWL)::Sublist 01…10`.

## What's in this folder

| File | Description |
|------|-------------|
| `awl_families.json` | The 570 headwords, their sublist, and all related word forms (parsed from eapfoundation.com). |
| `build_deck.py` | The generator. |
| `by_sublist/AWL_sublistNN.apkg` | Ready-to-import decks, **one per sublist**, with audio. |
| `AWL_RU_ZH_text_only.apkg` | Single-file deck with everything **except** audio (small, quick to import). |
| `cache/` | Cached translations & collocations so rebuilds are instant and don't re-hit any API. |

> Combined **with-audio** decks are split per sublist because a single
> `.apkg` with ~6,000 audio clips would exceed GitHub's 100 MB-per-file limit.
> The generated `media/` folder and the big combined `AWL_RU_ZH.apkg` are
> `.gitignore`d; rebuild them any time with the command below.

## How the content is generated (all free, no API keys)

| Field | Source |
|-------|--------|
| Part of speech, English definition, English example | **WordNet** (offline, via `nltk`) |
| RU / ZH translations, Russian definition | **Google Translate** (`deep-translator`) |
| Collocations (Сочетаемость) | **Datamuse** bigram API, filtered to real English words |
| Audio (Пример + Сочетаемость) | **gTTS** (Google Text-to-Speech) |

Everything is cached on disk, so a re-run only does the work that hasn't been
done yet — the build is fully resumable.

## Usage

```bash
pip install -r requirements.txt

# Full deck with audio, split into one .apkg per sublist:
python build_deck.py --split-by-sublist

# Text-only single-file deck (fast, small):
python build_deck.py --no-audio --out AWL_RU_ZH_text_only.apkg

# Quick test on the first 20 word forms:
python build_deck.py --limit 20

# Only certain sublists:
python build_deck.py --sublists 1 2 3
```

Then in Anki: **File → Import** the `.apkg` file(s).

## Notes & limitations

- Translations, definitions and collocations are produced automatically. They
  are good for study but not hand-curated; some rare word forms may have a
  definition drawn from the family headword or an approximate collocation.
- Audio is English text-to-speech for the example sentence and the collocations.
- Source of the word-family data: the AWL as published by Averil Coxhead,
  Victoria University of Wellington.
