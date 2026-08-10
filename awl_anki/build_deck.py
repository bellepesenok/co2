#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Build an Anki deck for the Academic Word List (AWL).

The deck contains ONE card per word form: every one of the 570 AWL headwords
*plus* every related word form (analyse -> analysed, analyser, analysis, ...),
so each inflection/derivation becomes its own separate card (~3147 cards).

Each card mirrors the reference layout:

    pull through (phr v)
    RU:            выкарабкаться, выжить
    ZH:            渡过难关, 痊愈
    Определение:   <russian definition>
    Пример:        <english example>          [audio]
    Сочетаемость:  <english collocations>     [audio]

Content sources (all free, no API key required):
  * Part of speech / English definition / example -> WordNet (offline, nltk)
  * RU + ZH translations, RU definition            -> Google Translate (deep_translator)
  * Collocations                                   -> Datamuse bigram API
  * Audio (Пример + Сочетаемость)                  -> gTTS (Google Text-to-Speech)

Everything is cached on disk so the build is fully resumable: re-running only
does the work that has not been done yet.

Usage:
    python build_deck.py                 # full deck, with audio
    python build_deck.py --no-audio      # skip TTS (much faster / smaller)
    python build_deck.py --limit 20      # quick test on the first 20 forms
    python build_deck.py --sublists 1 2  # only sublists 1 and 2
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import random
import re
import threading
import time
import urllib.parse
import urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed

HERE = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(HERE, "awl_families.json")
CACHE_DIR = os.path.join(HERE, "cache")
MEDIA_DIR = os.path.join(HERE, "media")

# ---------------------------------------------------------------------------
# Small on-disk cache helpers (thread-safe)
# ---------------------------------------------------------------------------
_locks: dict[str, threading.Lock] = {}
_caches: dict[str, dict] = {}
_caches_lock = threading.Lock()


def _cache_path(name: str) -> str:
    return os.path.join(CACHE_DIR, name + ".json")


def load_cache(name: str) -> dict:
    with _caches_lock:
        if name not in _caches:
            _locks[name] = threading.Lock()
            p = _cache_path(name)
            if os.path.exists(p):
                try:
                    _caches[name] = json.load(open(p, encoding="utf-8"))
                except Exception:
                    _caches[name] = {}
            else:
                _caches[name] = {}
    return _caches[name]


def save_cache(name: str) -> None:
    load_cache(name)
    os.makedirs(CACHE_DIR, exist_ok=True)
    p = _cache_path(name)
    with _locks[name]:
        tmp = p + f".tmp{os.getpid()}"
        json.dump(_caches[name], open(tmp, "w", encoding="utf-8"),
                  ensure_ascii=False)
        os.replace(tmp, p)


def cache_get(name: str, key: str):
    c = load_cache(name)
    with _locks[name]:
        return c.get(key)


def cache_set(name: str, key: str, value) -> None:
    c = load_cache(name)
    with _locks[name]:
        c[key] = value


# ---------------------------------------------------------------------------
# WordNet: part of speech, English definition, example
# ---------------------------------------------------------------------------
_wn = None
_wn_lock = threading.Lock()

POS_LABEL = {"n": "n", "v": "v", "a": "adj", "s": "adj", "r": "adv"}


def get_wordnet():
    global _wn
    if _wn is None:
        with _wn_lock:
            if _wn is None:
                import nltk
                for pkg in ("wordnet", "omw-1.4"):
                    try:
                        nltk.data.find(f"corpora/{pkg}")
                    except LookupError:
                        nltk.download(pkg, quiet=True)
                from nltk.corpus import wordnet as wn
                _ = wn.synsets("test")  # force load
                _wn = wn
    return _wn


def wordnet_info(word: str, headword: str) -> dict:
    """Return {pos, definition_en, example_en} using WordNet.

    Falls back from the exact form -> morphy base -> family headword so that
    inflected forms (which WordNet does not list directly) still get content.
    """
    wn = get_wordnet()
    candidates = wn.synsets(word.replace(" ", "_"))
    if not candidates:
        for pos in ("n", "v", "a", "r"):
            base = wn.morphy(word, pos)
            if base:
                s = wn.synsets(base, pos=pos)
                if s:
                    candidates = s
                    break
    if not candidates:
        candidates = wn.synsets(headword.replace(" ", "_"))
    if not candidates:
        return {"pos": "", "definition_en": "", "example_en": ""}

    syn = candidates[0]
    pos = POS_LABEL.get(syn.pos(), "")
    definition = syn.definition() or ""
    example = ""
    stem = re.sub(r"(ing|ed|es|s)$", "", word.lower())
    for s in candidates:
        for ex in s.examples():
            if word.lower() in ex.lower() or (len(stem) > 3 and stem in ex.lower()):
                example = ex
                break
        if example:
            break
    if not example:
        for s in candidates:
            if s.examples():
                example = s.examples()[0]
                break
    return {"pos": pos, "definition_en": definition, "example_en": example}


# ---------------------------------------------------------------------------
# Translation (Google via deep_translator) with retry + cache
# ---------------------------------------------------------------------------
def _retry(fn, *, tries=4, base=1.5):
    last = None
    for i in range(tries):
        try:
            return fn()
        except Exception as e:  # noqa: BLE001
            last = e
            time.sleep(base * (2 ** i) + random.random())
    raise last


def translate(text: str, target: str) -> str:
    text = (text or "").strip()
    if not text:
        return ""
    key = f"{target}\t{text}"
    cached = cache_get("translate", key)
    if cached is not None:
        return cached
    from deep_translator import GoogleTranslator

    def do():
        return GoogleTranslator(source="en", target=target).translate(text)

    try:
        out = _retry(do) or ""
    except Exception:
        out = ""
    cache_set("translate", key, out)
    return out


# ---------------------------------------------------------------------------
# Collocations via Datamuse bigram relations
# ---------------------------------------------------------------------------
STOPWORDS = {
    "the", "a", "an", "of", "to", "and", "or", "in", "on", "for", "with",
    "is", "are", "was", "were", "be", "been", "being", "it", "its", "this",
    "that", "these", "those", "as", "at", "by", "from", "into", "out", "up",
    "down", "would", "will", "can", "could", "should", "may", "might", "must",
    "i", "you", "he", "she", "they", "we", "him", "her", "them", "us", "his",
    "their", "our", "your", "my", "me", "not", "no", "do", "does", "did",
    "has", "have", "had", "so", "if", "then", "than", "but", "which", "who",
    "what", "when", "where", "how", "s", "t", "re", "ll", "ve", "d", "m",
    "there", "here", "one", "all", "any", "some", "such", "more", "most",
}


_TOKEN_RE = re.compile(r"^[a-z][a-z'-]{1,}$")

# Common non-English function words that leak into Datamuse's Google-Books
# bigram data (esp. for British spellings like "analyse" -> "analyse des").
FOREIGN_STOPWORDS = {
    "des", "der", "die", "den", "dem", "das", "ein", "eine", "und", "von",
    "zur", "zum", "im", "les", "une", "la", "le", "el", "los", "las", "del",
    "della", "dei", "du", "aux", "ce", "cette", "deux", "donnees", "donnée",
    "données", "et", "y", "il", "lo", "pour", "avec", "dans", "sur", "que",
    "qui", "sono", "che", "col", "nel", "als", "auf", "mit", "sich", "wir",
    "sie", "nach", "por", "para", "con", "una", "uno", "ele", "als",
    "de", "en", "al", "et", "ou", "se", "je", "ne", "za", "na", "ki",
}


def is_english_word(token: str) -> bool:
    """True if `token` is a genuine, reasonably common English word.

    Used to drop Datamuse noise such as foreign-language bigram partners
    (des, une, der, ...) that show up for British spellings.
    """
    if not _TOKEN_RE.match(token) or token in FOREIGN_STOPWORDS:
        return False
    from wordfreq import zipf_frequency
    if zipf_frequency(token, "en") < 2.5:
        return False
    wn = get_wordnet()
    if wn.synsets(token):
        return True
    for pos in ("n", "v", "a", "r"):
        if wn.morphy(token, pos):
            return True
    return False


def _datamuse(params: dict) -> list[str]:
    url = "https://api.datamuse.com/words?" + urllib.parse.urlencode(params)

    def do():
        req = urllib.request.Request(url, headers={"User-Agent": "awl-anki"})
        with urllib.request.urlopen(req, timeout=15) as r:
            return json.loads(r.read().decode("utf-8"))

    try:
        data = _retry(do, tries=3)
    except Exception:
        return []
    out = []
    for item in data:
        w = (item.get("word") or "").strip().lower()
        if w and w not in STOPWORDS and is_english_word(w):
            out.append(w)
    return out


def collocations(word: str, pos: str) -> str:
    key = word.lower()
    cached = cache_get("collocations", key)
    if cached is not None:
        return cached
    word = word.strip()
    if " " in word:
        followers = _datamuse({"rel_bga": word, "max": 8})
        phrases = [f"{word} {f}" for f in followers[:3]]
    else:
        followers = _datamuse({"rel_bga": word, "max": 8})
        preceders = _datamuse({"rel_bgb": word, "max": 8})
        phrases = [f"{word} {f}" for f in followers[:2]]
        phrases += [f"{p} {word}" for p in preceders[:2]]
    seen, out = set(), []
    for ph in phrases:
        if ph.lower() not in seen:
            seen.add(ph.lower())
            out.append(ph)
    result = ", ".join(out[:4])
    cache_set("collocations", key, result)
    return result


# ---------------------------------------------------------------------------
# Audio via gTTS (cached to media/ as sha1.mp3)
# ---------------------------------------------------------------------------
def audio_file(text: str) -> str:
    """Create (or reuse) an mp3 for `text`; return its basename, or "".

    The filename is content-addressed so identical text is only synthesised
    once and the build is resumable.
    """
    text = (text or "").strip()
    if not text:
        return ""
    h = hashlib.sha1(text.encode("utf-8")).hexdigest()[:16]
    fname = f"awl_{h}.mp3"
    fpath = os.path.join(MEDIA_DIR, fname)
    if os.path.exists(fpath) and os.path.getsize(fpath) > 0:
        return fname
    os.makedirs(MEDIA_DIR, exist_ok=True)
    from gtts import gTTS

    def do():
        tts = gTTS(text=text, lang="en")
        tmp = fpath + f".tmp{threading.get_ident()}"
        tts.save(tmp)
        os.replace(tmp, fpath)

    try:
        _retry(do, tries=4, base=2.0)
    except Exception:
        return ""
    return fname


# ---------------------------------------------------------------------------
# Build the flat list of word-form entries from the family data
# ---------------------------------------------------------------------------
def build_entries(sublists=None, limit=None) -> list[dict]:
    families = json.load(open(DATA_FILE, encoding="utf-8"))
    entries, seen = [], set()
    for fam in families:
        if sublists and fam["sublist"] not in sublists:
            continue
        head = fam["headword"]
        forms = [(head, True)] + [(f, False) for f in fam["related"]]
        for word, is_head in forms:
            wl = word.lower()
            if wl in seen:
                continue
            seen.add(wl)
            entries.append({
                "word": word,
                "headword": head,
                "sublist": fam["sublist"],
                "is_headword": is_head,
            })
    if limit:
        entries = entries[:limit]
    return entries


# ---------------------------------------------------------------------------
# Per-entry content enrichment (network heavy) -> mutates entry in place
# ---------------------------------------------------------------------------
def enrich(entry: dict, with_audio: bool) -> dict:
    word, head = entry["word"], entry["headword"]
    wn = wordnet_info(word, head)
    entry["pos"] = wn["pos"]
    entry["definition_en"] = wn["definition_en"]
    entry["example_en"] = wn["example_en"]
    entry["ru"] = translate(word, "ru")
    entry["zh"] = translate(word, "zh-CN")
    entry["definition_ru"] = translate(wn["definition_en"], "ru")
    entry["collocations"] = collocations(word, wn["pos"])
    if with_audio:
        entry["example_audio"] = audio_file(entry["example_en"])
        entry["collocations_audio"] = audio_file(
            entry["collocations"].replace(",", ";"))
    else:
        entry["example_audio"] = ""
        entry["collocations_audio"] = ""
    return entry


# ---------------------------------------------------------------------------
# genanki model + deck assembly
# ---------------------------------------------------------------------------
MODEL_ID = 1980700000123  # fixed so re-imports update instead of duplicate
DECK_ID = 1980700000456

CSS = """
.card {
  font-family: -apple-system, "Segoe UI", Arial, sans-serif;
  font-size: 18px;
  color: #222;
  background: #f4f4f4;
  text-align: left;
  padding: 18px 22px;
}
.headword { font-size: 26px; font-weight: 600; margin-bottom: 2px; }
.pos { color: #666; font-weight: 400; font-size: 20px; }
hr { border: none; border-top: 1px solid #ccc; margin: 14px 0; }
.row { margin: 10px 0; line-height: 1.5; }
.label { font-weight: 700; margin-right: 6px; }
.zh { font-size: 20px; }
.colloc { color: #2f7d32; }
.tag { color: #999; font-size: 13px; margin-top: 14px; }
.replay-button svg circle { fill: #fff; stroke: #333; }
.replay-button svg path { fill: #333; }
"""

FRONT = """
<div class="headword">{{Word}} {{#POS}}<span class="pos">({{POS}})</span>{{/POS}}</div>
"""

BACK = """
<div class="headword">{{Word}} {{#POS}}<span class="pos">({{POS}})</span>{{/POS}}</div>
<hr>
{{#RU}}<div class="row"><span class="label">RU:</span>{{RU}}</div>{{/RU}}
{{#ZH}}<div class="row zh"><span class="label">ZH:</span>{{ZH}}</div>{{/ZH}}
{{#DefinitionRU}}<div class="row"><span class="label">Определение:</span>{{DefinitionRU}}</div>{{/DefinitionRU}}
{{#Example}}<div class="row"><span class="label">Пример:</span>{{Example}} {{ExampleAudio}}</div>{{/Example}}
{{#Collocations}}<div class="row colloc"><span class="label">Сочетаемость:</span>{{Collocations}} {{CollocationsAudio}}</div>{{/Collocations}}
<div class="tag">AWL Sublist {{Sublist}} · family: {{Family}}</div>
"""


def make_model():
    import genanki
    return genanki.Model(
        MODEL_ID,
        "AWL Word (RU/ZH)",
        fields=[
            {"name": "Word"}, {"name": "POS"}, {"name": "RU"}, {"name": "ZH"},
            {"name": "DefinitionRU"}, {"name": "Example"},
            {"name": "ExampleAudio"}, {"name": "Collocations"},
            {"name": "CollocationsAudio"}, {"name": "Family"},
            {"name": "Sublist"},
        ],
        templates=[{"name": "Card 1", "qfmt": FRONT, "afmt": BACK}],
        css=CSS,
    )


def sound_tag(fname: str) -> str:
    return f"[sound:{fname}]" if fname else ""


def make_note(model, e: dict):
    import genanki
    ex_audio = sound_tag(e.get("example_audio", ""))
    col_audio = sound_tag(e.get("collocations_audio", ""))
    note = genanki.Note(
        model=model,
        fields=[
            e["word"], e.get("pos", ""), e.get("ru", ""), e.get("zh", ""),
            e.get("definition_ru", ""), e.get("example_en", ""), ex_audio,
            e.get("collocations", ""), col_audio, e["headword"],
            str(e["sublist"]),
        ],
        tags=[
            f"AWL::sublist{e['sublist']:02d}",
            f"family::{e['headword'].replace(' ', '_')}",
            "headword" if e["is_headword"] else "wordform",
        ],
        guid=genanki.guid_for("awl", e["word"].lower()),
    )
    return note


def entry_media(e: dict) -> list[str]:
    out = []
    if e.get("example_audio"):
        out.append(os.path.join(MEDIA_DIR, e["example_audio"]))
    if e.get("collocations_audio"):
        out.append(os.path.join(MEDIA_DIR, e["collocations_audio"]))
    return out


def build_package(entries: list[dict], out_path: str) -> None:
    """One .apkg with a sub-deck per sublist (Academic Word List (AWL)::Sublist NN)."""
    import genanki
    model = make_model()
    decks, media = [], []
    by_sub = {}
    for e in entries:
        by_sub.setdefault(e["sublist"], []).append(e)
    for sub in sorted(by_sub):
        deck = genanki.Deck(DECK_ID + sub,
                            f"Academic Word List (AWL)::Sublist {sub:02d}")
        for e in by_sub[sub]:
            deck.add_note(make_note(model, e))
            media += entry_media(e)
        decks.append(deck)
    pkg = genanki.Package(decks)
    pkg.media_files = sorted(set(media))
    pkg.write_to_file(out_path)


def build_split_by_sublist(entries: list[dict], out_dir: str) -> list[str]:
    """One .apkg per sublist so each file stays comfortably under 100 MB
    (important because GitHub rejects single files larger than 100 MB)."""
    import genanki
    os.makedirs(out_dir, exist_ok=True)
    model = make_model()
    by_sub = {}
    for e in entries:
        by_sub.setdefault(e["sublist"], []).append(e)
    paths = []
    for sub in sorted(by_sub):
        deck = genanki.Deck(DECK_ID + sub,
                            f"Academic Word List (AWL)::Sublist {sub:02d}")
        media = []
        for e in by_sub[sub]:
            deck.add_note(make_note(model, e))
            media += entry_media(e)
        pkg = genanki.Package(deck)
        pkg.media_files = sorted(set(media))
        p = os.path.join(out_dir, f"AWL_sublist{sub:02d}.apkg")
        pkg.write_to_file(p)
        paths.append(p)
    return paths


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main() -> None:
    ap = argparse.ArgumentParser(description="Build the AWL Anki deck.")
    ap.add_argument("--out", default=os.path.join(HERE, "AWL_RU_ZH.apkg"))
    ap.add_argument("--limit", type=int, default=None,
                    help="only process the first N word forms (for testing)")
    ap.add_argument("--sublists", type=int, nargs="*", default=None,
                    help="restrict to these sublist numbers")
    ap.add_argument("--no-audio", action="store_true",
                    help="skip gTTS audio generation")
    ap.add_argument("--split-by-sublist", action="store_true",
                    help="write one .apkg per sublist into <out dir>/by_sublist")
    ap.add_argument("--workers", type=int, default=8)
    ap.add_argument("--save-every", type=int, default=50)
    args = ap.parse_args()

    with_audio = not args.no_audio
    entries = build_entries(sublists=args.sublists, limit=args.limit)
    total = len(entries)
    print(f"[awl] {total} word forms to process (audio={'on' if with_audio else 'off'})",
          flush=True)

    get_wordnet()  # warm up once in the main thread

    done = 0
    lock = threading.Lock()
    with ThreadPoolExecutor(max_workers=args.workers) as ex:
        futs = {ex.submit(enrich, e, with_audio): e for e in entries}
        for fut in as_completed(futs):
            e = futs[fut]
            try:
                fut.result()
            except Exception as exc:  # noqa: BLE001
                print(f"[warn] {e['word']}: {exc}", flush=True)
            with lock:
                done += 1
                if done % args.save_every == 0 or done == total:
                    for name in ("translate", "collocations"):
                        if name in _caches:
                            save_cache(name)
                    print(f"[awl] {done}/{total} ({done*100//total}%)", flush=True)

    for name in ("translate", "collocations"):
        if name in _caches:
            save_cache(name)

    print("[awl] assembling .apkg ...", flush=True)
    build_package(entries, args.out)
    # also dump the enriched dataset for inspection / reuse
    json.dump(entries, open(os.path.join(HERE, "awl_enriched.json"), "w",
                            encoding="utf-8"), ensure_ascii=False, indent=2)
    size_mb = os.path.getsize(args.out) / 1e6
    print(f"[awl] wrote {args.out} ({size_mb:.1f} MB, {total} cards)", flush=True)

    if args.split_by_sublist:
        out_dir = os.path.join(os.path.dirname(args.out) or ".", "by_sublist")
        paths = build_split_by_sublist(entries, out_dir)
        for p in paths:
            print(f"[awl] wrote {p} ({os.path.getsize(p)/1e6:.1f} MB)", flush=True)


if __name__ == "__main__":
    main()
