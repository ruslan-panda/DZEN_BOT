from __future__ import annotations

import argparse
import base64
import html
import io
import json
import logging
import math
import os
import re
import sys
import threading
import time
import xml.etree.ElementTree as ET
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, TimeoutError as FuturesTimeoutError, as_completed
from datetime import date, datetime
from pathlib import Path
from urllib.parse import quote

import requests
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from PIL import Image, ImageDraw
from requests.adapters import HTTPAdapter
from telethon.sessions import StringSession
from telethon.sync import TelegramClient
from urllib3.util.retry import Retry

BASE_DIR = Path(__file__).resolve().parent

# ----------------- СТРУКТУРИРОВАННОЕ ЛОГИРОВАНИЕ ДЛЯ VPS -----------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("AlgoMoneyAutopilot")


# ----------------- АВТОЗАГРУЗЧИК .ENV -----------------
def load_custom_env() -> None:
    """Загружает переменные окружения из .env в корне проекта или .venv."""
    candidate_paths = [
        BASE_DIR / ".env",
        BASE_DIR / ".venv" / ".env",
        BASE_DIR / "venv" / ".env",
        BASE_DIR / "config.env",
    ]
    for path in candidate_paths:
        if path.is_file():
            with open(path, "r", encoding="utf-8") as f:
                for raw_line in f:
                    line = raw_line.strip()
                    if not line or line.startswith("#") or "=" not in line:
                        continue
                    key, val = line.split("=", 1)
                    key = key.strip()
                    val = val.strip().strip('"').strip("'")
                    if key and key not in os.environ:
                        os.environ[key] = val
            log.info("Конфигурация загружена из: %s", path)
            return


load_custom_env()


def parse_tg_channel(val: str) -> int | str:
    val = val.strip()
    if val == "@algo_money_sync":
        return "@algo_money_tech_dzen_sync"
    return int(val) if val.lstrip("-").isdigit() else val


# ======================= КОНФИГУРАЦИЯ ИЗ .ENV =======================
OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY", "").strip()
DZEN_CHANNEL_URL = os.getenv("DZEN_CHANNEL_URL", "https://dzen.ru/algomoney_ru").strip()

ENABLE_TELEGRAM = os.getenv("ENABLE_TELEGRAM", "True").lower() in ("true", "1", "yes")
SAVE_LOCAL_DOCX = os.getenv("SAVE_LOCAL_DOCX", "True").lower() in ("true", "1", "yes")
TG_API_ID = int(os.getenv("TG_API_ID", "0"))
TG_API_HASH = os.getenv("TG_API_HASH", "").strip()
TG_STRING_SESSION_ENV = os.getenv("TG_STRING_SESSION", "").strip()
TG_TECH_CHANNEL = parse_tg_channel(os.getenv("TG_TECH_CHANNEL", "@algo_money_tech_dzen_sync"))
TG_PUBLIC_CHANNEL = parse_tg_channel(os.getenv("TG_PUBLIC_CHANNEL", "@algomoney_ru"))
DZEN_SYNC_WAIT_SECONDS = int(os.getenv("DZEN_SYNC_WAIT_SECONDS", "45"))

SERVER_HEADLESS = os.getenv("SERVER_HEADLESS", "True").lower() in ("true", "1", "yes")

# Лимиты для ОДНОГО сообщения Telegram (4096 HTML-симв.; >3000 видимых симв. = Статья на Дзене)
TG_SINGLE_MSG_HTML_LIMIT = 4050
ARTICLE_MIN_CHARS = 3350
ARTICLE_TARGET_MAX_CHARS = 3700
REQUIRED_SECTIONS_COUNT = 3

# Дата окончания бесплатного периода Dots3-Note Preview
DOTS3_FREE_DEADLINE = date(2026, 12, 31)

PROXY_URL = os.getenv("PROXY_URL", "").strip()
if PROXY_URL:
    os.environ["HTTP_PROXY"] = PROXY_URL
    os.environ["HTTPS_PROXY"] = PROXY_URL
    ru_bypass = (
        "dzen.ru,.dzen.ru,habr.com,.habr.com,lenta.ru,.lenta.ru,"
        "kommersant.ru,.kommersant.ru,rbc.ru,.rbc.ru,localhost,127.0.0.1"
    )
    existing_no_proxy = os.environ.get("NO_PROXY", "")
    os.environ["NO_PROXY"] = f"{existing_no_proxy},{ru_bypass}".strip(",")

CHANNEL_NAME = "Алгоритм Денег"
CHANNEL_THEME = (
    "Серьёзная финансовая аналитика, личные финансы, инвестиции, фондовый рынок, "
    "дивиденды, налоги и вклады простым и точным языком с расчётами в рублях и процентах."
)

RSS_FEEDS = [
    ("РБК Экономика", "https://rssexport.rbc.ru/rbcnews/news/30/full.rss"),
    ("Коммерсантъ Экономика", "https://www.kommersant.ru/RSS/section-economics.xml"),
    ("Lenta Экономика", "https://lenta.ru/rss/news/economics"),
    ("Хабр Финансы", "https://habr.com/ru/rss/hubs/finance/articles/?fl=ru"),
]

TG_STRING_SESSION_FILE = BASE_DIR / "tg_session_string.txt"
HISTORY_FILE = BASE_DIR / "used_topics.txt"
KNOWLEDGE_BASE_DIR = BASE_DIR / "knowledge_base"
OUTPUT_DIR = BASE_DIR / "articles"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8",
}


# ----------------- HTTP-СЕССИИ С ПУЛОМ СОЕДИНЕНИЙ -----------------
def build_http_session(retries: int = 1, direct: bool = False) -> requests.Session:
    session = requests.Session()
    session.headers.update(HEADERS)
    if direct:
        session.trust_env = False
    retry = Retry(
        total=retries,
        backoff_factor=0.3,
        status_forcelist=[500, 502, 504],
        allowed_methods=["GET", "POST"],
    )
    adapter = HTTPAdapter(max_retries=retry, pool_connections=16, pool_maxsize=16)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    return session


HTTP = build_http_session(retries=1, direct=False)
HTTP_DIRECT = build_http_session(retries=0, direct=True)


# ======================================================================
# РОУТЕР ТОЛЬКО ПРОВЕРЕННЫХ РАБОЧИХ МОДЕЛЕЙ OPENROUTER (:free)
# ======================================================================
MODEL_PARTICLE = "particle-ai/particle-1:free"
MODEL_DOTS3 = "dots-studio/dots-3-note-preview:free"
MODEL_NEMOTRON_ULTRA = "nvidia/nemotron-3-ultra-550b-a55b:free"
MODEL_ERNIE = "baidu/ernie-5.5-flash-preview:free"

ReasoningMode = str | None
ModelSpec = tuple[str, ReasoningMode]

# Жёстко исключаем все модели, которые системно падали (402/403/404/429/timeout/пустой ответ)
EXCLUDED_MODEL_SUBSTRINGS = (
    "embed",
    "rerank",
    "voice",
    "heygen",
    "fish-audio",
    "s2.1",
    "flint",
    "north-mini",
    "1.7b",
    "content-safety",
    "inkling",
    "lightning",
    "hydrogen",
    "gemma-4-31b",
    "ling-3.2",
    "gemini-",
    ":batch",
    "-vl",
)

MODEL_TIMEOUTS: dict[str, int] = {
    MODEL_NEMOTRON_ULTRA: 110,
    MODEL_ERNIE: 75,
    MODEL_DOTS3: 70,
    MODEL_PARTICLE: 65,
}

SESSION_SKIP_KEYS: set[ModelSpec] = set()
_CATALOG_LOCK = threading.Lock()
OPENROUTER_FREE_MODELS: list[str] = []
CATALOG_LOADED: bool = False


def has_valid_openrouter_key() -> bool:
    return bool(
        OPENROUTER_API_KEY
        and OPENROUTER_API_KEY.startswith("sk-or-")
        and "ВСТАВЬ" not in OPENROUTER_API_KEY
    )


def discover_openrouter_free_models() -> list[str]:
    """Загружает актуальный список рабочих :free моделей OpenRouter (исключая сбойные)."""
    global CATALOG_LOADED
    with _CATALOG_LOCK:
        if CATALOG_LOADED:
            return OPENROUTER_FREE_MODELS
        CATALOG_LOADED = True

        try:
            resp = HTTP.get("https://openrouter.ai/api/v1/models", timeout=(5, 8))
            if resp.status_code == 200:
                for item in resp.json().get("data", []):
                    slug = item.get("id", "").strip()
                    slug_low = slug.lower()
                    if not slug:
                        continue
                    is_free = slug_low.endswith(":free") or str((item.get("pricing") or {}).get("prompt", "")) == "0"
                    if is_free and not any(bad in slug_low for bad in EXCLUDED_MODEL_SUBSTRINGS):
                        OPENROUTER_FREE_MODELS.append(slug)

                log.info(
                    "Каталог OpenRouter: %d проверенных :free моделей (%s)",
                    len(OPENROUTER_FREE_MODELS),
                    ", ".join(OPENROUTER_FREE_MODELS[:5]),
                )
        except Exception as e:
            log.warning("Каталог OpenRouter недоступен (%s) — используем встроенную четвёрку.", str(e)[:70])

        return OPENROUTER_FREE_MODELS


def build_queue(primary: list[ModelSpec], tail_limit: int = 3) -> list[ModelSpec]:
    """Убирает дубликаты и добавляет живой резерв из каталога OpenRouter."""
    unique: list[ModelSpec] = []
    seen_slugs: set[str] = set()

    for slug, mode in primary:
        if not slug or (slug, mode) in SESSION_SKIP_KEYS or slug in seen_slugs:
            continue
        if slug == MODEL_DOTS3 and datetime.now().date() > DOTS3_FREE_DEADLINE:
            continue
        unique.append((slug, mode))
        seen_slugs.add(slug)

    added = 0
    for slug in discover_openrouter_free_models():
        if slug not in seen_slugs and (slug, "off") not in SESSION_SKIP_KEYS:
            unique.append((slug, "off"))
            seen_slugs.add(slug)
            added += 1
            if added >= tail_limit:
                break
    return unique


def get_model_timeout(model_name: str, default_sec: int = 70) -> int:
    low = model_name.lower()
    for key, t_sec in MODEL_TIMEOUTS.items():
        if key in low:
            return t_sec
    return default_sec


# ----------------- ПАМЯТЬ И ПАРАЛЛЕЛЬНЫЙ СБОР RSS -----------------
def load_used_topics() -> list[str]:
    if not HISTORY_FILE.is_file():
        return []
    with open(HISTORY_FILE, "r", encoding="utf-8") as f:
        return [line.strip() for line in f if line.strip()]


def save_used_topic(topic: str) -> None:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    with open(HISTORY_FILE, "a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] {topic}\n")


def _is_topic_already_used(title: str, used_topics: list[str]) -> bool:
    title_words = {w.lower() for w in re.findall(r"[а-яА-Яa-zA-Z0-9]{4,}", title)}
    if not title_words:
        return False
    for entry in used_topics[-60:]:
        clean_entry = re.sub(r"^\[.*?\]\s*(?:\[.*?\])?\s*", "", entry)
        entry_words = {w.lower() for w in re.findall(r"[а-яА-Яa-zA-Z0-9]{4,}", clean_entry)}
        if not entry_words:
            continue
        overlap = len(title_words & entry_words) / max(1, min(len(title_words), len(entry_words)))
        if overlap >= 0.65:
            return True
    return False


def _fetch_single_rss(source_name: str, rss_url: str) -> list[dict]:
    items: list[dict] = []
    try:
        resp = HTTP_DIRECT.get(rss_url, timeout=(3.0, 5.0))
        if resp.status_code == 200:
            root = ET.fromstring(resp.content)
            for item in root.findall(".//item")[:12]:
                title = (item.findtext("title") or "").strip()
                desc_raw = (item.findtext("description") or "").strip()
                full_text_raw = (
                    item.findtext("{http://purl.org/rss/1.0/modules/content/}encoded") or desc_raw
                )
                clean_desc = BeautifulSoup(full_text_raw, "html.parser").get_text(" ", strip=True)
                if title:
                    items.append(
                        {
                            "source": source_name,
                            "title": title,
                            "description": clean_desc[:1200],
                        }
                    )
    except Exception as e:
        log.warning("RSS '%s' пропущен: %s", source_name, str(e)[:65])
    return items


def collect_fresh_rss_parallel(used_topics: list[str]) -> list[dict]:
    """Параллельно собирает свежие новости из РБК, Коммерсантъ, Lenta и Хабр Финансы."""
    all_items: list[dict] = []
    seen_titles: set[str] = set()

    with ThreadPoolExecutor(max_workers=len(RSS_FEEDS)) as pool:
        futures = [pool.submit(_fetch_single_rss, name, url) for name, url in RSS_FEEDS]
        for fut in as_completed(futures):
            for entry in fut.result():
                norm_title = entry["title"].lower()
                if norm_title in seen_titles or _is_topic_already_used(entry["title"], used_topics):
                    continue
                seen_titles.add(norm_title)
                all_items.append(entry)

    return all_items[:36]


# ----------------- АВТО-РЕМОНТНИК JSON И ВЫЗОВЫ OPENROUTER -----------------
def repair_malformed_json(candidate: str) -> dict:
    """Ремонтирует неэкранированные кавычки, спецсимволы, переносы строк и обрезанный JSON."""
    start_idx = candidate.find("{")
    if start_idx != -1:
        candidate = candidate[start_idx:]

    end_idx = candidate.rfind("}")
    if end_idx != -1:
        candidate = candidate[: end_idx + 1]

    s = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", " ", candidate)
    s = re.sub(r'\\(?!["\\/bfnrtu])', r"\\\\", s)

    out: list[str] = []
    in_string = False
    escape = False

    for i, ch in enumerate(s):
        if escape:
            out.append(ch)
            escape = False
            continue
        if ch == "\\":
            out.append(ch)
            escape = True
            continue
        if ch == '"':
            if not in_string:
                in_string = True
                out.append(ch)
            else:
                rest = s[i + 1 :].lstrip()
                if not rest or rest[0] in (":", ",", "}", "]"):
                    in_string = False
                    out.append(ch)
                else:
                    out.append("«")
            continue
        if in_string and ch in ("\n", "\r"):
            out.append("\\n")
            continue
        out.append(ch)

    fixed = "".join(out)
    if in_string:
        fixed += '"'

    fixed = re.sub(r'"\s*\n\s*"', '",\n"', fixed)
    fixed = re.sub(r"}\s*\n\s*{", "},\n{", fixed)
    fixed = re.sub(r']\s*\n\s*"', '],\n"', fixed)
    fixed = re.sub(r",\s*([}\]])", r"\1", fixed)

    try:
        return json.loads(fixed)
    except json.JSONDecodeError:
        pass

    trimmed = fixed.rstrip()
    last_comma = max(trimmed.rfind(","), trimmed.rfind("{"), trimmed.rfind("["))
    if last_comma > len(trimmed) - 180 and last_comma > 0 and trimmed[last_comma] == ",":
        trimmed = trimmed[:last_comma]

    stack: list[str] = []
    in_str = False
    esc = False
    for ch in trimmed:
        if esc:
            esc = False
            continue
        if ch == "\\":
            esc = True
            continue
        if ch == '"':
            in_str = not in_str
            continue
        if not in_str:
            if ch in ("{", "["):
                stack.append("}" if ch == "{" else "]")
            elif ch in ("}", "]") and stack:
                stack.pop()

    if in_str:
        trimmed += '"'
    trimmed = re.sub(r",\s*$", "", trimmed)
    repaired = trimmed + "".join(reversed(stack))
    repaired = re.sub(r",\s*([}\]])", r"\1", repaired)
    return json.loads(repaired)


def extract_clean_json(raw_text: str) -> dict:
    if not raw_text:
        raise ValueError("Пустой ответ от модели")

    text = re.sub(r"<think>.*?</think>", "", raw_text, flags=re.DOTALL).strip()
    if "</think>" in text:
        text = text.split("</think>")[-1].strip()

    text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\s*```$", "", text)

    start_idx = text.find("{")
    if start_idx == -1:
        raise ValueError("JSON-объект {...} не найден в ответе")

    end_idx = text.rfind("}")
    candidate = text[start_idx : end_idx + 1] if end_idx > start_idx else text[start_idx:]

    try:
        return json.loads(candidate)
    except json.JSONDecodeError:
        return repair_malformed_json(candidate)


def _run_with_hard_timeout(fn, timeout_sec: int, model_label: str):
    pool = ThreadPoolExecutor(max_workers=1)
    future = pool.submit(fn)
    try:
        return future.result(timeout=timeout_sec)
    except FuturesTimeoutError:
        pool.shutdown(wait=False, cancel_futures=True)
        raise RuntimeError(f"TIMEOUT: {model_label} превысила лимит {timeout_sec} сек")
    except Exception:
        pool.shutdown(wait=False, cancel_futures=True)
        raise
    finally:
        pool.shutdown(wait=False)


def _openrouter_chat_worker(
    model_name: str,
    system_prompt: str,
    user_prompt: str,
    timeout_sec: int,
    reasoning_mode: ReasoningMode = "off",
    max_tokens: int = 12000,
    temperature: float = 0.55,
) -> str:
    url = "https://openrouter.ai/api/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json",
        "HTTP-Referer": "https://dzen.ru/algomoney_ru",
        "X-Title": "AlgoMoney VPS Pipeline",
    }
    base_payload: dict = {
        "model": model_name,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        "temperature": temperature,
        "max_tokens": max_tokens,
    }

    payload = dict(base_payload)
    if reasoning_mode == "off":
        payload["reasoning"] = {"enabled": False}
    elif reasoning_mode in ("low", "medium", "high"):
        payload["reasoning"] = {"enabled": True, "effort": reasoning_mode}

    read_timeout = max(12, timeout_sec - 3)
    resp = HTTP.post(url, headers=headers, json=payload, timeout=(8, read_timeout))
    if resp.status_code == 400 and "reasoning" in payload:
        resp = HTTP.post(url, headers=headers, json=base_payload, timeout=(8, read_timeout))

    if resp.status_code != 200:
        raise RuntimeError(f"HTTP {resp.status_code}: {resp.text[:160]}")

    body = resp.json()
    if "error" in body:
        raise RuntimeError(f"OpenRouter Error: {body['error']}")

    choices = body.get("choices", [])
    if not choices:
        raise RuntimeError("Пустой choices")

    content = (choices[0].get("message", {}).get("content") or "").strip()
    if not content:
        finish_reason = choices[0].get("finish_reason", "unknown")
        raise RuntimeError(f"Пустой content от {model_name} (finish_reason={finish_reason})")
    return content


def invoke_llm_json(
    candidates: list[ModelSpec],
    stage_title: str,
    system_prompt: str,
    user_prompt: str,
    required_keys: list[str],
    validator=None,
    max_tokens: int = 12000,
) -> dict:
    """Диспетчер вызовов OpenRouter: сбойная модель исключается с 1-го раза до конца запуска."""
    if not has_valid_openrouter_key():
        raise RuntimeError("OPENROUTER_API_KEY не задан или некорректен в .env")

    last_error = None
    for idx, (model_name, reasoning_mode) in enumerate(candidates, start=1):
        skip_key = (model_name, reasoning_mode)
        if not model_name or skip_key in SESSION_SKIP_KEYS:
            continue

        timeout_sec = get_model_timeout(model_name, default_sec=70)
        mode_tag = f" [reasoning={reasoning_mode}]" if reasoning_mode else ""

        try:
            log.info(
                "[%s] Запрос к %s%s (%d/%d, таймаут %d сек)...",
                stage_title,
                model_name,
                mode_tag,
                idx,
                len(candidates),
                timeout_sec,
            )
            raw_text = _run_with_hard_timeout(
                lambda: _openrouter_chat_worker(
                    model_name,
                    system_prompt,
                    user_prompt,
                    timeout_sec=timeout_sec,
                    reasoning_mode=reasoning_mode,
                    max_tokens=max_tokens,
                ),
                timeout_sec=timeout_sec,
                model_label=model_name,
            )

            parsed = extract_clean_json(raw_text)
            if validator:
                parsed = validator(parsed) or parsed
            missing = [k for k in required_keys if k not in parsed]
            if missing:
                raise ValueError(f"Отсутствуют обязательные поля JSON: {missing}")

            log.info("  -> Успех: %s%s (%d симв.)", model_name, mode_tag, len(raw_text))
            return parsed

        except Exception as e:
            last_error = e
            err_str = str(e)
            SESSION_SKIP_KEYS.add(skip_key)
            if any(code in err_str for code in ("HTTP 400", "HTTP 402", "HTTP 403", "HTTP 404", "HTTP 429")):
                for mode in (None, "off", "low"):
                    SESSION_SKIP_KEYS.add((model_name, mode))

            log.warning("Модель %s%s исключена (%s). Переключение...", model_name, mode_tag, err_str[:90])
            time.sleep(0.2)

    raise RuntimeError(f"Этап [{stage_title}] не выполнен ни одной моделью: {last_error}")


# ======================================================================
# ЭТАП 1: ФИЛЬТРАЦИЯ ТЕМ (PARTICLE-1 / DOTS3 / ERNIE)
# ======================================================================
def filter_topics(rss_items: list[dict]) -> list[dict]:
    if not rss_items:
        return []

    candidates = build_queue(
        [
            (MODEL_PARTICLE, "off"),
            (MODEL_DOTS3, "off"),
            (MODEL_ERNIE, "off"),
        ],
        tail_limit=2,
    )

    batch = rss_items[:15]
    headlines_text = "\n".join(
        f"{i}. [{item['source']}] {item['title']}" for i, item in enumerate(batch)
    )

    sys_prompt = (
        "Ты — фильтр новостей для серьёзного финансово-аналитического издания «Алгоритм Денег». "
        "Оцени надёжность источника и пригодность новости для финансового разбора с расчётами в рублях. "
        "Отсеивай кликбейт, светскую хронику и воду. Отвечай СТРОГО валидным JSON."
    )
    user_prompt = f"""
Оцени каждый заголовок по шкале от 0.0 до 1.0 (вероятность того, что тема надёжна и подходит для серьёзного финансового анализа):

{headlines_text}

Верни СТРОГО JSON:
{{
  "scored": [
    {{"id": 0, "score": 0.92, "reliable": true}},
    {{"id": 1, "score": 0.25, "reliable": false}}
  ]
}}
"""
    try:
        res = invoke_llm_json(
            candidates=candidates,
            stage_title="1/6 Фильтр тем",
            system_prompt=sys_prompt,
            user_prompt=user_prompt,
            required_keys=["scored"],
            max_tokens=4096,
        )
        approved: list[tuple[float, dict]] = []
        for entry in res.get("scored", []):
            idx = entry.get("id")
            score = float(entry.get("score", 0))
            reliable = bool(entry.get("reliable", score >= 0.65))
            if isinstance(idx, int) and 0 <= idx < len(batch) and reliable and score >= 0.65:
                approved.append((score, batch[idx]))

        approved.sort(key=lambda x: x[0], reverse=True)
        if approved:
            vetted = [item for _, item in approved[:10]]
            log.info(
                "[✓] Фильтр отобрал %d тем из %d (топ: «%s»)",
                len(vetted),
                len(batch),
                vetted[0]["title"][:70],
            )
            return vetted
    except Exception as e:
        log.warning("Фильтр тем пропущен (%s), используем исходный список RSS.", str(e)[:75])

    return batch[:10]


# ======================================================================
# ЭТАП 2: ЛОКАЛЬНЫЙ RAG (TF-IDF + ЧИСЛОВОЙ РЕРАНК ПО RSS И KNOWLEDGE_BASE)
# ======================================================================
def _load_knowledge_base_chunks() -> list[dict]:
    chunks: list[dict] = []
    KNOWLEDGE_BASE_DIR.mkdir(parents=True, exist_ok=True)
    for ext in ("*.txt", "*.md"):
        for file_path in sorted(KNOWLEDGE_BASE_DIR.glob(ext)):
            try:
                text = file_path.read_text(encoding="utf-8").strip()
                if not text:
                    continue
                for para in re.split(r"\n\s*\n", text):
                    clean_p = para.strip()
                    if len(clean_p) >= 40:
                        chunks.append({"source": f"База знаний ({file_path.name})", "text": clean_p[:900]})
            except Exception as e:
                log.warning("Ошибка чтения %s: %s", file_path.name, e)
    return chunks


def _tokenize_ru_en(text: str) -> list[str]:
    return re.findall(r"[а-яА-Яa-zA-Z0-9]{3,}", text.lower())


def retrieve_rag_context(topic_query: str, rss_items: list[dict]) -> str:
    corpus: list[dict] = []
    for item in rss_items:
        full_entry = f"{item['title']}. {item['description']}".strip()
        if len(full_entry) >= 35:
            corpus.append({"source": item["source"], "text": full_entry[:800]})
    corpus.extend(_load_knowledge_base_chunks())

    if not corpus:
        return "Дополнительная локальная фактура отсутствует; опирайся на проверенные факты из брифа."

    docs_tokens = [_tokenize_ru_en(c["text"]) for c in corpus]
    q_tokens = _tokenize_ru_en(topic_query)
    if not q_tokens:
        return "\n".join(f"- [{c['source']}] {c['text']}" for c in corpus[:5])

    df: Counter = Counter()
    for tokens in docs_tokens:
        for tok in set(tokens):
            df[tok] += 1
    n_docs = len(docs_tokens)

    def to_tfidf(tokens: list[str]) -> dict[str, float]:
        tf = Counter(tokens)
        return {
            tok: (count / max(1, len(tokens))) * (math.log((n_docs + 1) / (df.get(tok, 0) + 1)) + 1.0)
            for tok, count in tf.items()
        }

    q_vec = to_tfidf(q_tokens)
    q_norm = math.sqrt(sum(v * v for v in q_vec.values())) or 1.0
    q_terms = set(q_tokens)

    scored: list[tuple[float, dict]] = []
    for chunk, tokens in zip(corpus, docs_tokens):
        d_vec = to_tfidf(tokens)
        d_norm = math.sqrt(sum(v * v for v in d_vec.values())) or 1.0
        sim = sum(q_vec[t] * d_vec.get(t, 0.0) for t in q_vec) / (q_norm * d_norm)
        num_bonus = 0.25 if re.search(r"\d+(?:[.,]\d+)?\s*(?:%|млрд|млн|руб|₽)", chunk["text"]) else 0.0
        exact_bonus = 0.05 * len(q_terms & set(tokens))
        scored.append((sim + num_bonus + exact_bonus, chunk))

    scored.sort(key=lambda x: x[0], reverse=True)
    top_5 = [c for _, c in scored[:5]]
    log.info("[✓] RAG отобрал %d опорных фрагментов (из %d документов).", len(top_5), len(corpus))
    return "\n".join(f"- [{c['source']}] {c['text']}" for c in top_5)


# ======================================================================
# ГАРАНТИРОВАННОЕ АВТО-ФОРМАТИРОВАНИЕ (ЖИРНЫЙ, КУРСИВ, ВРЕЗКИ)
# ======================================================================
BOLD_PATTERNS = [
    r"(?<!\*)\b(\d+(?:[\s.,]\d+)*\s*(?:%|процент(?:а|ов)?|млрд|млн|тыс\.?|руб(?:лей|ля|\.)?|₽|\$|€|год(?:а|лет)?|месяц(?:а|ев)?|п\.п\.))\b(?!\*)",
    r"(?<!\*)\b(ключев(?:ая|ой|ую)\s+ставк[а-я]+|инфляци[яиюей]+|дивиденд[а-я]*|чист(?:ая|ой)\s+прибыл[а-я]+|сложн(?:ый|ого)\s+процент[а-я]*|ндфл|фнс|цб\s+рф|офз|иис)\b(?!\*)",
]

ITALIC_PHRASES = [
    "важно понимать",
    "главный нюанс",
    "обратите внимание",
    "на практике",
    "простыми словами",
    "реальная доходность",
    "покупательская способность",
    "подушка безопасности",
    "точка безубыточности",
    "диверсификация портфеля",
    "эффективная ставка",
    "налоговый вычет",
    "капитализация процентов",
    "долговая нагрузка",
    "инвестиционный горизонт",
    "ликвидность активов",
    "налоговая база",
    "не является инвестиционной рекомендацией",
]


def _clean_unclosed_markdown(text: str) -> str:
    if not text:
        return ""
    if text.count("**") % 2 != 0:
        text = text.replace("**", "")
    if len(re.findall(r"(?<!\*)\*(?!\*)", text)) % 2 != 0:
        text = re.sub(r"(?<!\*)\*(?!\*)", "", text)
    return text.strip()


def ensure_rich_formatting(text: str) -> str:
    """Гарантирует наличие **жирного** (на цифрах/терминах) и *курсива* (на акцентах) в абзаце."""
    if not text or not isinstance(text, str):
        return ""

    text = _clean_unclosed_markdown(text)
    text = re.sub(r"(?<![\w*])_([^_\n]+?)_(?![\w*])", r"*\1*", text)

    if "**" not in text:
        for pat in BOLD_PATTERNS:
            text = re.sub(pat, r"**\1**", text, count=2, flags=re.IGNORECASE)

    text_no_bold = re.sub(r"\*\*.*?\*\*", "", text)
    if "*" not in text_no_bold:
        applied_italic = False
        for phrase in ITALIC_PHRASES:
            pattern = re.compile(rf"(?<!\*)\b({re.escape(phrase)})\b(?!\*)", re.IGNORECASE)
            if pattern.search(text):
                text = pattern.sub(r"*\1*", text, count=1)
                applied_italic = True
                break

        if not applied_italic:
            sentences = re.split(r"(?<=[.!?])\s+", text)
            if sentences and "," in sentences[0] and "*" not in sentences[0]:
                first_part, rest_part = sentences[0].split(",", 1)
                if 12 < len(first_part) < 62:
                    sentences[0] = f"*{first_part}*,{rest_part}"
                    text = " ".join(sentences)

    return _clean_unclosed_markdown(text)


def apply_formatting_to_article(data: dict) -> dict:
    """Применяет авто-форматирование ко всем текстовым полям статьи и ТГ-поста."""
    data["article_lead"] = ensure_rich_formatting(data.get("article_lead", ""))
    data["call_to_action"] = ensure_rich_formatting(data.get("call_to_action", ""))
    if data.get("tg_public_post"):
        data["tg_public_post"] = ensure_rich_formatting(data["tg_public_post"])

    for sec in data.get("sections", []):
        sec["paragraphs"] = [ensure_rich_formatting(p) for p in sec.get("paragraphs", []) if p.strip()]
        if sec.get("calculation_example"):
            sec["calculation_example"] = ensure_rich_formatting(sec["calculation_example"])
        sec["bullet_points"] = [ensure_rich_formatting(b) for b in sec.get("bullet_points", []) if b.strip()]
        if sec.get("smart_note"):
            sec["smart_note"] = _clean_unclosed_markdown(sec["smart_note"])
    return data


# ======================================================================
# ЭТАПЫ 3–6: РЕДАКЦИОННЫЙ КОНВЕЙЕР
# ======================================================================
def calculate_dzen_article_length(data: dict) -> int:
    total = len(data.get("article_title", "")) + len(data.get("article_lead", ""))
    total += len(data.get("call_to_action", ""))
    for sec in data.get("sections", []):
        total += len(sec.get("subtitle", ""))
        total += sum(len(p) for p in sec.get("paragraphs", []))
        total += len(sec.get("calculation_example", ""))
        total += sum(len(b) for b in sec.get("bullet_points", []))
        total += len(sec.get("smart_note", ""))
    return total


def normalize_article_draft(parsed: dict) -> dict:
    secs = parsed.get("sections")
    if not isinstance(secs, list) or len(secs) < 2:
        raise ValueError(f"Недостаточно разделов в sections: {len(secs) if isinstance(secs, list) else 0}")

    if len(secs) == 2:
        longest_idx = max(range(2), key=lambda i: len(" ".join(secs[i].get("paragraphs", []))))
        big_sec = secs[longest_idx]
        paras = big_sec.get("paragraphs", [])
        if len(paras) >= 2:
            split_sec = {
                "subtitle": big_sec.get("smart_note") or "Практические выводы и алгоритм действий",
                "paragraphs": paras[1:],
                "calculation_example": big_sec.get("calculation_example", ""),
                "bullet_points": big_sec.get("bullet_points", []),
                "smart_note": "",
            }
            big_sec["paragraphs"] = paras[:1]
            big_sec["calculation_example"] = ""
            secs.insert(longest_idx + 1, split_sec)

    parsed["sections"] = secs[:REQUIRED_SECTIONS_COUNT]

    for idx, sec in enumerate(parsed["sections"], start=1):
        if not isinstance(sec.get("paragraphs"), list) or not sec["paragraphs"]:
            fallback_p = sec.get("text") or sec.get("content") or sec.get("summary") or ""
            sec["paragraphs"] = [fallback_p] if fallback_p else ["Аналитический разбор ключевых показателей."]
        sec.setdefault("subtitle", f"Ключевой фактор №{idx}")
        sec.setdefault("calculation_example", "")
        sec.setdefault("bullet_points", [])
        sec.setdefault("smart_note", "")

    if not str(parsed.get("article_lead", "")).strip():
        parsed["article_lead"] = parsed["sections"][0]["paragraphs"][0][:360]

    if not str(parsed.get("call_to_action", "")).strip():
        parsed["call_to_action"] = (
            "Грамотный расчёт цифр и учёт налоговых правил позволяют сохранить капитал даже в период волатильности. "
            "А как вы сейчас распределяете свои сбережения? Делитесь мнением в комментариях."
        )

    parsed.setdefault(
        "cover_prompt_en",
        "Modern 3D editorial financial illustration, banking vault and glowing chart, deep navy and gold aesthetic, clean studio lighting, no text, no letters, 8k",
    )
    parsed.setdefault("cover_search_ru", parsed.get("article_title", "экономика финансы деньги")[:55])
    return parsed


def md_to_tg_html(text: str) -> str:
    """Конвертирует Markdown в валидный Telegram HTML."""
    if not text:
        return ""
    escaped = html.escape(_clean_unclosed_markdown(text), quote=False)
    escaped = re.sub(r"\*\*\*(.+?)\*\*\*", r"<b><i>\1</i></b>", escaped, flags=re.DOTALL)
    escaped = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", escaped, flags=re.DOTALL)
    escaped = re.sub(r"`([^`\n]+?)`", r"<code>\1</code>", escaped)
    escaped = re.sub(r"(?<!\*)\*([^\*\n]+?)\*(?!\*)", r"<i>\1</i>", escaped)
    escaped = re.sub(r"(?<![\w])_([^_\n]+?)_(?![\w])", r"<i>\1</i>", escaped)
    return escaped


def build_single_dzen_telegram_post(data: dict, cover_url: str = "") -> str:
    """
    Собирает структурированную статью для ОДНОГО сообщения Telegram (<= 4 050 HTML-симв.)
    с обложкой через скрытый link_preview, жирными акцентами, курсивом и <blockquote>-врезками.
    """
    sections = data.get("sections", [])[:REQUIRED_SECTIONS_COUNT]
    safe_cover_url = html.escape(cover_url, quote=True) if cover_url else ""
    hidden_cover = f'<a href="{safe_cover_url}">‍</a>' if safe_cover_url else ""
    clean_title = data["article_title"].strip().rstrip(".")

    header = (
        f"{hidden_cover}"
        f"<b>{md_to_tg_html(clean_title)}.</b>\n\n"
        f"{md_to_tg_html(data['article_lead'].strip())}\n\n"
    )
    cta = f"<b>▪ Что в итоге?</b>\n{md_to_tg_html(data.get('call_to_action', '').strip())}\n\n"
    tags = " ".join(
        f"#{re.sub(r'[^а-яА-Яa-zA-Z0-9_]', '', t.replace(' ', '_'))}"
        for t in data.get("meta_tags", [])[:5]
    )
    footer = cta + tags.strip()

    def render_body(trim_factor: float = 1.0, keep_notes: bool = True, max_bullets: int = 2) -> str:
        blocks: list[str] = []
        for idx, sec in enumerate(sections, start=1):
            sub = re.sub(r"^\d+[\.\)]\s*", "", sec.get("subtitle", "")).strip()
            b_txt = f"<b>▪ {idx}. {md_to_tg_html(sub)}</b>\n\n"

            paras = [p.strip() for p in sec.get("paragraphs", []) if p.strip()]
            for p_idx, p in enumerate(paras):
                if trim_factor < 1.0 and p_idx > 0:
                    sentences = re.split(r"(?<=[.!?])\s+", p)
                    keep_n = max(1, int(len(sentences) * trim_factor))
                    p = " ".join(sentences[:keep_n])
                b_txt += md_to_tg_html(p) + "\n\n"

            calc_ex = sec.get("calculation_example", "").strip()
            if calc_ex:
                b_txt += f"<blockquote>🧮 <b>Расчёт:</b> {md_to_tg_html(calc_ex)}</blockquote>\n\n"

            bullets = [b.strip() for b in sec.get("bullet_points", []) if b.strip()][:max_bullets]
            for b in bullets:
                b_txt += f"• {md_to_tg_html(b)}\n"
            if bullets:
                b_txt += "\n"

            smart_note = sec.get("smart_note", "").strip()
            if smart_note and keep_notes:
                b_txt += f"<blockquote>📊 <i>{md_to_tg_html(smart_note)}</i></blockquote>\n\n"

            blocks.append(b_txt)
        return "".join(blocks)

    profiles = [
        (1.0, True, 2),
        (0.85, True, 2),
        (0.70, True, 2),
        (0.60, False, 2),
        (0.50, False, 1),
        (0.35, False, 0),
    ]
    for factor, keep_notes, max_b in profiles:
        candidate = header + render_body(factor, keep_notes, max_b) + footer
        if len(candidate) <= TG_SINGLE_MSG_HTML_LIMIT:
            return candidate

    allowed_body = max(200, TG_SINGLE_MSG_HTML_LIMIT - len(header) - len(footer))
    body_paragraphs = render_body(0.35, False, 0).split("\n\n")
    safe_body_parts: list[str] = []
    curr_len = 0
    for part in body_paragraphs:
        if curr_len + len(part) + 2 > allowed_body:
            break
        safe_body_parts.append(part)
        curr_len += len(part) + 2

    return header + "\n\n".join(safe_body_parts) + "\n\n" + footer


def enforce_strict_article_bounds(data: dict) -> dict:
    data = normalize_article_draft(data)
    data = apply_formatting_to_article(data)
    placeholder_url = "https://iili.io/placeholder_cover_image.jpg"

    for _ in range(18):
        if len(build_single_dzen_telegram_post(data, cover_url=placeholder_url)) <= TG_SINGLE_MSG_HTML_LIMIT - 40:
            break
        longest_sec = None
        longest_p_idx = -1
        max_p_len = 0
        for sec in data["sections"]:
            for idx, p in enumerate(sec.get("paragraphs", [])):
                if len(p) > max_p_len and len(p) > 160:
                    max_p_len = len(p)
                    longest_sec = sec
                    longest_p_idx = idx
        if longest_sec is None:
            break
        sentences = re.split(r"(?<=[.!?])\s+", longest_sec["paragraphs"][longest_p_idx])
        if len(sentences) > 1:
            longest_sec["paragraphs"][longest_p_idx] = _clean_unclosed_markdown(" ".join(sentences[:-1]))
        else:
            trimmed = longest_sec["paragraphs"][longest_p_idx][: max_p_len - 100].rsplit(" ", 1)[0] + "."
            longest_sec["paragraphs"][longest_p_idx] = _clean_unclosed_markdown(trimmed)
    return data


def run_analyst_stage(vetted_news: list[dict], used_topics: list[str]) -> dict:
    candidates = build_queue(
        [
            (MODEL_DOTS3, "off"),
            (MODEL_PARTICLE, "off"),
            (MODEL_ERNIE, "off"),
        ],
        tail_limit=3,
    )

    news_list_str = "\n".join(
        f"- [{n['source']}] {n['title']}: {n['description'][:220]}" for n in vetted_news[:8]
    )
    used_str = "\n".join(f"- {u}" for u in used_topics[-30:]) if used_topics else "Пока пусто."

    sys_prompt = (
        f"Ты — главный финансовый стратег издания «{CHANNEL_NAME}». "
        "Выбери одну главную тему дня и спроектируй аналитический план из РОВНО 3 разделов "
        "с расчётами в рублях и процентах. Отвечай СТРОГО валидным JSON. "
        'Внутри русского текста используй только кавычки «елочки», никогда не ставь двойные кавычки " внутри строк.'
    )
    user_prompt = f"""
Выбери ОДНУ самую значимую для кошелька россиян тему из отобранных новостей и составь план статьи на **ровно 3 раздела** (целевой объём {ARTICLE_MIN_CHARS} – {ARTICLE_TARGET_MAX_CHARS} символов).

ОТОБРАННЫЕ НОВОСТИ:
{news_list_str}

УЖЕ ИСПОЛЬЗОВАННЫЕ ТЕМЫ (НЕ ПОВТОРЯТЬ):
{used_str}

Верни СТРОГО JSON:
{{
  "chosen_topic": "Краткая формулировка выбранной темы (1 строка)",
  "rag_search_query": "Ключевые слова для поиска цифр, ставок и фактов по этой теме (5-8 слов)",
  "why_actual": "Почему тема критически важна прямо сейчас",
  "sections_plan": [
    {{"title": "Логический блок 1 (Суть и цифры)", "focus": "Что произошло и какие цифры важны"}},
    {{"title": "Логический блок 2 (Расчёт выгоды/потерь)", "focus": "Конкретный пример в рублях и процентах"}},
    {{"title": "Логический блок 3 (Риски, налоги и алгоритм действий)", "focus": "Что делать инвестору/вкладчику"}}
  ]
}}
"""
    try:
        return invoke_llm_json(
            candidates=candidates,
            stage_title="3/6 Аналитик (Dots3 / Particle)",
            system_prompt=sys_prompt,
            user_prompt=user_prompt,
            required_keys=["chosen_topic", "rag_search_query", "sections_plan"],
            max_tokens=6000,
        )
    except Exception as e:
        top_item = vetted_news[0]
        log.warning("Аналитик переведён на авто-план (%s): «%s»", str(e)[:70], top_item["title"][:65])
        return {
            "chosen_topic": top_item["title"],
            "rag_search_query": top_item["title"],
            "why_actual": f"Актуальная экономическая повестка ({top_item['source']}): {top_item['description'][:180]}",
            "sections_plan": [
                {"title": "Что произошло и какие цифры стоят за новостью", "focus": top_item["description"][:200]},
                {"title": "Расчёт выгоды и потерь в рублях для частного капитала", "focus": "Сценарии для сумм 100 000 ₽ и 500 000 ₽"},
                {"title": "Скрытые риски, налоги и алгоритм действий", "focus": "Практические шаги для защиты сбережений"},
            ],
        }


def run_writer_stage(analyst_plan: dict, rag_facts: str) -> dict:
    candidates = build_queue(
        [
            (MODEL_NEMOTRON_ULTRA, "off"),
            (MODEL_DOTS3, "off"),
            (MODEL_PARTICLE, "off"),
            (MODEL_ERNIE, "off"),
        ],
        tail_limit=3,
    )

    sys_prompt = (
        f"Ты — ведущий финансовый журналист «{CHANNEL_NAME}». Концепция: {CHANNEL_THEME} "
        "Пиши предметно, живым русским языком, опираясь строго на RAG-фактуру и помечая расчётные сценарии как модели (без гарантий доходности). "
        "Обязательно используй Markdown-выделение **жирным** для цифр и выводов и *курсивом* для терминов. "
        "Отвечай СТРОГО валидным JSON. Внутри русского текста используй ТОЛЬКО кавычки-елочки «...»."
    )
    user_prompt = f"""
Напиши аналитическую статью для Яндекс Дзена за ОДИН проход по плану Аналитика и фактам из RAG.

ПЛАН АНАЛИТИКА:
{json.dumps(analyst_plan, ensure_ascii=False)}

ПРОВЕРЕННАЯ ФАКТУРА (RAG):
{rag_facts}

ЖЁСТКИЕ ТРЕБОВАНИЯ:
1. **Целевой объём всей статьи: строго {ARTICLE_MIN_CHARS} – {ARTICLE_TARGET_MAX_CHARS} символов** (чтобы статья целиком вошла в 1 сообщение Telegram без обрезки!).
2. **Структура: РОВНО 3 раздела в массиве `sections`** (подзаголовки `subtitle` без номера в начале).
  - В каждом разделе:
    - `paragraphs`: **2 плотных абзаца** по 3 предложения с **жирным выделением цифр и выводов** (`**...**`) и *курсивом для терминов* (`*...*`).
    - `calculation_example`: модельный математический расчёт в рублях и процентах (без гарантий будущей доходности).
    - `bullet_points`: **2 конкретных тезиса**.
    - `smart_note`: краткий вывод от «Алгоритма Денег» (1 предложение).
3. **Обложка (`cover_prompt_en` и `cover_search_ru`):**
  - `cover_prompt_en`: английский промпт (12–15 слов) для 3D/editorial иллюстрации по теме статьи (СТРОГО без текста, букв, цифр и надписей!).
  - `cover_search_ru`: 3–4 русских слова для поиска фото по теме.

Верни СТРОГО JSON:
{{
  "chosen_topic": "{analyst_plan.get('chosen_topic', '')}",
  "why_actual": "{analyst_plan.get('why_actual', '')}",
  "article_title": "Серьёзный аналитический заголовок без кликбейта (до 95 символов)",
  "cover_prompt_en": "Modern 3D editorial financial illustration matching main topic, deep navy and gold, studio lighting, no text, no letters, no numbers, 8k",
  "cover_search_ru": "запрос из 3-4 русских слов по теме статьи",
  "meta_tags": ["финансы", "экономика", "инвестиции", "деньги", "налоги"],
  "article_lead": "Сильное вступление статьи (3 предложения, ~280 символов) с **жирным** и *курсивом*.",
  "sections": [
    {{
      "subtitle": "Подзаголовок раздела",
      "paragraphs": [
        "Первый абзац (3 предложения) с **ключевыми цифрами** и *важным нюансом*.",
        "Второй абзац (3 предложения) с разбором механики и примером."
      ],
      "calculation_example": "Модельный расчёт в рублях и процентах с **жирным итогом** и *курсивом для условий*.",
      "bullet_points": [
        "**Фактор 1:** конкретика и цифры",
        "**Фактор 2:** практический вывод"
      ],
      "smart_note": "Короткий итог раздела."
    }}
  ],
  "call_to_action": "Итоговый вывод статьи (2 предложения) и нейтральный вопрос читателям"
}}
"""
    return invoke_llm_json(
        candidates=candidates,
        stage_title="4/6 Автор (Nemotron 3 Ultra / Dots3)",
        system_prompt=sys_prompt,
        user_prompt=user_prompt,
        required_keys=["article_title", "article_lead", "sections", "call_to_action"],
        validator=normalize_article_draft,
        max_tokens=12000,
    )


def run_editor_and_tg_stage(draft: dict) -> dict:
    candidates = build_queue(
        [
            (MODEL_PARTICLE, "off"),
            (MODEL_DOTS3, "off"),
            (MODEL_NEMOTRON_ULTRA, "off"),
        ],
        tail_limit=2,
    )

    sys_prompt = (
        "Ты — шеф-редактор финансового издания «Алгоритм Денег». "
        "Убери канцелярит и ИИ-штампы, ОБЯЗАТЕЛЬНО сохрани и расставь **жирный шрифт** (цифры, ставки, главные выводы) "
        "и *курсив* (термины, оговорки), проконтролируй объём статьи строго "
        f"{ARTICLE_MIN_CHARS} – {ARTICLE_TARGET_MAX_CHARS} символов, "
        "и напиши отдельный короткий пост для Telegram (700–850 символов) с **жирным** и *курсивом*, который даёт отдельный инсайт. "
        "Отвечай СТРОГО валидным JSON. Используй только кавычки «елочки»."
    )
    user_prompt = f"""
Отредактируй черновик статьи (сохранив РОВНО 3 раздела, **жирные выделения**, *курсив* и объём {ARTICLE_MIN_CHARS} – {ARTICLE_TARGET_MAX_CHARS} символов) и напиши уникальный пост для Telegram (`tg_public_post`, 700–850 символов):

ЧЕРНОВИК СТАТЬИ:
{json.dumps(draft, ensure_ascii=False)}

Верни СТРОГО JSON с обновлёнными полями `article_title`, `article_lead`, `sections` (ровно 3 объекта), `call_to_action` и новым полем `tg_public_post`.
"""
    try:
        edited = invoke_llm_json(
            candidates=candidates,
            stage_title="5/6 Редактор и ТГ-пост (Particle / Dots3)",
            system_prompt=sys_prompt,
            user_prompt=user_prompt,
            required_keys=["article_title", "article_lead", "sections", "call_to_action", "tg_public_post"],
            validator=normalize_article_draft,
            max_tokens=12000,
        )
        edited["chosen_topic"] = draft.get("chosen_topic", edited.get("article_title", ""))
        edited["why_actual"] = draft.get("why_actual", "")
        edited["cover_prompt_en"] = draft.get("cover_prompt_en", "")
        edited["cover_search_ru"] = draft.get("cover_search_ru", "")
        edited["meta_tags"] = draft.get("meta_tags", ["финансы", "экономика", "инвестиции", "деньги"])
        return enforce_strict_article_bounds(edited)
    except Exception as e:
        log.warning("ИИ-редактура пропущена (%s), применяем алгоритмическую нормализацию.", str(e)[:75])
        if not draft.get("tg_public_post"):
            sec1_calc = draft["sections"][0].get("calculation_example", "") if draft.get("sections") else ""
            draft["tg_public_post"] = (
                f"**{draft['article_title']}**\n\n"
                f"{draft['article_lead'][:400]}\n\n"
                f"{sec1_calc[:220]}\n\n"
                "*Подробные расчёты в рублях и пошаговый алгоритм* — в **новой статье на Дзене** 👇"
            ).strip()
        return enforce_strict_article_bounds(draft)


def run_moderation_check(data: dict) -> dict:
    candidates = build_queue(
        [
            (MODEL_PARTICLE, "off"),
            (MODEL_DOTS3, "off"),
        ],
        tail_limit=2,
    )

    full_article_preview = (
        f"Заголовок: {data['article_title']}\n"
        f"Лид: {data['article_lead']}\n"
        + "\n".join(
            f"Раздел: {s['subtitle']} | {' '.join(s.get('paragraphs', []))}"
            for s in data.get("sections", [])
        )
        + f"\nИтог: {data.get('call_to_action', '')}"
    )

    sys_prompt = (
        "Ты — модератор контента и антикликбейт-фильтр для Яндекс Дзена и Telegram. "
        "Проверь текст на стоп-паттерны Дзена: жёлтый кликбейт, необоснованные гарантии доходности, токсичность. "
        "Отвечай СТРОГО валидным JSON."
    )
    user_prompt = f"""
Проверь финальный заголовок, статью и пост для Telegram:

СТАТЬЯ ДЛЯ ДЗЕНА:
{full_article_preview[:4200]}

ПОСТ ДЛЯ TELEGRAM:
{data.get('tg_public_post', '')}

Верни СТРОГО JSON:
{{
  "flagged": false,
  "issues": [],
  "fix_instructions": ""
}}
"""
    try:
        verdict = invoke_llm_json(
            candidates=candidates,
            stage_title="6/6 Модерация (Particle / Dots3)",
            system_prompt=sys_prompt,
            user_prompt=user_prompt,
            required_keys=["flagged"],
            max_tokens=2048,
        )
        if not verdict.get("flagged"):
            log.info("[✓] Модерация пройдена: стоп-паттерны и кликбейт не обнаружены.")
            return enforce_strict_article_bounds(data)

        issues = verdict.get("issues", [])
        fix_hint = verdict.get("fix_instructions", "Убери кликбейт и смягчи формулировки.")
        log.warning("[!] Обнаружены стоп-паттерны (%d шт.) — запускаем точечную правку...", len(issues))

        fix_prompt = f"""
Исправь замечания модерации ({issues}; рекомендация: {fix_hint}) в заголовке, лиде, выводах и ТГ-посте, ОБЯЗАТЕЛЬНО сохранив деловой тон, **жирный шрифт** (`**...**`) и *курсив* (`*...*`).

Текущие данные:
{{
  "article_title": {json.dumps(data["article_title"], ensure_ascii=False)},
  "article_lead": {json.dumps(data["article_lead"], ensure_ascii=False)},
  "call_to_action": {json.dumps(data["call_to_action"], ensure_ascii=False)},
  "tg_public_post": {json.dumps(data["tg_public_post"], ensure_ascii=False)}
}}

Верни СТРОГО JSON с этими 4 исправленными полями.
"""
        fixed_fields = invoke_llm_json(
            candidates=candidates,
            stage_title="6.1/6 Точечная правка",
            system_prompt="Ты — редактор комплаенса. Переформулируй проблемные фразы чисто, сохраняя **жирный** и *курсив*. Отвечай СТРОГО JSON.",
            user_prompt=fix_prompt,
            required_keys=["article_title", "article_lead", "call_to_action", "tg_public_post"],
            max_tokens=4096,
        )
        for k in ("article_title", "article_lead", "call_to_action", "tg_public_post"):
            if isinstance(fixed_fields.get(k), str) and len(fixed_fields[k].strip()) > 15:
                data[k] = fixed_fields[k].strip()
        log.info("[✓] Текст успешно очищен от стоп-паттернов.")
    except Exception as e:
        log.info("Проверка модерации пропущена (%s).", str(e)[:65])

    return enforce_strict_article_bounds(data)


def generate_standalone_tg_news(vetted_news: list[dict], used_topics: list[str]) -> dict:
    candidates = build_queue(
        [
            (MODEL_PARTICLE, "off"),
            (MODEL_DOTS3, "off"),
            (MODEL_ERNIE, "off"),
        ],
        tail_limit=3,
    )

    news_block = "\n".join(
        f"- [{n['source']}] {n['title']}: {n['description'][:240]}" for n in vetted_news[:10]
    )
    used_block = "\n".join(f"- {u}" for u in used_topics[-30:]) if used_topics else "Пока пусто."

    sys_prompt = (
        f"Ты — редактор оперативного финансового Telegram-канала «{CHANNEL_NAME}». "
        "Пиши чётко, по делу, с цифрами, **жирным** и *курсивом*. Отвечай СТРОГО валидным JSON. "
        'Внутри текста используй только кавычки «елочки».'
    )
    user_prompt = f"""
СВЕЖИЕ ОТФИЛЬТРОВАННЫЕ НОВОСТИ:
{news_block}

УЖЕ ОПУБЛИКОВАННЫЕ ТЕМЫ (НЕ ПОВТОРЯТЬ):
{used_block}

Выбери самую важную финансовую новость часа и напиши пост для Telegram строго на **700–850 символов**:
1. Суть события без воды.
2. Что это значит для наших денег / вкладов / инвестиций (конкретные цифры).
3. Вопрос подписчикам.

Верни СТРОГО JSON:
{{
  "chosen_topic": "Краткая суть новости для базы истории",
  "post_title": "Яркий деловой заголовок поста",
  "cover_prompt_en": "Modern 3D editorial financial news illustration, banking and market analytics, dark navy and gold, no text, no letters, 8k",
  "cover_search_ru": "запрос из 3-4 русских слов для фото",
  "post_body_md": "Текст поста (700-850 символов) с **жирным** на цифрах, *курсивом* на терминах и `моноширинными ставками`.",
  "tags": ["новости", "финансы", "экономика"]
}}
"""
    res = invoke_llm_json(
        candidates=candidates,
        stage_title="TG-Новости (Particle / Dots3)",
        system_prompt=sys_prompt,
        user_prompt=user_prompt,
        required_keys=["chosen_topic", "post_title", "post_body_md"],
        max_tokens=4096,
    )
    res["post_body_md"] = ensure_rich_formatting(res.get("post_body_md", ""))
    return res


# ======================================================================
# ВИЗУАЛЬНЫЙ ДВИЖОК (ПОИСК ФОТО -> FLUX -> ВАРИАТИВНЫЙ АРТ + CDN)
# ======================================================================
import random


def validate_and_convert_image(raw_bytes: bytes) -> bytes | None:
    if not raw_bytes or len(raw_bytes) < 12_000:
        return None
    try:
        img = Image.open(io.BytesIO(raw_bytes))
        if img.width < 550 or img.height < 300:
            return None
        img = img.convert("RGB")
        if img.width > 1400:
            ratio = 1280 / img.width
            img = img.resize((1280, int(img.height * ratio)), Image.Resampling.LANCZOS)
        out = io.BytesIO()
        img.save(out, format="JPEG", quality=92, optimize=True)
        return out.getvalue()
    except Exception:
        return None


def render_clean_financial_artwork() -> bytes:
    """Резервная чистая обложка 1280x720 без текста со случайной палитрой и графиком."""
    palettes = [
        ((11, 24, 48), (24, 56, 102), (242, 194, 72), (64, 214, 156)),
        ((14, 42, 36), (28, 84, 70), (110, 232, 175), (245, 204, 92)),
        ((32, 20, 56), (68, 42, 112), (242, 175, 88), (112, 214, 255)),
        ((20, 28, 44), (48, 66, 98), (96, 185, 245), (245, 198, 78)),
        ((38, 22, 28), (82, 42, 54), (245, 188, 92), (104, 224, 178)),
    ]
    bg_top, bg_bottom, accent, highlight = random.choice(palettes)
    width, height = 1280, 720
    grid_col = tuple(min(255, c + 24) for c in bg_top)

    img = Image.new("RGB", (width, height), bg_top)
    draw = ImageDraw.Draw(img)

    for y in range(height):
        t = y / height
        r = int(bg_top[0] + (bg_bottom[0] - bg_top[0]) * t)
        g = int(bg_top[1] + (bg_bottom[1] - bg_top[1]) * t)
        b = int(bg_top[2] + (bg_bottom[2] - bg_top[2]) * t)
        draw.line([(0, y), (width, y)], fill=(r, g, b))

    for x in range(120, width - 80, 130):
        draw.line([(x, 80), (x, height - 80)], fill=grid_col, width=1)
    for y in range(110, height - 70, 100):
        draw.line([(100, y), (width - 100, y)], fill=grid_col, width=1)

    bars = [random.randint(110, 210) + idx * random.randint(18, 32) for idx in range(8)]
    bars = [min(420, b) for b in bars]
    trend_points: list[tuple[int, int]] = []

    for idx, bh in enumerate(bars):
        bx0 = 180 + idx * 120
        bx1 = bx0 + 64
        by1 = 600
        by0 = by1 - bh
        bar_color = tuple(min(255, c + (38 if idx == len(bars) - 1 else 20)) for c in bg_bottom)
        draw.rounded_rectangle([bx0, by0, bx1, by1], radius=10, fill=bar_color)
        trend_points.append((bx0 + 32, max(120, by0 - random.randint(10, 45))))

    draw.line(trend_points, fill=accent, width=7)
    for idx, (px, py) in enumerate(trend_points):
        dot_color = highlight if idx == len(trend_points) - 1 else accent
        radius = 13 if idx == len(trend_points) - 1 else 9
        draw.ellipse((px - radius - 4, py - radius - 4, px + radius + 4, py + radius + 4), fill=bg_top)
        draw.ellipse((px - radius, py - radius, px + radius, py + radius), fill=dot_color)

    draw.rounded_rectangle([64, 54, width - 64, height - 54], radius=24, outline=accent, width=3)

    out = io.BytesIO()
    img.save(out, format="JPEG", quality=93)
    return out.getvalue()


def _search_real_photo(query_ru: str, prompt_en: str) -> tuple[bytes, str] | None:
    """Ищет уникальную тематическую фотографию (Bing Images -> Wikimedia Commons)."""
    # 1. Bing Images (отлично работает без прокси с GitHub Actions)
    try:
        search_url = f"https://www.bing.com/images/search?q={quote(query_ru)}&form=HDRSC2&first=1"
        resp = HTTP_DIRECT.get(search_url, timeout=(5, 8))
        if resp.status_code == 200:
            urls = re.findall(
                r'murl(?:"|")\s*:\s*(?:"|")(https?://[^"&]+?\.(?:jpg|jpeg|png)(?:\?[^"&]*)?)(?:"|")',
                resp.text,
                flags=re.IGNORECASE,
            )
            candidates = list(dict.fromkeys(urls[:10]))
            random.shuffle(candidates)
            for candidate_url in candidates[:6]:
                try:
                    img_resp = HTTP_DIRECT.get(candidate_url, timeout=(4, 7))
                    if img_resp.status_code == 200:
                        valid = validate_and_convert_image(img_resp.content)
                        if valid:
                            return valid, candidate_url
                except Exception:
                    continue
    except Exception:
        pass

    # 2. Wikimedia Commons (надёжный международный архив без блокировок)
    try:
        wiki_q = " ".join(re.findall(r"[A-Za-z]{4,}", prompt_en)[:4]) or "finance banking economy"
        api_url = (
            "https://commons.wikimedia.org/w/api.php"
            f"?action=query&generator=search&gsrsearch=filetype:bitmap+{quote(wiki_q)}"
            "&gsrlimit=10&prop=imageinfo&iiprop=url&iiurlwidth=1280&format=json"
        )
        resp = HTTP_DIRECT.get(api_url, timeout=(5, 8))
        if resp.status_code == 200:
            pages = list(resp.json().get("query", {}).get("pages", {}).values())
            random.shuffle(pages)
            for page in pages[:6]:
                ii = page.get("imageinfo", [])
                if not ii:
                    continue
                img_url = ii[0].get("thumburl") or ii[0].get("url")
                if not img_url or not img_url.lower().endswith((".jpg", ".jpeg", ".png")):
                    continue
                try:
                    img_resp = HTTP_DIRECT.get(img_url, timeout=(4, 7))
                    if img_resp.status_code == 200:
                        valid = validate_and_convert_image(img_resp.content)
                        if valid:
                            return valid, img_url
                except Exception:
                    continue
    except Exception:
        pass

    return None


def _upload_freeimage(session: requests.Session, jpeg_bytes: bytes) -> str | None:
    b64_img = base64.b64encode(jpeg_bytes).decode("ascii")
    resp = session.post(
        "https://freeimage.host/api/1/upload",
        data={
            "key": "6d207e02198a847aa98d0a2a901485a5",
            "action": "upload",
            "source": b64_img,
            "format": "json",
        },
        timeout=(8, 15),
    )
    if resp.status_code == 200:
        url = resp.json().get("image", {}).get("url")
        if url and url.startswith("http"):
            return url
    return None


def _upload_catbox(session: requests.Session, jpeg_bytes: bytes, filename: str) -> str | None:
    resp = session.post(
        "https://catbox.moe/user/api.php",
        data={"reqtype": "fileupload"},
        files={"fileToUpload": (filename, jpeg_bytes, "image/jpeg")},
        timeout=(8, 15),
    )
    if resp.status_code == 200 and resp.text.strip().startswith("https://"):
        return resp.text.strip()
    return None


def _upload_uguu(session: requests.Session, jpeg_bytes: bytes, filename: str) -> str | None:
    resp = session.post(
        "https://uguu.se/upload",
        files={"files[]": (filename, jpeg_bytes, "image/jpeg")},
        timeout=(8, 15),
    )
    if resp.status_code == 200:
        files = resp.json().get("files", [])
        if files and str(files[0].get("url", "")).startswith("http"):
            return files[0]["url"]
    return None


def _upload_litterbox(session: requests.Session, jpeg_bytes: bytes, filename: str) -> str | None:
    resp = session.post(
        "https://litterbox.catbox.moe/resources/internals/api.php",
        data={"reqtype": "fileupload", "time": "72h"},
        files={"fileToUpload": (filename, jpeg_bytes, "image/jpeg")},
        timeout=(8, 15),
    )
    if resp.status_code == 200 and resp.text.strip().startswith("https://"):
        return resp.text.strip()
    return None


def upload_image_to_public_host_parallel(jpeg_bytes: bytes, filename: str = "cover.jpg") -> str | None:
    tasks = [
        lambda: _upload_freeimage(HTTP_DIRECT, jpeg_bytes),
        lambda: _upload_catbox(HTTP_DIRECT, jpeg_bytes, filename),
        lambda: _upload_uguu(HTTP_DIRECT, jpeg_bytes, filename),
        lambda: _upload_litterbox(HTTP_DIRECT, jpeg_bytes, filename),
    ]
    if PROXY_URL:
        tasks.append(lambda: _upload_freeimage(HTTP, jpeg_bytes))
        tasks.append(lambda: _upload_catbox(HTTP, jpeg_bytes, filename))

    with ThreadPoolExecutor(max_workers=len(tasks)) as pool:
        futures = [pool.submit(fn) for fn in tasks]
        for fut in as_completed(futures, timeout=18):
            try:
                url = fut.result()
                if url:
                    return url
            except Exception:
                continue
    return None


def prepare_single_cover_image(meta: dict) -> tuple[bytes, str]:
    """
    Готовит уникальную обложку 1280x720:
    1) Реальное фото по теме статьи (Bing Images / Wikimedia Commons);
    2) AI-генерация (Flux / Pollinations, без текста);
    3) Вариативный финансовый арт без текста (Pillow).
    """
    query_ru = (
        meta.get("cover_search_ru")
        or meta.get("chosen_topic")
        or meta.get("article_title")
        or "экономика финансы деньги"
    )
    prompt_en = meta.get(
        "cover_prompt_en",
        "Modern 3D editorial financial illustration, banking vault and growth chart, dark navy and gold, no text, no letters, 8k",
    )

    jpeg_bytes: bytes | None = None
    direct_url: str = ""
    origin_label = "вариативный арт (без текста)"

    # Уровень 1: Реальная тематическая фотография из сети
    found_photo = _search_real_photo(query_ru, prompt_en)
    if found_photo:
        jpeg_bytes, direct_url = found_photo
        origin_label = "реальное фото из сети"

    # Уровень 2: AI-генерация (Flux / Pollinations)
    if not jpeg_bytes:
        styled_prompt = (
            f"{prompt_en}, minimalist editorial magazine cover art, clean composition, "
            "no text, no words, no letters, no numbers, no watermark"
        )
        encoded_prompt = quote(styled_prompt[:280])
        seed = random.randint(10000, 99999999)
        for model_q in ("&model=flux", ""):
            ai_url = (
                f"https://image.pollinations.ai/prompt/{encoded_prompt}"
                f"?width=1280&height=720&nologo=true&seed={seed}{model_q}"
            )
            try:
                resp = HTTP_DIRECT.get(ai_url, timeout=(8, 25))
                if resp.status_code == 200:
                    valid = validate_and_convert_image(resp.content)
                    if valid:
                        jpeg_bytes = valid
                        direct_url = ai_url
                        origin_label = "AI-генерация (Flux)"
                        break
            except Exception:
                continue

    # Уровень 3: Вариативный графический арт без текста (Pillow)
    if not jpeg_bytes:
        jpeg_bytes = render_clean_financial_artwork()

    cdn_url = None
    try:
        cdn_url = upload_image_to_public_host_parallel(jpeg_bytes, filename="article_cover.jpg")
    except Exception:
        pass

    final_url = cdn_url or direct_url

    if final_url:
        try:
            HTTP_DIRECT.head(final_url, timeout=(3, 5))
        except Exception:
            pass

    log.info(
        "[✓] Обложка готова (%s, %d КБ) -> %s",
        origin_label,
        len(jpeg_bytes) // 1024,
        final_url[:75] if final_url else "локальный буфер",
    )
    return jpeg_bytes, final_url

# ----------------- СБОРКА ЛОКАЛЬНОГО DOCX -----------------
def add_formatted_paragraph(doc: Document, text: str, is_bullet: bool = False):
    p = doc.add_paragraph(style="List Bullet" if is_bullet else None)
    p.paragraph_format.space_after = Pt(6 if is_bullet else 8)

    pattern = re.compile(
        r"(\*\*\*.*?\*\*\*|\*\*.*?\*\*|`.*?`|\*[^\*\n]+?\*|(?<!\w)_[^_\n]+?_(?!\w))"
    )
    for part in pattern.split(text):
        if not part:
            continue
        if part.startswith("***") and part.endswith("***") and len(part) > 6:
            run = p.add_run(part[3:-3])
            run.bold = True
            run.italic = True
        elif part.startswith("**") and part.endswith("**") and len(part) > 4:
            p.add_run(part[2:-2]).bold = True
        elif part.startswith("`") and part.endswith("`") and len(part) > 2:
            run = p.add_run(part[1:-1])
            run.font.name = "Courier New"
            run.font.bold = True
            run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
        elif (
            (part.startswith("*") and part.endswith("*"))
            or (part.startswith("_") and part.endswith("_"))
        ) and len(part) > 2:
            p.add_run(part[1:-1]).italic = True
        else:
            p.add_run(part)
    return p


def create_docx(data: dict, filepath: Path) -> None:
    doc = Document()
    sections = data.get("sections", [])

    doc.add_heading("Служебный блок", level=2)
    meta = doc.add_paragraph()
    meta.add_run(f"Почему актуально: {data.get('why_actual', '')}\n").italic = True
    meta.add_run(
        f"Объём статьи: {calculate_dzen_article_length(data)} симв. | Разделов: {len(sections)}\n"
    ).bold = True
    meta.add_run("Теги: ").bold = True
    meta.add_run(", ".join(data.get("meta_tags", [])) + "\n")
    doc.add_paragraph("═" * 42)

    doc.add_heading(data["article_title"], level=1)

    cover_bytes = data.get("_cover_bytes")
    if cover_bytes:
        doc.add_picture(io.BytesIO(cover_bytes), width=Inches(6.0))

    add_formatted_paragraph(doc, data["article_lead"])

    for idx, sec in enumerate(sections, start=1):
        doc.add_heading(sec.get("subtitle", f"Раздел {idx}"), level=2)

        for p_text in sec.get("paragraphs", []):
            if p_text.strip():
                add_formatted_paragraph(doc, p_text.strip())

        calc_example = sec.get("calculation_example", "").strip()
        if calc_example:
            calc_p = add_formatted_paragraph(doc, f"🧮 **Практический расчёт:** {calc_example}")
            calc_p.paragraph_format.left_indent = Inches(0.25)

        for item in sec.get("bullet_points", []):
            if item.strip():
                add_formatted_paragraph(doc, item.strip(), is_bullet=True)

        smart_note = sec.get("smart_note", "").strip()
        if smart_note:
            note_p = add_formatted_paragraph(doc, f"📊 **Алгоритм Денег:** *{smart_note}*")
            note_p.paragraph_format.left_indent = Inches(0.3)

    doc.add_heading("Что в итоге?", level=2)
    add_formatted_paragraph(doc, data.get("call_to_action", ""))

    doc.save(str(filepath))


# ======================================================================
# ПУБЛИКАЦИЯ В ДЗЕН (1 ЦЕЛЬНОЕ СООБЩЕНИЕ + 1 ОБЛОЖКА) И TELEGRAM
# ======================================================================
def get_authorized_tg_client() -> TelegramClient:
    session_str = TG_STRING_SESSION_ENV
    if not session_str and TG_STRING_SESSION_FILE.is_file():
        session_str = TG_STRING_SESSION_FILE.read_text(encoding="utf-8").strip()

    if SERVER_HEADLESS and not session_str:
        raise RuntimeError(
            "На сервере (SERVER_HEADLESS=True) не найдена сессия Telegram! "
            "Задайте TG_STRING_SESSION в .env или создайте tg_session_string.txt."
        )

    client = TelegramClient(StringSession(session_str), TG_API_ID, TG_API_HASH)
    client.start()

    new_session_str = client.session.save()
    if new_session_str and new_session_str != session_str:
        TG_STRING_SESSION_FILE.write_text(new_session_str, encoding="utf-8")
        log.info("Ключ сессии Telegram сохранён в: %s", TG_STRING_SESSION_FILE)

    return client


def find_published_dzen_url(article_title: str) -> str:
    clean_channel_url = DZEN_CHANNEL_URL.replace("://m.dzen.ru/", "://dzen.ru/")
    if not clean_channel_url or "ТВОЙ_КАНАЛ" in clean_channel_url:
        return "https://dzen.ru"
    title_words = [w.lower() for w in re.findall(r"\w+", article_title) if len(w) > 4][:3]
    try:
        resp = HTTP_DIRECT.get(clean_channel_url, timeout=(3.0, 5.0))
        if resp.status_code == 200:
            soup = BeautifulSoup(resp.text, "html.parser")
            for a_tag in soup.find_all("a", href=True):
                href = a_tag["href"]
                text = a_tag.get_text(" ", strip=True).lower()
                if ("/a/" in href or "/b/" in href) and any(w in text for w in title_words):
                    return href if href.startswith("http") else f"https://dzen.ru{href}"
    except Exception:
        pass
    return clean_channel_url


def build_safe_tg_caption(title_html: str, body_md: str, footer_html: str, limit: int = 1020) -> str:
    max_body_chars = max(120, limit - len(title_html) - len(footer_html) - 20)
    raw_body = body_md.strip()
    if len(raw_body) > max_body_chars:
        raw_body = raw_body[:max_body_chars].rsplit(" ", 1)[0].rstrip(",;:-") + "…"
    caption = f"{title_html}{md_to_tg_html(raw_body)}{footer_html}"
    if len(caption) > limit:
        shrink = max(80, max_body_chars - (len(caption) - limit) - 20)
        raw_body = body_md.strip()[:shrink].rsplit(" ", 1)[0].rstrip(",;:-") + "…"
        caption = f"{title_html}{md_to_tg_html(raw_body)}{footer_html}"
    return caption


def publish_full_pipeline_to_tg(data: dict) -> None:
    if not ENABLE_TELEGRAM:
        return

    cover_bytes = data.get("_cover_bytes")
    cover_url = data.get("_cover_url", "")
    single_dzen_html = build_single_dzen_telegram_post(data, cover_url=cover_url)
    visible_len = len(BeautifulSoup(single_dzen_html, "html.parser").get_text())

    client = get_authorized_tg_client()
    try:
        log.info(
            "[1/2] Отправка цельной статьи в тех. канал (%s) для Дзена "
            "(видимый текст: %d симв., HTML: %d/%d симв., обложка: %s)...",
            TG_TECH_CHANNEL,
            visible_len,
            len(single_dzen_html),
            TG_SINGLE_MSG_HTML_LIMIT,
            "да" if cover_url else "нет",
        )
        try:
            client.send_message(
                TG_TECH_CHANNEL,
                message=single_dzen_html,
                parse_mode="html",
                link_preview=True,
                invert_media=True,
            )
        except TypeError:
            client.send_message(
                TG_TECH_CHANNEL,
                message=single_dzen_html,
                parse_mode="html",
                link_preview=True,
            )

        if DZEN_SYNC_WAIT_SECONDS > 0:
            log.info("Ожидание %d сек для импорта статьи на Дзене...", DZEN_SYNC_WAIT_SECONDS)
            time.sleep(DZEN_SYNC_WAIT_SECONDS)

        dzen_url = html.escape(find_published_dzen_url(data["article_title"]), quote=True)
        tg_insight = ensure_rich_formatting(data.get("tg_public_post", "").strip())
        tags_html = " ".join(
            f"#{re.sub(r'[^а-яА-Яa-zA-Z0-9_]', '', t.replace(' ', '_'))}"
            for t in data.get("meta_tags", [])[:4]
        )
        public_caption = build_safe_tg_caption(
            title_html=f"💰 <b>{md_to_tg_html(data['article_title'])}</b>\n\n",
            body_md=tg_insight,
            footer_html=f'\n\n👉 <a href="{dzen_url}"><b>Читать полный разбор в Дзене</b></a>\n\n{tags_html}',
        )

        log.info("[2/2] Публикация уникального поста с обложкой в %s...", TG_PUBLIC_CHANNEL)
        if cover_bytes:
            photo = io.BytesIO(cover_bytes)
            photo.name = "cover.jpg"
            client.send_file(TG_PUBLIC_CHANNEL, file=photo, caption=public_caption, parse_mode="html")
        else:
            client.send_message(TG_PUBLIC_CHANNEL, message=public_caption, parse_mode="html")
        log.info("[✓] Единая статья на Дзене и пост в основном Telegram-канале опубликованы!")
    finally:
        client.disconnect()


def publish_standalone_news_to_tg(news_data: dict) -> None:
    if not ENABLE_TELEGRAM:
        return
    img_bytes, _ = prepare_single_cover_image(news_data)
    tags = " ".join(
        f"#{re.sub(r'[^а-яА-Яa-zA-Z0-9_]', '', t.replace(' ', '_'))}"
        for t in news_data.get("tags", ["деньги", "новости"])
    )
    caption = build_safe_tg_caption(
        title_html=f"⚡️ <b>{md_to_tg_html(news_data['post_title'])}</b>\n\n",
        body_md=news_data["post_body_md"],
        footer_html=f"\n\n{tags}",
    )

    client = get_authorized_tg_client()
    try:
        log.info("Публикация новостного поста в %s...", TG_PUBLIC_CHANNEL)
        if img_bytes:
            photo = io.BytesIO(img_bytes)
            photo.name = "news.jpg"
            client.send_file(TG_PUBLIC_CHANNEL, file=photo, caption=caption, parse_mode="html")
        else:
            client.send_message(TG_PUBLIC_CHANNEL, message=caption, parse_mode="html")
        log.info("[✓] Оперативный новостной пост успешно опубликован в Telegram!")
    finally:
        client.disconnect()


# ======================================================================
# ТОЧКА ВХОДА
# ======================================================================
def main() -> None:
    parser = argparse.ArgumentParser(
        description="VPS-автопилот «Алгоритм Денег» (быстрый пайплайн на проверенных моделях)"
    )
    parser.add_argument(
        "--mode",
        choices=["full", "tg_news"],
        default="full",
        help="full — 1 аналитическая статья в день (Дзен + локальный .docx + ТГ); tg_news — оперативный пост в ТГ",
    )
    args = parser.parse_args()

    if SAVE_LOCAL_DOCX:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    start_ts = time.perf_counter()
    log.info(
        "Запуск «%s» | Режим: %s | Лимит ТГ-поста: %d HTML-симв. | Headless: %s",
        CHANNEL_NAME,
        args.mode,
        TG_SINGLE_MSG_HTML_LIMIT,
        SERVER_HEADLESS,
    )

    used_topics = load_used_topics()

    # Параллельный старт: RSS + актуальный каталог OpenRouter
    with ThreadPoolExecutor(max_workers=2) as boot_pool:
        rss_future = boot_pool.submit(collect_fresh_rss_parallel, used_topics)
        catalog_future = boot_pool.submit(discover_openrouter_free_models)
        raw_news = rss_future.result()
        catalog_future.result()

    log.info(
        "Собрано свежих новостей (без дублей): %d | В истории тем: %d (за %.1f сек)",
        len(raw_news),
        len(used_topics),
        time.perf_counter() - start_ts,
    )

    if not raw_news:
        log.warning("Свежих уникальных новостей в RSS не найдено. Завершение работы.")
        return

    # 1. Фильтрация тем
    vetted_news = filter_topics(raw_news)

    if args.mode == "tg_news":
        news_data = generate_standalone_tg_news(vetted_news, used_topics)
        log.info("Тема ТГ-новости: «%s» (%d симв.)", news_data["post_title"], len(news_data["post_body_md"]))
        publish_standalone_news_to_tg(news_data)
        save_used_topic(f"[TG] {news_data['chosen_topic']}")
        log.info("[✓] Режим tg_news завершён за %.1f сек.", time.perf_counter() - start_ts)
        return

    # 2. Аналитика и план на 3 раздела
    analyst_plan = run_analyst_stage(vetted_news, used_topics)
    log.info("[✓] Утверждена тема дня: «%s»", analyst_plan["chosen_topic"])

    # 3. Локальный RAG по RSS и knowledge_base/
    rag_query = f"{analyst_plan['chosen_topic']} {analyst_plan.get('rag_search_query', '')}".strip()
    rag_facts = retrieve_rag_context(rag_query, vetted_news)

    # 4. Написание статьи
    draft_article = run_writer_stage(analyst_plan, rag_facts)

    # 5. Редактура и уникальный ТГ-пост
    edited_article = run_editor_and_tg_stage(draft_article)

    # 6 + 7. Параллельно: модерация + подготовка обложки с параллельной загрузкой на CDN
    cover_meta = {
        "cover_search_ru": edited_article.get("cover_search_ru"),
        "cover_prompt_en": edited_article.get("cover_prompt_en"),
        "chosen_topic": edited_article.get("chosen_topic"),
        "article_title": edited_article.get("article_title"),
    }
    with ThreadPoolExecutor(max_workers=2) as final_pool:
        mod_future = final_pool.submit(run_moderation_check, edited_article)
        cover_future = final_pool.submit(prepare_single_cover_image, cover_meta)
        final_article = mod_future.result()
        cover_bytes, cover_url = cover_future.result()

    final_article["_cover_bytes"] = cover_bytes
    final_article["_cover_url"] = cover_url

    log.info(
        "Финальный текст готов: «%s» | Объём: %d симв. | Разделов: %d",
        final_article["article_title"],
        calculate_dzen_article_length(final_article),
        len(final_article["sections"]),
    )

    # 8. Локальный архив .docx и запись в историю тем
    if SAVE_LOCAL_DOCX:
        date_str = datetime.now().strftime("%Y-%m-%d_%H-%M")
        safe_title = "".join(
            c for c in final_article["article_title"] if c.isalnum() or c in (" ", "-", "_")
        ).strip()
        local_filepath = OUTPUT_DIR / f"{date_str}_{safe_title[:40]}.docx"
        create_docx(final_article, local_filepath)
        log.info("Архивный .docx сохранён локально: %s", local_filepath)

    save_used_topic(f"[DZEN] {final_article['chosen_topic']}")

    # 9. Публикация в Дзен и основной Telegram-канал
    try:
        publish_full_pipeline_to_tg(final_article)
    except Exception as e:
        log.error("Ошибка публикации в Telegram/Дзен: %s", e)

    log.info("[✓] Весь цикл завершён за %.1f сек.", time.perf_counter() - start_ts)


if __name__ == "__main__":
    main()