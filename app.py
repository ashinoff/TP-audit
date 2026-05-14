"""
Power Atlas — поиск абонентов, привязанных не к своей ТП.

Streamlit-приложение. Загружает Excel-выгрузку, анализирует распределение абонентов
по ТП и выдаёт ранжированный список подозрительных привязок с обоснованием.
Данные не сохраняются: всё обрабатывается в памяти процесса.
"""
from __future__ import annotations

import io
import re
import unicodedata
from collections import Counter, defaultdict
from dataclasses import dataclass
from typing import Optional

import pandas as pd
import streamlit as st


# ═══════════════════════════════════════ КОНФИГУРАЦИЯ СТРАНИЦЫ ════════════════════════════════════════

APP_TITLE = "Power Atlas"
APP_SUBTITLE = "Аудит привязок абонентов к ТП"

st.set_page_config(
    page_title=f"{APP_TITLE} — {APP_SUBTITLE}",
    page_icon="◐",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ═══════════════════════════════════════ СХЕМА КОЛОНОК ════════════════════════════════════════

@dataclass(frozen=True)
class FieldSpec:
    """Описание ожидаемой колонки в выгрузке."""
    key: str
    label: str
    required: bool
    role: str
    exact: tuple[str, ...] = ()
    prefix: tuple[str, ...] = ()
    contains: tuple[str, ...] = ()
    avoid_contains: tuple[str, ...] = ()


SCHEMA: tuple[FieldSpec, ...] = (
    FieldSpec(
        key="tp", label="ТП", required=True,
        role="Идентификатор трансформаторной подстанции",
        exact=("тп", "тп номер", "номер тп", "наименование тп", "код тп", "tp"),
        prefix=("тп ",),
        contains=("трансформатор",),
        avoid_contains=("ту", "тус", "стек"),
    ),
    FieldSpec(
        key="pod", label="Подключение", required=False,
        role="Питающая ПС (вышестоящая подстанция)",
        exact=("подключение", "пс", "питающая пс", "источник питания",
               "вышестоящая пс", "наименование пс", "питающая подстанция"),
        contains=("подключ", "питающ"),
    ),
    FieldSpec(
        key="feeder", label="Фидер", required=False,
        role="Линия 6/10 кВ от ПС до ТП",
        exact=("фидер", "присоединение"),
        contains=("фидер",),
    ),
    FieldSpec(
        key="np", label="Населённый пункт", required=False,
        role="Город, посёлок, село",
        exact=("населенный пункт", "населённый пункт", "город", "нп", "пункт"),
        contains=("населен", "населён"),
    ),
    FieldSpec(
        key="street", label="Улица", required=False,
        role="Название улицы или СНТ",
        exact=("улица", "ул"),
        contains=("улиц",),
    ),
    FieldSpec(
        key="house", label="Дом", required=False,
        role="Номер дома или участка",
        exact=("дом", "номер дома", "№ дома"),
    ),
    FieldSpec(
        key="ls", label="Лицевой счёт", required=False,
        role="ЛС или ЛС СТЕК — выводится в отчёте",
        exact=("лс", "лс / лс стек", "лс/лс стек", "лицевой счёт",
               "лицевой счет", "номер лс"),
        contains=("лицев",),
    ),
    FieldSpec(
        key="contract", label="Договор", required=False,
        role="Наименование договора / абонент",
        exact=("наименование договора", "договор", "контрагент",
               "плательщик", "абонент"),
        contains=("договор", "контрагент"),
    ),
    FieldSpec(
        key="num", label="№ п/п", required=False,
        role="Порядковый номер строки в исходнике",
        exact=("№ п/п", "n п/п", "n пп", "пп", "№", "n"),
    ),
    FieldSpec(
        key="tu_code", label="Код ТУ", required=False,
        role="Код точки учёта",
        exact=("код ту",),
    ),
    FieldSpec(
        key="tu_name", label="ТУ", required=False,
        role="Наименование точки учёта",
        exact=("ту", "наименование ту", "точка учёта", "точка учета"),
        avoid_contains=("стек", "код"),
    ),
)


# ═══════════════════════════════════════ УТИЛИТЫ НОРМАЛИЗАЦИИ ════════════════════════════════════════

_STREET_SUFFIX_RE = re.compile(
    r"\b(ул|улица|пер|переулок|пр|проспект|пр кт|пл|площадь|ш|шоссе|снт|с/т|сдт|"
    r"тер|тсн|днп|нп|кп|туп|тупик|б р|бульвар|кт|мкр|микрорайон)\b\.?",
    flags=re.IGNORECASE,
)
_PUNCT_RE = re.compile(r"[«»\"'.,\-()/]")
_SPACES_RE = re.compile(r"\s+")


def _norm_header(s: str) -> str:
    """Канонизирует заголовок: NFKC, lowercase, ё→е, тире/подчёркивания → пробел, точки и кавычки — прочь."""
    s = str(s)
    s = unicodedata.normalize("NFKC", s)
    s = s.lower().strip()
    s = s.replace("ё", "е")
    s = re.sub(r"[\u2010-\u2015_–—-]", " ", s)
    s = re.sub(r"[\.\"'`]", " ", s)
    s = _SPACES_RE.sub(" ", s).strip()
    return s


def detect_columns(df: pd.DataFrame) -> dict[str, Optional[str]]:
    """
    Сопоставление колонок по системе приоритетов:
        100  — exact-совпадение нормализованных строк
         80  — заголовок начинается с alias-префикса (или равен ему)
         60  — alias встречается как отдельное слово в заголовке
         40  — alias встречается как подстрока в заголовке
    avoid_contains блокирует совпадение. Каждая колонка достаётся максимум одному полю.
    """
    headers = list(df.columns)
    normed = {h: _norm_header(h) for h in headers}

    candidates: list[tuple[int, str, str]] = []
    for field in SCHEMA:
        for col, n_col in normed.items():
            if any(av in n_col for av in field.avoid_contains):
                continue
            score = 0
            if any(_norm_header(a) == n_col for a in field.exact):
                score = 100
            elif any(n_col.startswith(_norm_header(a) + " ") or n_col == _norm_header(a)
                     for a in field.prefix):
                score = 80
            elif any(re.search(rf"\b{re.escape(_norm_header(a))}\b", n_col)
                     for a in (*field.exact, *field.contains)):
                score = 60
            elif any(_norm_header(a) in n_col for a in field.contains):
                score = 40
            if score > 0:
                candidates.append((score, field.key, col))

    candidates.sort(key=lambda x: (-x[0], x[1]))
    result: dict[str, Optional[str]] = {f.key: None for f in SCHEMA}
    used: set[str] = set()
    for _, field_key, col in candidates:
        if result[field_key] is None and col not in used:
            result[field_key] = col
            used.add(col)
    return result


# ═══════════════════════════════════════ НОРМАЛИЗАЦИЯ ЗНАЧЕНИЙ ════════════════════════════════════════

def norm_street(s) -> Optional[str]:
    if pd.isna(s):
        return None
    s = str(s).lower().strip()
    s = _STREET_SUFFIX_RE.sub(" ", s)
    s = _PUNCT_RE.sub(" ", s)
    s = _SPACES_RE.sub(" ", s).strip()
    return s or None


def norm_house(h) -> Optional[str]:
    if pd.isna(h):
        return None
    s = str(h).lower().strip().replace(" ", "")
    return s or None


def norm_np(s) -> str:
    if pd.isna(s):
        return ""
    return str(s).lower().strip()


# ═══════════════════════════════════════ ОСНОВНОЙ АНАЛИЗ ════════════════════════════════════════

def analyze(
    df_raw: pd.DataFrame,
    cols: dict[str, Optional[str]],
    *,
    min_tp_size: int = 3,
    min_score: int = 30,
    weight_pod: int = 50,
    weight_feeder: int = 50,
    weight_np: int = 20,
    weight_street_orphan: int = 25,
    weight_addr_home: int = 30,
    np_dominance_threshold: float = 0.8,
) -> tuple[pd.DataFrame, dict]:
    if not cols.get("tp"):
        raise ValueError("Не выбрана колонка с ТП — без неё анализ невозможен.")

    df = df_raw.copy()
    df["_tp"] = df[cols["tp"]]
    df["_feeder"] = df[cols["feeder"]] if cols.get("feeder") else pd.Series([None] * len(df), index=df.index)
    df["_pod"] = df[cols["pod"]] if cols.get("pod") else pd.Series([None] * len(df), index=df.index)
    df["_np_raw"] = df[cols["np"]] if cols.get("np") else pd.Series([None] * len(df), index=df.index)
    df["_street_raw"] = df[cols["street"]] if cols.get("street") else pd.Series([None] * len(df), index=df.index)
    df["_house_raw"] = df[cols["house"]] if cols.get("house") else pd.Series([None] * len(df), index=df.index)

    df["_np"] = df["_np_raw"].apply(norm_np)
    df["_street"] = df["_street_raw"].apply(norm_street)
    df["_house"] = df["_house_raw"].apply(norm_house)

    street_np_tp: dict[tuple, Counter] = defaultdict(Counter)
    addr_tp: dict[tuple, Counter] = defaultdict(Counter)

    for _, row in df.iterrows():
        tp = row["_tp"]
        if pd.isna(tp):
            continue
        if pd.notna(row["_street"]):
            street_np_tp[(row["_street"], row["_np"])][tp] += 1
        if pd.notna(row["_street"]) and pd.notna(row["_house"]):
            addr_tp[(row["_np"], row["_street"], row["_house"])][tp] += 1

    tp_modes: dict[str, dict] = {}
    for tp, sub in df.groupby("_tp"):
        if pd.isna(tp):
            continue
        feeder_top = sub["_feeder"].mode()
        pod_top = sub["_pod"].mode()
        np_top = sub["_np"].mode()
        np_top_val = np_top.iloc[0] if not np_top.empty else ""
        tp_modes[tp] = {
            "feeder_top": feeder_top.iloc[0] if not feeder_top.empty else None,
            "pod_top": pod_top.iloc[0] if not pod_top.empty else None,
            "np_top": np_top_val,
            "np_share": (sub["_np"] == np_top_val).mean() if np_top_val else 0.0,
            "streets": Counter(sub["_street"].dropna()),
            "size": len(sub),
        }

    signal_counts = Counter()
    records = []

    for _, row in df.iterrows():
        tp = row["_tp"]
        if pd.isna(tp) or tp not in tp_modes:
            continue
        m = tp_modes[tp]
        if m["size"] < min_tp_size:
            continue

        score = 0
        reasons: list[str] = []

        if pd.notna(row["_pod"]) and m["pod_top"] and row["_pod"] != m["pod_top"]:
            score += weight_pod
            reasons.append(f"ПС «{row['_pod']}» ≠ «{m['pod_top']}» (норма на ТП)")
            signal_counts["ПС"] += 1

        if pd.notna(row["_feeder"]) and m["feeder_top"] and row["_feeder"] != m["feeder_top"]:
            score += weight_feeder
            reasons.append(f"Фидер «{row['_feeder']}» ≠ «{m['feeder_top']}» (норма на ТП)")
            signal_counts["Фидер"] += 1

        if (
            row["_np"]
            and m["np_top"]
            and row["_np"] != m["np_top"]
            and m["np_share"] >= np_dominance_threshold
        ):
            score += weight_np
            reasons.append(
                f"Нас. пункт «{row['_np_raw']}» (на ТП доминирует «{m['np_top']}», "
                f"{m['np_share']:.0%})"
            )
            signal_counts["Нас.пункт"] += 1

        if pd.notna(row["_street"]):
            count_here = m["streets"].get(row["_street"], 0)
            if count_here == 1:
                neighbours = street_np_tp.get((row["_street"], row["_np"]), Counter())
                others = Counter({k: v for k, v in neighbours.items() if k != tp})
                if others:
                    best_other_tp, best_other_n = others.most_common(1)[0]
                    dominant_count_here = (
                        m["streets"].most_common(1)[0][1] if m["streets"] else 0
                    )
                    if best_other_n >= 5 and dominant_count_here >= 5:
                        score += weight_street_orphan
                        reasons.append(
                            f"Улица «{row['_street_raw']}» — 1 абонент на ТП, "
                            f"но {best_other_n} на «{best_other_tp}»"
                        )
                        signal_counts["Улица-одиночка"] += 1

        if pd.notna(row["_street"]) and pd.notna(row["_house"]):
            tps_for_addr = addr_tp.get((row["_np"], row["_street"], row["_house"]), Counter())
            total_at_addr = sum(tps_for_addr.values())
            here_at_addr = tps_for_addr.get(tp, 0)
            if total_at_addr >= 4 and (here_at_addr / total_at_addr) <= 0.25:
                best_tp, best_n = tps_for_addr.most_common(1)[0]
                if best_tp != tp and best_n >= 3:
                    score += weight_addr_home
                    reasons.append(
                        f"Адрес «{row['_street_raw']} {row['_house_raw']}» — здесь "
                        f"{here_at_addr} из {total_at_addr}, преобладает ТП «{best_tp}» "
                        f"({best_n})"
                    )
                    signal_counts["Адрес"] += 1

        if score >= min_score:
            rec = {
                "Балл": score,
                "ТП (текущая)": tp,
                "Причины": "; ".join(reasons),
            }
            for key in ("num", "ls", "contract", "tu_code", "tu_name"):
                col = cols.get(key)
                if col:
                    rec[col] = row[col]
            for src_key, label in (
                ("_np_raw", "Нас. пункт"),
                ("_street_raw", "Улица"),
                ("_house_raw", "Дом"),
                ("_pod", "Подключение"),
                ("_feeder", "Фидер"),
            ):
                rec[label] = row[src_key]
            records.append(rec)

    result = pd.DataFrame(records)
    if not result.empty:
        result = result.sort_values(["Балл", "ТП (текущая)"], ascending=[False, True])
        result = result.reset_index(drop=True)

    summary = {
        "rows_total": len(df),
        "tps_total": df["_tp"].nunique(dropna=True),
        "tps_analyzed": sum(1 for m in tp_modes.values() if m["size"] >= min_tp_size),
        "suspicious_total": len(result),
        "signal_counts": dict(signal_counts),
    }
    return result, summary


# ═══════════════════════════════════════ СВОДКА И ЭКСПОРТ ════════════════════════════════════════

def per_tp_summary(result: pd.DataFrame) -> pd.DataFrame:
    if result.empty:
        return pd.DataFrame()
    return (
        result.groupby("ТП (текущая)")
        .agg(
            Подозрительных=("Балл", "size"),
            Сумма_баллов=("Балл", "sum"),
            Макс_балл=("Балл", "max"),
        )
        .sort_values(["Макс_балл", "Сумма_баллов"], ascending=False)
        .reset_index()
    )


def make_xlsx_bytes(result: pd.DataFrame, summary: dict, tp_agg: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        (result if not result.empty else pd.DataFrame({"Результат": ["Ничего не найдено"]})).to_excel(
            writer, sheet_name="Подозрительные", index=False
        )
        (tp_agg if not tp_agg.empty else pd.DataFrame({"Результат": ["Ничего не найдено"]})).to_excel(
            writer, sheet_name="По ТП", index=False
        )
        sig = summary.get("signal_counts", {})
        stats_rows = [
            ("Всего строк во входе", summary.get("rows_total", 0)),
            ("Всего ТП", summary.get("tps_total", 0)),
            ("ТП, прошедших фильтр размера", summary.get("tps_analyzed", 0)),
            ("Подозрительных записей", summary.get("suspicious_total", 0)),
            ("— по сигналу «ПС»", sig.get("ПС", 0)),
            ("— по сигналу «Фидер»", sig.get("Фидер", 0)),
            ("— по сигналу «Нас.пункт»", sig.get("Нас.пункт", 0)),
            ("— по сигналу «Улица-одиночка»", sig.get("Улица-одиночка", 0)),
            ("— по сигналу «Адрес»", sig.get("Адрес", 0)),
        ]
        pd.DataFrame(stats_rows, columns=["Показатель", "Значение"]).to_excel(
            writer, sheet_name="Статистика", index=False
        )
        method = pd.DataFrame(
            [
                ("ПС", 50,
                 "Подключение (питающая ПС) отличается от доминирующего на ТП. "
                 "Все абоненты одной ТП физически питаются от одной ПС — "
                 "расхождение почти всегда означает ошибку."),
                ("Фидер", 50,
                 "Фидер отличается от доминирующего на ТП. "
                 "Физически невозможно — следовательно ошибка."),
                ("Нас. пункт", 20,
                 "Нас. пункт отличается от доминирующего, "
                 "если тот охватывает ≥80% абонентов ТП."),
                ("Улица-одиночка", 25,
                 "На текущей ТП эта улица — единственная, "
                 "при этом в этом же нас. пункте на другой ТП этой улицы ≥5 абонентов."),
                ("Адрес", 30,
                 "Связка «улица + дом» в основном (≥75%) относится к другой ТП "
                 "при наличии ≥4 записей по адресу."),
            ],
            columns=["Сигнал", "Баллы по умолчанию", "Описание"],
        )
        method.to_excel(writer, sheet_name="Методика", index=False)

        from openpyxl.utils import get_column_letter
        for sheet_name in writer.book.sheetnames:
            ws = writer.book[sheet_name]
            for col_idx, col in enumerate(ws.columns, start=1):
                max_len = 0
                for cell in col:
                    v = cell.value
                    if v is not None:
                        max_len = max(max_len, min(len(str(v)), 80))
                ws.column_dimensions[get_column_letter(col_idx)].width = max_len + 2
    return buf.getvalue()


# ═══════════════════════════════════════ ВИЗУАЛЬНАЯ ОБОЛОЧКА ════════════════════════════════════════

CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght,SOFT@9..144,300..900,0..100&family=Instrument+Sans:ital,wght@0,400..700;1,400..700&family=JetBrains+Mono:wght@300;400;500;600&display=swap');

:root {
    --bg: #FAF7F0;
    --bg-soft: #F2EDE0;
    --bg-card: #FFFCF5;
    --ink: #1A1614;
    --ink-soft: #4A4540;
    --ink-mute: #8A857F;
    --rule: #D9D2C2;
    --accent: #B8540C;
    --accent-soft: #E6A368;
    --accent-bg: #FCE8D3;
    --good: #4F6B3A;
    --warn: #B8540C;
    --bad: #9B2D1F;
}

header[data-testid="stHeader"] { display: none; }
#MainMenu, footer { visibility: hidden; }
.stDeployButton { display: none; }

.stApp {
    background:
        radial-gradient(ellipse at top right, #F5EDD8 0%, transparent 60%),
        radial-gradient(ellipse at bottom left, #F0E6CE 0%, transparent 55%),
        #FAF7F0;
    color: var(--ink);
    font-family: 'Instrument Sans', system-ui, sans-serif;
    font-feature-settings: "ss01" on, "ss02" on, "cv11" on;
}

section[data-testid="stSidebar"] {
    background: #F2EDE0;
    border-right: 1px solid var(--rule);
}
section[data-testid="stSidebar"] * {
    font-family: 'Instrument Sans', system-ui, sans-serif;
}
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3 {
    font-family: 'Fraunces', serif;
    font-weight: 500;
    letter-spacing: -0.01em;
    color: var(--ink);
}

.main .block-container {
    padding-top: 2.5rem;
    padding-bottom: 6rem;
    max-width: 1200px;
}

h1, h2, h3, h4 {
    font-family: 'Fraunces', serif !important;
    font-weight: 500;
    letter-spacing: -0.02em;
    color: var(--ink);
    font-variation-settings: "opsz" 96, "SOFT" 0;
}
h1 { font-size: 3.2rem; line-height: 1.05; }
h2 { font-size: 2rem; line-height: 1.15; margin-top: 2rem; }
h3 { font-size: 1.4rem; }

.pa-hero {
    border-top: 1px solid var(--ink);
    border-bottom: 1px solid var(--ink);
    padding: 1.5rem 0 2rem 0;
    margin-bottom: 2.5rem;
    display: grid;
    grid-template-columns: 1fr auto;
    gap: 2rem;
    align-items: end;
}
.pa-hero-title {
    font-family: 'Fraunces', serif;
    font-weight: 400;
    font-size: 4.5rem;
    line-height: 0.95;
    letter-spacing: -0.04em;
    color: var(--ink);
    font-variation-settings: "opsz" 144, "SOFT" 30;
}
.pa-hero-title em {
    font-style: italic;
    font-weight: 300;
    color: var(--accent);
}
.pa-hero-meta {
    text-align: right;
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.72rem;
    line-height: 1.6;
    letter-spacing: 0.04em;
    text-transform: uppercase;
    color: var(--ink-soft);
}
.pa-hero-meta .lbl { color: var(--ink-mute); }
.pa-hero-meta .val { color: var(--ink); }
.pa-hero-sub {
    grid-column: 1 / -1;
    margin-top: 0.8rem;
    font-family: 'Instrument Sans', sans-serif;
    font-size: 1.05rem;
    line-height: 1.5;
    color: var(--ink-soft);
    max-width: 38em;
}

.pa-section-num {
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.72rem;
    letter-spacing: 0.15em;
    color: var(--accent);
    text-transform: uppercase;
    margin-bottom: 0.2rem;
    display: block;
}
.pa-section-rule {
    height: 1px;
    background: var(--ink);
    margin: 0.4rem 0 1.2rem 0;
}

.pa-card {
    background: var(--bg-card);
    border: 1px solid var(--rule);
    padding: 1.3rem 1.5rem;
    box-shadow: 0 1px 0 var(--rule);
}

.pa-schema {
    border: 1px solid var(--ink);
    margin-top: 0.5rem;
    background: var(--bg-card);
}
.pa-schema-row {
    display: grid;
    grid-template-columns: 32px 200px 1fr 110px;
    border-bottom: 1px solid var(--rule);
    align-items: stretch;
}
.pa-schema-row:last-child { border-bottom: none; }
.pa-schema-row.head {
    background: var(--ink);
    color: #FAF7F0;
}
.pa-schema-row.head > div {
    padding: 0.55rem 0.9rem;
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.7rem;
    letter-spacing: 0.12em;
    text-transform: uppercase;
}
.pa-schema-row > div {
    padding: 0.7rem 0.9rem;
    font-size: 0.94rem;
}
.pa-schema-row .num {
    font-family: 'JetBrains Mono', monospace;
    color: var(--ink-mute);
    font-size: 0.78rem;
    border-right: 1px solid var(--rule);
    text-align: center;
    padding-top: 0.85rem;
}
.pa-schema-row .name {
    font-family: 'Fraunces', serif;
    font-size: 1.05rem;
    font-weight: 500;
    color: var(--ink);
    border-right: 1px solid var(--rule);
}
.pa-schema-row .role {
    color: var(--ink-soft);
    border-right: 1px solid var(--rule);
}
.pa-schema-row .req {
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.7rem;
    text-align: center;
    padding-top: 0.85rem;
    letter-spacing: 0.08em;
    text-transform: uppercase;
}
.pa-schema-row .req.required { color: var(--accent); font-weight: 600; }
.pa-schema-row .req.optional { color: var(--ink-mute); }

[data-testid="stMetric"] {
    background: var(--bg-card);
    border: 1px solid var(--rule);
    padding: 1rem 1.2rem;
}
[data-testid="stMetricLabel"] {
    font-family: 'JetBrains Mono', monospace !important;
    font-size: 0.7rem !important;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    color: var(--ink-mute) !important;
}
[data-testid="stMetricValue"] {
    font-family: 'Fraunces', serif !important;
    font-size: 2.4rem !important;
    color: var(--ink) !important;
    font-weight: 400 !important;
    letter-spacing: -0.02em;
    line-height: 1 !important;
}

.stButton button, .stDownloadButton button {
    font-family: 'Instrument Sans', sans-serif !important;
    font-weight: 600 !important;
    letter-spacing: 0.02em !important;
    border-radius: 0 !important;
    border: 1px solid var(--ink) !important;
    background: var(--ink) !important;
    color: var(--bg) !important;
    padding: 0.7rem 1.4rem !important;
    transition: all 0.15s ease;
}
.stButton button:hover, .stDownloadButton button:hover {
    background: var(--accent) !important;
    border-color: var(--accent) !important;
    color: #FFFFFF !important;
    transform: translateY(-1px);
}

[data-testid="stFileUploader"] section {
    background: var(--bg-card);
    border: 2px dashed var(--ink-soft) !important;
    border-radius: 0 !important;
    padding: 2rem !important;
}
[data-testid="stFileUploader"] section:hover {
    border-color: var(--accent) !important;
    background: var(--accent-bg);
}

.stTabs [data-baseweb="tab-list"] {
    gap: 0;
    border-bottom: 1px solid var(--ink);
    background: transparent;
}
.stTabs [data-baseweb="tab"] {
    border-radius: 0 !important;
    background: transparent !important;
    color: var(--ink-soft) !important;
    font-family: 'JetBrains Mono', monospace !important;
    font-size: 0.78rem !important;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    padding: 0.7rem 1.2rem !important;
    border-bottom: 2px solid transparent !important;
    margin-bottom: -1px !important;
}
.stTabs [aria-selected="true"] {
    color: var(--ink) !important;
    border-bottom-color: var(--accent) !important;
    background: transparent !important;
}

[data-testid="stDataFrame"] {
    border: 1px solid var(--rule);
}

.stAlert {
    border-radius: 0 !important;
    border-left: 3px solid var(--accent);
    background: var(--bg-card) !important;
}

.stSlider [data-baseweb="slider"] [role="slider"] {
    background: var(--accent) !important;
    border-color: var(--ink) !important;
}

.stSelectbox [data-baseweb="select"] {
    border-radius: 0 !important;
}

.pa-caption {
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.7rem;
    color: var(--ink-mute);
    letter-spacing: 0.08em;
    text-transform: uppercase;
}

.pa-legend {
    margin-top: 1.5rem;
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.78rem;
    color: var(--ink-soft);
    line-height: 1.9;
}
.pa-legend .h { color: var(--bad); font-weight: 600; }
.pa-legend .m { color: var(--accent); font-weight: 600; }
.pa-legend .l { color: var(--ink-mute); font-weight: 600; }
</style>
"""


def render_hero(rows_loaded: Optional[int] = None) -> None:
    today = pd.Timestamp.today().strftime("%d · %m · %Y")
    meta_rows = ""
    if rows_loaded:
        formatted = f"{rows_loaded:,}".replace(",", " ")
        meta_rows = f'<div><span class="lbl">Загружено</span> · <span class="val">{formatted}</span></div>'
    st.markdown(
        f"""
        <div class="pa-hero">
            <div>
                <div class="pa-section-num">Аудит распределительной сети</div>
                <div class="pa-hero-title">Power<br><em>Atlas</em></div>
            </div>
            <div class="pa-hero-meta">
                <div><span class="lbl">Версия</span> · <span class="val">1.0</span></div>
                <div><span class="lbl">Дата</span> · <span class="val">{today}</span></div>
                {meta_rows}
            </div>
            <div class="pa-hero-sub">
                Поиск абонентов, которые с большой вероятностью привязаны не к своей трансформаторной
                подстанции. Загрузите выгрузку — получите ранжированный список подозрительных привязок
                с обоснованием. Файл обрабатывается в памяти процесса и нигде не сохраняется.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_section_header(num: str, title: str) -> None:
    st.markdown(
        f"""
        <div style="margin-top: 2.5rem;">
            <span class="pa-section-num">— {num}</span>
            <h2 style="margin: 0.2rem 0 0 0;">{title}</h2>
            <div class="pa-section-rule"></div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_schema_table() -> None:
    rows = []
    for i, field in enumerate(SCHEMA, start=1):
        req_class = "required" if field.required else "optional"
        req_text = "Обязательно" if field.required else "Желательно"
        rows.append(
            f"""
            <div class="pa-schema-row">
                <div class="num">{i:02d}</div>
                <div class="name">{field.label}</div>
                <div class="role">{field.role}</div>
                <div class="req {req_class}">{req_text}</div>
            </div>
            """
        )
    st.markdown(
        f"""
        <div class="pa-schema">
            <div class="pa-schema-row head">
                <div></div>
                <div>Колонка</div>
                <div>Назначение</div>
                <div>Статус</div>
            </div>
            {''.join(rows)}
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_method_card() -> None:
    rows = [
        ("ПС", 50, "Питающая ПС отличается от доминирующей на ТП. Сильнейший сигнал — физически невозможно."),
        ("Фидер", 50, "Фидер отличается от доминирующего на ТП. Так же как ПС — физически невозможно."),
        ("Нас. пункт", 20, "Населённый пункт отличается, если доминирующий охватывает не менее 80% ТП."),
        ("Улица", 25, "Улица встречается на ТП один раз, но в этом же нас. пункте на другой ТП — пять и более раз."),
        ("Адрес", 30, "Связка «улица + дом» в основном относится к другой ТП (минимум 4 записи)."),
    ]
    html = '<div class="pa-schema"><div class="pa-schema-row head"><div></div><div>Сигнал</div><div>Логика</div><div>Балл</div></div>'
    for i, (name, pts, desc) in enumerate(rows, start=1):
        html += f"""
        <div class="pa-schema-row">
            <div class="num">{i:02d}</div>
            <div class="name">{name}</div>
            <div class="role">{desc}</div>
            <div class="req required">{pts}</div>
        </div>
        """
    html += "</div>"
    st.markdown(html, unsafe_allow_html=True)


# ═══════════════════════════════════════ ОСНОВНОЙ UI ════════════════════════════════════════

def main() -> None:
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
    render_hero()

    with st.sidebar:
        st.markdown(
            '<div class="pa-section-num">— Настройки</div>'
            '<h3 style="margin:0.2rem 0 1rem 0">Параметры</h3>',
            unsafe_allow_html=True,
        )
        min_score = st.slider("Минимальный балл для отчёта", 20, 150, 30, step=5)
        min_tp_size = st.slider(
            "Минимум абонентов на ТП", 2, 20, 3,
            help="ТП с меньшим числом абонентов пропускаются — нет статистики.",
        )
        np_dom = st.slider(
            "Доминирование нас. пункта", 0.5, 1.0, 0.8, step=0.05,
            help="Доля абонентов в доминирующем нас. пункте, чтобы сигнал сработал.",
        )

        st.markdown('<div class="pa-section-rule"></div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="pa-section-num">— Веса сигналов</div>',
            unsafe_allow_html=True,
        )
        w_pod = st.number_input("ПС (Подключение)", 0, 200, 50, step=5)
        w_feeder = st.number_input("Фидер", 0, 200, 50, step=5)
        w_np = st.number_input("Нас. пункт", 0, 200, 20, step=5)
        w_street = st.number_input("Улица-одиночка", 0, 200, 25, step=5)
        w_addr = st.number_input("Адрес у другой ТП", 0, 200, 30, step=5)

        st.markdown('<div class="pa-section-rule"></div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="pa-caption">Файл не покидает процесс. Данные не сохраняются.</div>',
            unsafe_allow_html=True,
        )

    # ── Секция 01: схема входных колонок ────────────────────────────────────────
    render_section_header("01 · Структура входных данных", "Какие колонки нужны")
    st.markdown(
        '<p style="color:var(--ink-soft);max-width:42em;margin-top:-0.5rem;">'
        'Приложение автоматически распознаёт колонки выгрузки. Названия и порядок '
        'могут отличаться — главное, чтобы заголовки были осмысленными. После загрузки '
        'можно проверить и при необходимости скорректировать сопоставление.</p>',
        unsafe_allow_html=True,
    )
    render_schema_table()

    # ── Секция 02: загрузка файла ────────────────────────────────────────────────
    render_section_header("02 · Загрузка", "Excel-выгрузка абонентов")

    uploaded = st.file_uploader(
        "Перетащите файл или нажмите для выбора",
        type=["xlsx", "xlsm"],
        accept_multiple_files=False,
        label_visibility="visible",
    )

    if uploaded is None:
        render_section_header("03 · Методика", "Как считаются баллы")
        st.markdown(
            '<p style="color:var(--ink-soft);max-width:42em;margin-top:-0.5rem;">'
            'Для каждого абонента считается сумма баллов по пяти независимым сигналам. '
            'Сигналы независимы: чем больше совпало — тем выше уверенность в ошибке привязки.</p>',
            unsafe_allow_html=True,
        )
        render_method_card()
        st.markdown(
            """
            <div class="pa-legend">
                <span class="h">≥ 100</span> &nbsp; почти наверняка ошибка<br>
                <span class="m">50 — 99</span> &nbsp; высокая вероятность, проверить вручную<br>
                <span class="l">30 — 49</span> &nbsp; кандидат на проверку
            </div>
            """,
            unsafe_allow_html=True,
        )
        return

    try:
        with st.spinner("Читаю файл…"):
            xls = pd.ExcelFile(uploaded)
            sheet_name = xls.sheet_names[0]
            if len(xls.sheet_names) > 1:
                sheet_name = st.selectbox("Лист", xls.sheet_names, index=0)
            df_raw = pd.read_excel(xls, sheet_name=sheet_name)
    except Exception as exc:
        st.error(f"Не удалось прочитать файл: {exc}")
        return

    # ── Секция 03: сопоставление колонок ────────────────────────────────────────
    rows_formatted = f"{len(df_raw):,}".replace(",", " ")
    render_section_header(
        "03 · Сопоставление",
        f"Найдено в файле: {rows_formatted} строк, {len(df_raw.columns)} колонок",
    )

    auto_cols = detect_columns(df_raw)
    required_missing = [
        f.label for f in SCHEMA if f.required and not auto_cols.get(f.key)
    ]
    detected_count = sum(1 for v in auto_cols.values() if v is not None)

    st.markdown(
        f'<div class="pa-caption" style="margin-bottom:0.8rem;">'
        f'Автоматически распознано: {detected_count} из {len(SCHEMA)}. '
        f'Раскройте блок ниже, чтобы проверить и при необходимости поправить.</div>',
        unsafe_allow_html=True,
    )

    with st.expander("Сопоставление колонок (нажмите, чтобы проверить)", expanded=bool(required_missing)):
        st.markdown(
            '<p style="color:var(--ink-soft);font-size:0.9rem;">'
            'Если автоматически найденные колонки неверны — выберите нужные вручную. '
            '<strong>Колонка ТП обязательна</strong> — без неё анализ невозможен.</p>',
            unsafe_allow_html=True,
        )
        options = [None] + list(df_raw.columns)
        manual_cols = {}
        c1, c2, c3 = st.columns(3)
        for i, field in enumerate(SCHEMA):
            target = (c1, c2, c3)[i % 3]
            with target:
                default = auto_cols.get(field.key)
                idx = options.index(default) if default in options else 0
                label = f"{field.label}{' *' if field.required else ''}"
                manual_cols[field.key] = st.selectbox(
                    label, options, index=idx, key=f"col_{field.key}",
                    help=field.role,
                )
        cols = manual_cols

    if not cols.get("tp"):
        st.error("**Не указана колонка с ТП** — без неё анализ невозможен. Раскройте блок сопоставления выше и выберите её вручную.")
        return

    # ── Секция 04: анализ ───────────────────────────────────────────────────────
    render_section_header("04 · Результат", "Подозрительные привязки")

    with st.spinner("Анализирую…"):
        try:
            result, summary = analyze(
                df_raw, cols,
                min_tp_size=min_tp_size,
                min_score=min_score,
                weight_pod=w_pod,
                weight_feeder=w_feeder,
                weight_np=w_np,
                weight_street_orphan=w_street,
                weight_addr_home=w_addr,
                np_dominance_threshold=np_dom,
            )
        except Exception as exc:
            st.error(f"Ошибка анализа: {exc}")
            return

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Строк во входе", f"{summary['rows_total']:,}".replace(",", " "))
    m2.metric("ТП всего", f"{summary['tps_total']:,}".replace(",", " "))
    m3.metric("В анализе", f"{summary['tps_analyzed']:,}".replace(",", " "))
    m4.metric("Подозрительных", f"{summary['suspicious_total']:,}".replace(",", " "))

    sig = summary.get("signal_counts", {})
    if sig:
        st.markdown('<div style="height:0.5rem;"></div>', unsafe_allow_html=True)
        s1, s2, s3, s4, s5 = st.columns(5)
        s1.metric("По ПС", sig.get("ПС", 0))
        s2.metric("По фидеру", sig.get("Фидер", 0))
        s3.metric("По нас. пункту", sig.get("Нас.пункт", 0))
        s4.metric("Улица-одиночка", sig.get("Улица-одиночка", 0))
        s5.metric("Адрес у другой ТП", sig.get("Адрес", 0))

    if result.empty:
        st.warning("Ничего не найдено при текущих настройках. Попробуйте снизить порог балла в боковой панели.")
        return

    tp_agg = per_tp_summary(result)

    st.markdown('<div style="height:1.5rem;"></div>', unsafe_allow_html=True)

    tab1, tab2, tab3 = st.tabs(["— Подозрительные", "— Сводка по ТП", "— Экспорт"])

    with tab1:
        f1, f2 = st.columns([2, 1])
        with f1:
            tp_filter = st.text_input("Фильтр по ТП (подстрока)", "")
        with f2:
            min_show = st.slider("Минимальный балл к отображению", min_score, 200, min_score, step=5)
        view = result.copy()
        if tp_filter:
            view = view[view["ТП (текущая)"].astype(str).str.contains(tp_filter, case=False, na=False)]
        view = view[view["Балл"] >= min_show]
        st.dataframe(view, hide_index=True, use_container_width=True, height=520)

    with tab2:
        st.dataframe(tp_agg, hide_index=True, use_container_width=True, height=520)

    with tab3:
        st.markdown(
            '<p style="color:var(--ink-soft);">'
            'Итоговый файл содержит четыре листа: <em>Подозрительные</em>, '
            '<em>По ТП</em>, <em>Статистика</em>, <em>Методика</em>. '
            'Колонки автоматически отформатированы по ширине.</p>',
            unsafe_allow_html=True,
        )
        xlsx_bytes = make_xlsx_bytes(result, summary, tp_agg)
        st.download_button(
            "Скачать отчёт ( .xlsx )",
            data=xlsx_bytes,
            file_name=f"power-atlas_{pd.Timestamp.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )


if __name__ == "__main__":
    main()
