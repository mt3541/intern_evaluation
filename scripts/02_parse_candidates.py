"""
Stage 2: Parse cached CV/Portfolio text into structured per-candidate JSON.

For each unique candidate, merges CV + PORTFOLIO text and extracts:
  Identity:        name, university, department, status, gpa
  Background:     teknofest_level
  Heuristic 1-5: real_project, portfolio, coursework, tools, rc_uav, english
  Notes:          short auto-generated summary + raw signals for review

The 1-5 scores are best-effort heuristic guesses meant as a STARTING POINT.
Reviewers should edit extracted/candidates.json before running stage 3.

Tech-interview (O-S) and HR-feedback (N) columns are NOT filled — they require
the actual interview / HR meeting and must remain blank in the init version.
"""
import json, re, pathlib, unicodedata
from collections import defaultdict

ROOT = pathlib.Path(__file__).resolve().parent.parent
TXT_DIR = ROOT / "cv_text_cache"
OUT = ROOT / "extracted" / "candidates_auto.json"
OUT.parent.mkdir(exist_ok=True)

# ---------- Reference vocab (mirror of the workbook lookup sheets) ----------
UNIVERSITIES = {
    "METU (ODTÜ)": ["metu", "odtü", "odtu", "middle east technical"],
    "ITU": ["istanbul technical university", "istanbul teknik", "i̇tü", "itü"],
    "Bilkent": ["bilkent"],
    "Boğaziçi": ["boğaziçi", "bogazici", "bosphorus"],
    "Koç": ["koç university", "koc university", "koç üniversitesi"],
    "Sabancı": ["sabancı", "sabanci"],
    "Hacettepe": ["hacettepe"],
    "Yıldız Teknik": ["yıldız teknik", "yildiz teknik", "yildiz technical", " ytu "],
    "Gazi": ["gazi üniversitesi", "gazi university", "gazi univ"],
    "TOBB ETÜ": ["tobb"],
    "Ankara Üniversitesi": ["ankara üniversitesi", "ankara university"],
    "Ege": ["ege üniversitesi", "ege university"],
    "Erciyes": ["erciyes"],
    "Eskişehir Osmangazi": ["osmangazi"],
    "Eskişehir Teknik (Anadolu)": ["eskişehir teknik", "eskisehir teknik", "anadolu üniversitesi"],
    "Atılım": ["atılım", "atilim"],
    "İzmir Yüksek Teknoloji (IYTE)": ["iyte", "izmir yüksek", "izmir institute"],
    "Çukurova": ["çukurova", "cukurova"],
    "Selçuk": ["selçuk", "selcuk üniversitesi"],
    "Karadeniz Teknik (KTÜ)": ["karadeniz teknik", "ktu", "ktü"],
    "Samsun Üniversitesi": ["samsun üniversitesi", "samsun university"],
    "THK Üniversitesi": ["thk", "türk hava kurumu", "turkish aeronautical"],
}

DEPARTMENTS = {
    "Aerospace Engineering": ["aerospace eng", "uzay müh", "uçak ve uzay", "havacılık ve uzay", "havacilik ve uzay"],
    "Aeronautical Engineering": ["aeronautical eng", "uçak müh", "ucak muh"],
    "Astronautical Engineering": ["astronautical"],
    "Electrical & Electronics Engineering": ["electrical and electronics", "electrical & electronics",
                                              "elektrik elektronik", "elektrik-elektronik"],
    "Electronics Engineering": ["electronics eng", "elektronik müh", "elektronik ve haberle"],
    "Control Engineering": ["control and automation", "kontrol ve otomasyon", "control engineering",
                            "kontrol müh"],
    "Mechatronics Engineering": ["mechatronics", "mekatronik"],
    "Mechanical Engineering": ["mechanical eng", "makine müh", "makine muh"],
    "Aircraft Maintenance / Airframes": ["aircraft maintenance", "uçak bakım", "airframe"],
    "Computer Engineering": ["computer eng", "bilgisayar müh"],
    "Software Engineering": ["software eng", "yazılım müh"],
    "Industrial Engineering": ["industrial eng", "endüstri müh"],
    "Physics": ["physics", "fizik bölümü"],
}

# Tools (lowercase) — used for tools score
TOOLS_KEYWORDS = {
    "matlab": 1, "simulink": 1, "python": 1, "ardupilot": 1.5, "px4": 1.5,
    "ros": 1, "ros2": 1, "ros 2": 1, "sitl": 1.5, "gazebo": 1, "xflr5": 1.5,
    "avl": 1.5, "openvsp": 1.5, "ansys": 0.5, "solidworks": 0.5, "catia": 0.5,
    "c++": 0.5, "embedded c": 0.5, "labview": 0.5, "stm32": 0.5,
    "kalman": 1, "ekf": 1, "lqr": 1, "lqg": 1, "pid": 0.5,
    "mavlink": 1, "qgroundcontrol": 0.5, "mission planner": 0.5,
}

RC_KEYWORDS = ["quadcopter", "quadrotor", "drone", "uav", "iha", "model uçak", "rc plane", "rc aircraft",
               "fixed-wing", "fixed wing", "vtol", "multirotor", "flight test", "fpv"]

COURSE_KEYWORDS = ["flight mechanics", "uçuş mekaniği", "control systems", "kontrol sistem",
                   "aerodynamic", "aerodinamik", "estimation", "kalman", "modern control",
                   "linear systems", "doğrusal sistem", "dynamics", "dinamik", "signals and systems",
                   "sinyaller ve sistemler", "flight dynamics", "uçuş dinamiği", "guidance"]


_TR_MAP = str.maketrans({"ı": "i", "İ": "i", "ş": "s", "Ş": "s",
                          "ğ": "g", "Ğ": "g", "ç": "c", "Ç": "c",
                          "ü": "u", "Ü": "u", "ö": "o", "Ö": "o"})


def deaccent(s: str) -> str:
    s = s.translate(_TR_MAP)
    return ''.join(c for c in unicodedata.normalize('NFKD', s) if not unicodedata.combining(c))


def find_match(text_low: str, mapping: dict):
    """Return key whose synonym list has a hit; longest match wins."""
    hits = []
    for key, syns in mapping.items():
        for syn in syns:
            if syn in text_low:
                hits.append((len(syn), key))
    if not hits:
        return None
    hits.sort(reverse=True)
    return hits[0][1]


def detect_gpa(text: str):
    # Look for patterns like 3.34/4, 3.34 / 4.00, GPA: 3.4, AGNO 3.5, GNO: 2.94
    patterns = [
        r"(?:c?gpa|agno|gno|gano|not\s*ortalamas[ıi])\D{0,15}(\d[.,]\d{1,2})\s*/\s*(\d[.,]?\d{0,2})",
        r"(\d[.,]\d{1,2})\s*/\s*4(?:[.,]00?)?",
        r"(?:c?gpa|agno|gno|gano)\s*[:\-]?\s*(\d[.,]\d{1,2})",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            g = m.group(1).replace(",", ".")
            try:
                v = float(g)
                # if scale looked like /100 (e.g., 85/100), skip
                if v > 5:
                    continue
                return round(v, 2)
            except ValueError:
                pass
    # try /100 then convert to /4
    m = re.search(r"(?:c?gpa|gano|ortalamas[ıi])\D{0,10}(\d{2}[.,]?\d?)\s*/\s*100", text, re.IGNORECASE)
    if m:
        try:
            v = float(m.group(1).replace(",", "."))
            return round(v / 25.0, 2)
        except ValueError:
            pass
    return None


def detect_status(text: str):
    t = text.lower()
    # MSc / yuksek lisans — must be near a degree context (not "MSC Nastran" software)
    msc_ctx = (
        re.search(r"\b(m\.?sc|master['’]?s?)\s+(student|candidate|degree|in|of|thesis|programme|program)", t)
        or re.search(r"\b(yüksek\s+lisans|yuksek\s+lisans)\b", t)
        or re.search(r"\bgraduate\s+student\b", t)
    )
    if msc_ctx:
        if "graduated" in t and ("master" in t or "yüksek lisans" in t):
            return "MSc graduate"
        return "MSc student"
    # year hints (English + Turkish)
    if re.search(r"(4(\.|th|st)?\s*(year|sınıf|sinif|class)|senior\s+student|son sınıf|4\.\s*sınıf|4\.sınıf|4th-?year)", t):
        return "4th year"
    if re.search(r"(3(\.|rd)?\s*(year|sınıf|sinif|class)|3\.\s*sınıf|3\.sınıf|junior\s+student|3rd-?year|third year|third-year)", t):
        return "3rd year"
    if "aday mühendis" in t or "aday muhendis" in t:
        return "Fresh graduate (<1 yr)"
    # graduation date
    grads = re.findall(r"(?:graduat\w*|mezun)[^\n]{0,80}?(20\d{2})", t)
    grad_year = None
    for g in grads:
        y = int(g)
        if 2018 <= y <= 2028:
            grad_year = max(grad_year or 0, y)
    # Also from education dates "2019 - 2024" / "Sept 2020 – June 2024"
    for m in re.finditer(r"(20\d{2})\s*[-–—]\s*(20\d{2})", t):
        a, b = int(m.group(1)), int(m.group(2))
        if 2018 <= b <= 2030:
            grad_year = max(grad_year or 0, b)
    if grad_year:
        if grad_year > 2026:
            # still studying; approximate year by years remaining
            yrs_left = grad_year - 2026
            if yrs_left == 1:
                return "4th year"
            if yrs_left == 2:
                return "3rd year"
        elif grad_year == 2026:
            return "Fresh graduate (<1 yr)"
        elif grad_year == 2025:
            return "Fresh graduate (<1 yr)"
        elif grad_year in (2024, 2023):
            return "Graduated 1–2 yrs ago"
        elif grad_year == 2022:
            return "Graduated 2–3 yrs ago"
        elif grad_year < 2022:
            return "Graduated >3 yrs ago"
    return None


def detect_teknofest(text: str):
    t = text.lower()
    if "teknofest" not in t and "savaşan i̇ha" not in t and "savasan iha" not in t:
        return "None"
    if re.search(r"(podium|final(ist)?|1st|2nd|3rd|first place|second place|şampiyon|champion)", t):
        # Only count if near teknofest
        # Heuristic: assume connection
        return "Podium / finals"
    # Multiple participations
    if re.search(r"teknofest[^\n]{0,400}teknofest", t, re.DOTALL):
        return "Participated multiple times"
    if re.search(r"(participated|katıldı|katildim|attended|yarış)", t):
        return "Participated once"
    if re.search(r"(applied|başvur)", t):
        return "Applied only"
    return "Participated once"


def detect_english(text: str):
    t = text.lower()
    # Look near "english" or "ingilizce"
    for label, level in [
        ("native", 5), ("c2", 5), ("proficient", 5), ("fluent", 5),
        ("c1", 5), ("advanced", 4), ("ileri", 4),
        ("b2", 4), ("upper intermediate", 4), ("upper-intermediate", 4),
        ("b1", 3), ("intermediate", 3), ("orta", 3),
        ("a2", 2), ("elementary", 2),
        ("a1", 1), ("beginner", 1),
    ]:
        # search "english ... <label>" within 80 chars
        if re.search(rf"(english|i̇ngilizce|ingilizce)[^\n]{{0,80}}{label}", t):
            return level
        if re.search(rf"{label}[^\n]{{0,40}}(english|ingilizce)", t):
            return level
    # Look for TOEFL / IELTS scores
    m = re.search(r"toefl[^\n]{0,40}(\d{2,3})", t)
    if m:
        s = int(m.group(1))
        if s >= 100: return 5
        if s >= 80: return 4
        if s >= 60: return 3
        return 2
    m = re.search(r"ielts[^\n]{0,40}(\d[.,]?\d?)", t)
    if m:
        s = float(m.group(1).replace(",", "."))
        if s >= 7.5: return 5
        if s >= 6.5: return 4
        if s >= 5.5: return 3
        return 2
    if "english" in t or "i̇ngilizce" in t or "ingilizce" in t:
        return 3
    return None


def score_tools(text: str):
    t = text.lower()
    score = 0
    found = []
    for kw, w in TOOLS_KEYWORDS.items():
        if kw in t:
            score += w
            found.append(kw)
    if score >= 6: return 5, found
    if score >= 4: return 4, found
    if score >= 2.5: return 3, found
    if score >= 1: return 2, found
    return 1, found


def score_rc_uav(text: str):
    t = text.lower()
    hits = [k for k in RC_KEYWORDS if k in t]
    # Real flight test / built
    built = any(k in t for k in ["built", "flew", "yaptım", "uçurdum", "uçtu", "uçuş test", "flight test",
                                  "test flight", "uçtuğu"])
    if len(hits) >= 4 and built: return 5, hits
    if len(hits) >= 3 or built: return 4, hits
    if len(hits) >= 2: return 3, hits
    if len(hits) >= 1: return 2, hits
    return 1, hits


def score_coursework(text: str):
    t = text.lower()
    hits = [k for k in COURSE_KEYWORDS if k in t]
    if len(hits) >= 5: return 5, hits
    if len(hits) >= 3: return 4, hits
    if len(hits) >= 2: return 3, hits
    if len(hits) >= 1: return 2, hits
    return 1, hits


def score_real_project(text: str):
    t = text.lower()
    # Count "PROJECT" headings / project items / experience entries
    sig = 0
    project_signals = ["project", "proje", "internship", "staj", "experience", "deneyim"]
    for s in project_signals:
        sig += len(re.findall(s, t))
    gnc_relevance = sum(1 for k in ["uav", "iha", "drone", "quad", "flight controller", "guidance",
                                     "navigation", "kalman", "pid", "control", "kontrol", "autopilot",
                                     "ardupilot", "px4", "fixed-wing", "vtol"] if k in t)
    if gnc_relevance >= 5 and sig >= 4: return 5
    if gnc_relevance >= 3: return 4
    if gnc_relevance >= 2: return 3
    if gnc_relevance >= 1: return 2
    return 1


def score_portfolio(has_portfolio: bool, portfolio_len: int, cv_len: int):
    if not has_portfolio:
        return 2
    if portfolio_len < 500:  # essentially empty / image-only
        return 2
    if portfolio_len > 8000: return 5
    if portfolio_len > 4000: return 4
    if portfolio_len > 1500: return 3
    return 2


# ---------- Filename → canonical candidate id ----------
def canonical_id(stem: str) -> str:
    """Canonical = first 2 deaccented lowercase tokens of the filename stem.

    All CV/portfolio files for the same person share the first two name tokens
    (firstname_lastname or firstname_middlename), so this groups them reliably.
    """
    s = deaccent(stem).lower()
    tokens = [t for t in s.split("_") if t]
    return "_".join(tokens[:2])


def pretty_name(stem: str) -> str:
    s = stem
    s = re.sub(r"_PORTFOLIO.*$", "", s, flags=re.IGNORECASE)
    s = re.sub(r"_CV.*$", "", s, flags=re.IGNORECASE)
    s = re.sub(r"_Resume.*$", "", s, flags=re.IGNORECASE)
    s = re.sub(r"_cvpdf.*$", "", s, flags=re.IGNORECASE)
    parts = s.split("_")
    keep = []
    for p in parts[:3]:
        if not p or re.search(r"(cv|resume|portfolio|pdf|sonpdf|cvvpdf|tr|ypdf)", p, re.IGNORECASE):
            break
        keep.append(p)
    return " ".join(keep)


def main():
    # Group files by canonical_id
    groups = defaultdict(lambda: {"cv": [], "portfolio": []})
    for txt in sorted(TXT_DIR.glob("*.txt")):
        cid = canonical_id(txt.stem)
        is_port = "PORTFOLIO" in txt.stem.upper()
        groups[cid]["portfolio" if is_port else "cv"].append(txt)

    candidates = []
    for cid, files in sorted(groups.items()):
        cv_text = "\n".join(p.read_text(encoding="utf-8", errors="ignore") for p in files["cv"])
        port_text = "\n".join(p.read_text(encoding="utf-8", errors="ignore") for p in files["portfolio"])
        full = cv_text + "\n" + port_text
        text_low = deaccent(full.lower())
        # Detection
        uni = find_match(text_low, {k: [deaccent(s) for s in v] for k, v in UNIVERSITIES.items()})
        dept = find_match(text_low, {k: [deaccent(s) for s in v] for k, v in DEPARTMENTS.items()})
        gpa = detect_gpa(full)
        status = detect_status(full)
        teknofest = detect_teknofest(full)
        english = detect_english(full)
        tools_score, tools_found = score_tools(full)
        rc_score, rc_hits = score_rc_uav(full)
        course_score, course_hits = score_coursework(full)
        proj_score = score_real_project(full)
        port_score = score_portfolio(bool(files["portfolio"]), len(port_text), len(cv_text))
        # Pretty name from first CV file (fall back to portfolio)
        src = (files["cv"] or files["portfolio"])[0]
        name = pretty_name(src.stem)

        candidates.append({
            "id": cid,
            "source_files": {
                "cv": [p.name for p in files["cv"]],
                "portfolio": [p.name for p in files["portfolio"]],
            },
            "name": name,
            "university": uni or "Other",
            "department": dept or "Other",
            "academic_status": status or "",
            "gpa": gpa if gpa is not None else "",
            "teknofest": teknofest,
            "real_project": proj_score,
            "portfolio_score": port_score,
            "coursework": course_score,
            "tools": tools_score,
            "rc_uav": rc_score,
            "english": english if english is not None else 3,
            "hr_feedback": "",          # interview-driven; leave blank
            "controls_gnc": "",         # interview-driven
            "math_physics": "",
            "coding": "",
            "debugging": "",
            "communication": "",
            "disqualified": "N",
            "notes": "",
            "_signals": {
                "tools_found": tools_found,
                "rc_hits": rc_hits,
                "course_hits": course_hits,
                "cv_chars": len(cv_text),
                "portfolio_chars": len(port_text),
            },
        })

    OUT.write_text(json.dumps(candidates, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"Wrote {len(candidates)} candidates → {OUT.relative_to(ROOT)}")
    # Print warnings for low-confidence rows
    for c in candidates:
        warn = []
        if c["university"] == "Other": warn.append("uni")
        if c["department"] == "Other": warn.append("dept")
        if not c["academic_status"]: warn.append("status")
        if c["gpa"] == "": warn.append("gpa")
        if warn:
            print(f"  REVIEW {c['name']:30s}  missing: {','.join(warn)}")


if __name__ == "__main__":
    main()
