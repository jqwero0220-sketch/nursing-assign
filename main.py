"""
간호학과 스마트 실습지 최적 배정 시스템

버전 기록 (수정할 때마다 APP_VERSION 을 올립니다)
  v9.6  주소 자동 정리(오타·괄호·도 누락), 지도 새 주소(VPC) 지원, 로그인 없이 사용, 관리 실습 병원 목록 수정

실행:
    pip install fastapi uvicorn pandas numpy scipy openpyxl python-multipart
    set APP_PASSWORD=원하는암호        (Windows)  /  export APP_PASSWORD=원하는암호  (Mac·Linux)
    python main.py   →  http://localhost:8000

배정 방식
  - 단일 병원 모드 : 선택한 병원 기준으로 종합 매칭점수 순위를 매기고 정원까지 배정
  - 전체 병원 동시 배정 모드 : 모든 학생 x 모든 병원(정원만큼 슬롯)을 한 번에 헝가리안 알고리즘으로
    풀어 '전체 매칭점수 합'이 최대가 되도록 배정 (학생 1명 = 병원 1곳 보장)
"""
import datetime
import hashlib
import hmac
import io
import json
import math
import os
import re
import secrets
import sys
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from concurrent.futures import ThreadPoolExecutor, wait
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import random

import numpy as np
import pandas as pd
import uvicorn
from fastapi import Body, FastAPI, File, Form, Header, HTTPException, UploadFile
from fastapi.concurrency import run_in_threadpool
from fastapi.responses import HTMLResponse, StreamingResponse
from pydantic import BaseModel
from scipy.optimize import linear_sum_assignment

APP_VERSION = "v9.6"
app = FastAPI(title=f"간호학과 스마트 실습지 최적 배정 시스템 ({APP_VERSION})")

# ──────────────────────────────────────────────────────────────
# 보안: 암호는 환경변수로 관리하고, 인증 후에는 토큰으로 API를 보호한다.
# ──────────────────────────────────────────────────────────────
def _load_env_files() -> None:
    """main.py 와 같은 폴더의 keys.env / .env 를 읽어 환경변수로 등록(이미 설정된 값은 유지).
    '여기에…' 로 시작하는 자리표시 값은 무시한다. 이 파일은 외부에 공유하지 마세요."""
    here = Path(__file__).resolve().parent
    for name in ("keys.env", ".env"):
        p = here / name
        if not p.exists():
            continue
        try:
            for line in p.read_text(encoding="utf-8-sig").splitlines():
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                k, v = line.split("=", 1)
                v = v.strip().strip('"').strip("'")
                if v and not v.startswith("여기에"):
                    os.environ.setdefault(k.strip(), v)
        except Exception:
            pass


_load_env_files()
_ENV_PASSWORD = os.environ.get("APP_PASSWORD")
AUTH_REQUIRED = bool(_ENV_PASSWORD)   # APP_PASSWORD 를 설정하지 않으면 로그인 없이 바로 사용한다.
APP_PASSWORD = _ENV_PASSWORD or "no-password"
TOKEN_TTL_SEC = 8 * 60 * 60
_TOKEN_SECRET = hashlib.sha256(("tok|" + APP_PASSWORD).encode("utf-8")).digest()


def _make_token() -> str:
    """서명된 토큰: '만료시각.서명'. 서버가 재시작돼도 암호가 같으면 계속 유효하다."""
    exp = str(int(time.time()) + TOKEN_TTL_SEC)
    sig = hmac.new(_TOKEN_SECRET, exp.encode(), hashlib.sha256).hexdigest()
    return f"{exp}.{sig}"


def _token_ok(token: Optional[str]) -> bool:
    try:
        exp, sig = str(token).split(".", 1)
        good = hmac.new(_TOKEN_SECRET, exp.encode(), hashlib.sha256).hexdigest()
        return hmac.compare_digest(sig, good) and int(exp) >= time.time()
    except Exception:
        return False

MAX_UPLOAD_BYTES = 5 * 1024 * 1024
MAX_STUDENTS = 2000

MASTER_HOSPITAL_DB = {
    "3학년": {
        "1학기": {
            "성인1": [
                "가톨릭대학교 부천성모병원", "가톨릭대학교 성빈센트병원", "고려대학교안산병원",
                "순천향대학교부천병원", "순천향대학교서울병원", "용인세브란스병원",
                "인천기독병원", "중앙대학교광명병원(방중)", "한림대학교성심병원"
            ],
            "여성": [
                "가톨릭대학교 부천성모병원", "군포지샘병원", "봄빛병원(방중)",
                "순천향대학교서울병원", "우성여성병원(방중)", "인하대학교병원(방중)",
                "중앙대학교광명병원(방중)", "한빛병원"
            ]
        },
        "2학기": {
            "성인2": [
                "가톨릭대학교 부천성모병원", "가톨릭대학교 성빈센트병원", "고려대학교안산병원",
                "사랑의병원", "순천향대학교부천병원", "순천향대학교서울병원",
                "아주대학교병원", "인하대학교병원(방중)", "한도병원", "한림대학교성심병원"
            ],
            "아동": [
                "단원병원", "순천향대학교서울병원", "시화병원", "아이원병원",
                "웰봄병원", "부천서울어린이병원"
            ],
            "정신": [
                "가톨릭대학교 성빈센트병원", "계요병원", "군포시정신건강복지센터",
                "안산시정신건강복지센터", "안산시중독관리통합지원센터", "의왕시정신건강복지센터", "이음병원"
            ]
        }
    },
    "4학년": {
        "1학기": {
            "성인3": [
                "가톨릭대학교 부천성모병원", "부천세종병원", "사랑의병원",
                "순천향대학교부천병원", "아주대학교병원", "인천기독병원",
                "인하대학교병원(방중)", "중앙대학교광명병원(방중)", "한림대학교성심병원"
            ],
            "지역": [
                "근로복지공단안산병원", "단원보건소", "상록수보건소",
                "센트럴병원", "엘지이노텍"
            ],
            "관리": [
                "고려대학교구로병원", "단원병원", "서울특별시 서남병원",
                "순천향대학교부천병원", "순천향대학교서울병원", "시화병원",
                "아주대학교병원", "중앙대학교광명병원(방중)", "한림대학교동탄성심병원"
            ]
        },
        "2학기": {
            "종합": [
                "단원보건소", "단원병원", "마음건강센터", "상록수보건소",
                "순천향대학교서울병원", "시화병원", "안산시정신건강복지센터",
                "용인세브란스병원", "우성여성병원(방중)", "중앙대학교광명병원(방중)"
            ]
        }
    }
}

REQUIRED_COLS = ["학번", "이름", "성별", "GPA", "출생연도"]
OPTIONAL_COLS = ["주소", "과거실습지", "1지망병원", "2지망병원", "3지망병원"]
PREF_SCORES = (100.0, 70.0, 40.0)
NON_PREF_SCORE = 10.0
BIG_COST = 1e6


# ──────────────────────────────────────────────────────────────
# 유틸
# ──────────────────────────────────────────────────────────────
def norm_h(name) -> str:
    """병원명 비교용 정규화: '(방중)' 표기와 모든 공백 제거."""
    s = str(name or "")
    s = s.replace("(방중)", "")
    return re.sub(r"\s+", "", s)


def tie_break(student_id: str) -> float:
    """재현 가능한(프로세스가 바뀌어도 같은) 동점자 처리용 0~1 값. 점수에는 영향을 주지 않는다."""
    return int(hashlib.md5(student_id.encode("utf-8")).hexdigest()[:8], 16) / 0xFFFFFFFF


def norm_gender(v) -> str:
    s = str(v or "").strip().lower()
    if s.startswith("남") or s in ("m", "male"):
        return "남"
    if s.startswith("여") or s in ("f", "female"):
        return "여"
    return s


def hospitals_for(grade: str, semester: str, subject: str) -> List[str]:
    sem = MASTER_HOSPITAL_DB.get(grade, {}).get(semester, {})
    if subject == "ALL":
        raw = [h for lst in sem.values() for h in lst]
    else:
        raw = list(sem.get(subject, []))
    seen, out = set(), []
    for h in raw:
        k = norm_h(h)
        if k not in seen:
            seen.add(k)
            out.append(h)
    return sorted(out)


# ──────────────────────────────────────────────────────────────
# 통학 시간 추정 (동 단위 규칙 기반. 실제 서비스에서는 대중교통 경로 API로 교체)
# ──────────────────────────────────────────────────────────────
def estimate_transit(address: str, hospital: str) -> Tuple[int, int, int]:
    addr = str(address).strip()
    hosp = norm_h(hospital)
    base_time, transfers, walk_time = 35, 1, 9

    if "안산시" in addr:
        if "단원구" in addr:
            if "와동" in addr: base_time, walk_time = 32, 8
            elif "고잔동" in addr: base_time, walk_time = 20, 6
            elif "선부동" in addr: base_time, walk_time = 28, 9
            elif "초지동" in addr: base_time, walk_time = 25, 7
            elif "원곡동" in addr: base_time, walk_time = 30, 10
            else: base_time = 30
        elif "상록구" in addr:
            if "본오동" in addr: base_time, walk_time = 38, 11
            elif "사동" in addr: base_time, walk_time = 35, 10
            elif "일동" in addr or "이동" in addr: base_time, walk_time = 26, 7
            elif "성포동" in addr: base_time, walk_time = 28, 8
            else: base_time = 33
    elif "군포시" in addr:
        if "산본동" in addr: base_time, walk_time = 45, 10
        elif "금정동" in addr: base_time, walk_time = 42, 9
        elif "당동" in addr: base_time, walk_time = 48, 12
        else: base_time = 46
    elif "수원시" in addr:
        transfers = 2
        if "팔달구" in addr: base_time, walk_time = 50, 12
        elif "권선구" in addr: base_time, walk_time = 58, 14
        elif "영통구" in addr: base_time, walk_time = 65, 15
        else: base_time = 55
    elif "부천시" in addr:
        if "원미구" in addr or "중동" in addr or "상동" in addr: base_time, walk_time = 46, 10
        else: base_time, walk_time = 50, 12
    elif "안양시" in addr:
        if "동안구" in addr: base_time, walk_time = 40, 9
        else: base_time = 42

    if "광명병원" in hosp: base_time += 12
    elif "인하대" in hosp: base_time += 18; transfers = 2
    elif "성빈센트병원" in hosp: base_time += 8
    elif "계요병원" in hosp: base_time += 10

    return int(base_time), int(transfers), int(walk_time)


def satisfaction_index(mfi: float) -> int:
    """통학 부담(MFI) 기반 예상 만족도 지수 (규칙 기반, 40~98).
    실제 만족도 설문 데이터가 쌓이면 학습 모델로 교체할 수 있도록 함수로 분리."""
    return max(40, min(98, int(100 - mfi * 0.65)))


# ──────────────────────────────────────────────────────────────
# 실측 통학시간: 네이버 Geocoding(주소→좌표) + TMAP 대중교통(소요시간)
#   환경변수: NAVER_MAPS_KEY_ID, NAVER_MAPS_KEY, TMAP_APP_KEY
#   - 학생 주소는 '동 단위'까지만 외부로 보냅니다(상세 주소 미전송).
#   - 병원 주소는 hospital_addresses.json 에 직접 입력합니다(첫 실행 시 빈 양식이 자동 생성됨).
#   - 조회 결과는 transit_cache.json 에 저장되어 다음부터는 API를 다시 부르지 않습니다.
#   - 키가 없거나 조회에 실패한 구간은 규칙 기반 추정치로 자동 대체됩니다.
# ──────────────────────────────────────────────────────────────
BASE_DIR = Path(__file__).resolve().parent
NAVER_KEY_ID = os.environ.get("NAVER_MAPS_KEY_ID", "")
NAVER_KEY_SECRET = os.environ.get("NAVER_MAPS_KEY", "")
# 새 지도 서비스(VPC)는 maps.apigw.ntruss.com, 예전(Classic) 키는 naveropenapi.apigw.ntruss.com 을 쓴다.
NAVER_GEOCODE_URL = os.environ.get("NAVER_GEOCODE_URL", "https://maps.apigw.ntruss.com/map-geocode/v2/geocode")
NAVER_GEOCODE_URL_OLD = "https://naveropenapi.apigw.ntruss.com/map-geocode/v2/geocode"
TMAP_APP_KEY = os.environ.get("TMAP_APP_KEY", "")
TMAP_TRANSIT_URL = "https://apis.openapi.sk.com/transit/routes"
TRANSIT_AVAILABLE = bool(NAVER_KEY_ID and NAVER_KEY_SECRET)          # 네이버 키만 있어도 사용 가능
TRANSIT_MODE = "tmap" if TMAP_APP_KEY else "line"                   # tmap: 실제 대중교통 / line: 좌표 직선거리 추정
LINE_BASE_MIN, LINE_MIN_PER_KM = 12, 2.0   # 직선거리 추정식: 통학(분) = 12 + 2.0 × 직선거리(km)  (실측과 비교해 조정 가능)


def _rkey(dong: str, hosp_addr: str) -> str:
    return f"{TRANSIT_MODE}|{dong}||{hosp_addr}"


def haversine_km(a: Tuple[float, float], b: Tuple[float, float]) -> float:
    """(경도, 위도) 두 점 사이의 직선거리(km)."""
    (lon1, lat1), (lon2, lat2) = a, b
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dphi, dl = p2 - p1, math.radians(lon2 - lon1)
    h = math.sin(dphi / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * 6371.0 * math.asin(math.sqrt(h))


def line_estimate(s: Tuple[float, float], h: Tuple[float, float]) -> dict:
    d = haversine_km(s, h)
    return {"t": max(5, round(LINE_BASE_MIN + LINE_MIN_PER_KM * d)),
            "tr": 0 if d < 8 else (1 if d < 25 else 2), "w": 8, "src": "직선거리", "km": round(d, 1)}
HOSPITAL_ADDR_FILE = BASE_DIR / "hospital_addresses.json"
TRANSIT_CACHE_FILE = BASE_DIR / "transit_cache.json"
TRANSIT_BUDGET_SEC = 60      # 한 번의 배정에서 외부 조회에 쓸 최대 시간(초)
TRANSIT_WORKERS = 4

_cache_lock = threading.Lock()
_cache: Optional[dict] = None
_ADDR_MAP: Dict[str, str] = {}


def _load_cache() -> dict:
    global _cache
    if _cache is None:
        try:
            _cache = json.loads(TRANSIT_CACHE_FILE.read_text(encoding="utf-8"))
        except Exception:
            _cache = {}
        _cache.setdefault("geo", {})
        _cache.setdefault("route", {})
    return _cache


def _save_cache() -> None:
    try:
        with _cache_lock:
            TRANSIT_CACHE_FILE.write_text(json.dumps(_cache, ensure_ascii=False), encoding="utf-8")
    except Exception:
        pass


_CITY_PROV = {"시흥시": "경기도", "안산시": "경기도", "화성시": "경기도", "수원시": "경기도", "부천시": "경기도", "광명시": "경기도",
              "군포시": "경기도", "안양시": "경기도", "성남시": "경기도", "오산시": "경기도", "용인시": "경기도", "김포시": "경기도",
              "의왕시": "경기도", "평택시": "경기도", "고양시": "경기도", "파주시": "경기도", "과천시": "경기도", "하남시": "경기도",
              "구리시": "경기도", "광주시": "경기도"}


def clean_address(address: str) -> str:
    """사람이 손으로 쓴 주소를 지도가 알아듣게 다듬는다(오타·괄호·쉼표·도/시 누락)."""
    a = str(address or "").strip()
    a = re.split(r"[,/;]", a)[0]                       # '경기도 안산시, 경기도' → 첫 부분만
    a = re.sub(r"\(.*?\)", " ", a)                     # (기숙사) 제거
    a = a.replace("가평시", "가평군").replace("인청", "인천").replace("서울서", "서울시").replace("인천 광역시", "인천광역시")
    a = re.sub(r"^서울(?=\S*구)(?!특별시|시)", "서울시 ", a)  # '서울 강서구' → '서울시 강서구'
    a = re.sub(r"^경기(?!도)", "경기도", a)
    a = re.sub(r"(서울시|서울특별시|인천광역시|인천시)(?=\S)", r"\1 ", a)      # '서울시도봉구' → '서울시 도봉구'
    a = re.sub(r"(시)(?=[가-힣]+구)", r"\1 ", a)                              # '수원시권선구' → '수원시 권선구'
    a = re.sub(r"^경기도부천서", "경기도 부천시", a)
    a = re.sub(r"\s+", " ", a).strip()
    first = a.split(" ")[0] if a else ""
    if first in _CITY_PROV:                              # '시흥시' → '경기도 시흥시'
        a = _CITY_PROV[first] + " " + a
    return a


def dong_level(address: str) -> str:
    """상세 주소 제거: '동/읍/면'으로 끝나는 첫 토큰까지만 남긴다(없으면 앞 3토큰). 개인정보 최소화."""
    toks = clean_address(address).split()
    out: List[str] = []
    for t in toks:
        out.append(t)
        if len(t) > 1 and t.endswith(("동", "읍", "면")):
            return " ".join(out)
    return " ".join(toks[:3])


def _http_json(url: str, headers: dict, body: Optional[bytes] = None, timeout: int = 8) -> dict:
    req = urllib.request.Request(url, data=body, headers=headers, method="POST" if body is not None else "GET")
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read().decode("utf-8"))


def _geocode_raw(query: str) -> List[dict]:
    global NAVER_GEOCODE_URL
    q = "?" + urllib.parse.urlencode({"query": query})
    hdr = {"x-ncp-apigw-api-key-id": NAVER_KEY_ID, "x-ncp-apigw-api-key": NAVER_KEY_SECRET, "Accept": "application/json"}
    try:
        data = _http_json(NAVER_GEOCODE_URL + q, hdr)
    except urllib.error.HTTPError as e:
        # 새 주소에서 인증이 거절되면 옛 주소(Classic 키)로 한 번 더 시도하고, 되는 쪽을 기억한다.
        if e.code in (401, 403) and NAVER_GEOCODE_URL != NAVER_GEOCODE_URL_OLD:
            data = _http_json(NAVER_GEOCODE_URL_OLD + q, hdr)
            NAVER_GEOCODE_URL = NAVER_GEOCODE_URL_OLD
        else:
            raise
    return data.get("addresses") or []


def geocode(query: str) -> Optional[Tuple[float, float]]:
    """주소 → (경도, 위도). 못 찾으면 None(결과 캐시), 통신 오류는 예외."""
    c = _load_cache()
    if query in c["geo"]:
        v = c["geo"][query]
        return (v[0], v[1]) if v else None
    addrs = _geocode_raw(query)
    val = [float(addrs[0]["x"]), float(addrs[0]["y"])] if addrs else None
    with _cache_lock:
        c["geo"][query] = val
    return (val[0], val[1]) if val else None


def commute_dttm() -> str:
    """출근 시간대 기준(다음 평일 오전 8시, 한국시간). TMAP searchDttm 형식 yyyymmddhhmi."""
    kst = datetime.timezone(datetime.timedelta(hours=9))
    d = datetime.datetime.now(kst).date() + datetime.timedelta(days=1)
    while d.weekday() >= 5:
        d += datetime.timedelta(days=1)
    return d.strftime("%Y%m%d") + "0800"


def fetch_route(dong: str, hosp_addr: str) -> Optional[dict]:
    """동 → 병원 대중교통 소요시간. 성공 시 {'t': 분, 'tr': 환승, 'w': 도보분}을 캐시에 저장."""
    s, h = geocode(dong), geocode(hosp_addr)
    if not s or not h:
        return None
    if not TMAP_APP_KEY:  # TMAP 키가 없으면 좌표 직선거리로 추정
        res = line_estimate(s, h)
        with _cache_lock:
            _load_cache()["route"][_rkey(dong, hosp_addr)] = res
        return res
    body = json.dumps({
        "startX": str(s[0]), "startY": str(s[1]), "endX": str(h[0]), "endY": str(h[1]),
        "count": 3, "lang": 0, "format": "json", "searchDttm": commute_dttm(),
    }).encode("utf-8")
    data = _http_json(TMAP_TRANSIT_URL, {"appKey": TMAP_APP_KEY, "accept": "application/json",
                                         "content-type": "application/json"}, body)
    its = (((data.get("metaData") or {}).get("plan") or {}).get("itineraries")) or []
    if not its:
        return None
    best = min(its, key=lambda i: i.get("totalTime", 10 ** 9))
    res = {"t": max(1, round(best["totalTime"] / 60)),            # totalTime: 초
           "tr": int(best.get("transferCount", 0)),
           "w": max(0, round(best.get("totalWalkTime", 0) / 60)), "src": "TMAP"}  # totalWalkTime: 초
    with _cache_lock:
        _load_cache()["route"][_rkey(dong, hosp_addr)] = res
    return res


def all_hospital_names() -> List[str]:
    seen, out = set(), []
    for g in MASTER_HOSPITAL_DB.values():
        for s in g.values():
            for lst in s.values():
                for h in lst:
                    if norm_h(h) not in seen:
                        seen.add(norm_h(h))
                        out.append(h)
    return sorted(out)


def load_hospital_addresses() -> Dict[str, str]:
    """hospital_addresses.json → {정규화된 병원명: 주소}. 파일이 없으면 빈 양식을 만들어 준다."""
    if not HOSPITAL_ADDR_FILE.exists():
        try:
            HOSPITAL_ADDR_FILE.write_text(
                json.dumps({h: "" for h in all_hospital_names()}, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass
        return {}
    try:
        raw = json.loads(HOSPITAL_ADDR_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}
    return {norm_h(k): str(v).strip() for k, v in raw.items() if str(v).strip()}


def lookup_cached(dong: str, hospital: str) -> Optional[dict]:
    addr = _ADDR_MAP.get(norm_h(hospital))
    if not dong or not addr:
        return None
    return _load_cache()["route"].get(_rkey(dong, addr))


def prefetch_transit(students: List[dict], hospitals: List[str]) -> List[str]:
    """필요한 (동, 병원) 조합 중 캐시에 없는 것만 동시 조회. 안내 문구 목록을 돌려준다."""
    global _ADDR_MAP
    _ADDR_MAP = load_hospital_addresses()
    c = _load_cache()
    notes: List[str] = []
    missing = [h for h in hospitals if norm_h(h) not in _ADDR_MAP]
    if missing:
        notes.append(f"병원 주소가 등록되지 않은 {len(missing)}곳은 규칙 기반 추정치를 사용했습니다 "
                     f"(hospital_addresses.json 에 입력): {', '.join(missing[:5])}{' 외' if len(missing) > 5 else ''}")
    todo = set()
    for s in students:
        if not s.get("dong"):
            continue
        for h in hospitals:
            a = _ADDR_MAP.get(norm_h(h))
            if a and _rkey(s["dong"], a) not in c["route"]:
                todo.add((s["dong"], a))
    if not todo:
        return notes

    errs: List[str] = []

    def work(item):
        try:
            return fetch_route(*item) is not None
        except urllib.error.HTTPError as e:
            errs.append(f"HTTP {e.code}")
            return False
        except Exception as e:
            errs.append(type(e).__name__)
            return False

    ex = ThreadPoolExecutor(max_workers=TRANSIT_WORKERS)
    futs = [ex.submit(work, it) for it in todo]
    done, pending = wait(futs, timeout=TRANSIT_BUDGET_SEC)
    ex.shutdown(wait=False, cancel_futures=True)
    failed = sum(1 for f in done if not f.result()) + len(pending)
    _save_cache()
    if failed:
        why = f" (오류: {', '.join(sorted(set(errs))[:3])} — 키·Maps 선택·병원 주소를 확인하세요)" if errs else ""
        notes.append(f"지도 조회에 실패했거나 지연된 {failed}건(총 {len(todo)}건 중)은 규칙 기반 추정치로 대체했습니다.{why}")
    return notes


# ──────────────────────────────────────────────────────────────
# 파일 읽기 / 검증
# ──────────────────────────────────────────────────────────────
def read_table(contents: bytes, filename: str) -> pd.DataFrame:
    name = (filename or "").lower()
    if name.endswith(".csv"):
        for enc in ("utf-8-sig", "cp949"):
            try:
                return pd.read_csv(io.BytesIO(contents), dtype=str, keep_default_na=False, encoding=enc)
            except UnicodeDecodeError:
                continue
        raise ValueError("CSV 인코딩을 인식하지 못했습니다. UTF-8 또는 CP949(엑셀 기본)로 저장해 주세요.")
    if name.endswith(".xlsx"):
        return pd.read_excel(io.BytesIO(contents), dtype=str, keep_default_na=False)
    raise ValueError("지원하지 않는 파일 형식입니다. .xlsx 또는 .csv 파일을 올려 주세요.")


def parse_students(df: pd.DataFrame) -> Tuple[List[dict], List[str]]:
    df.columns = [str(c).strip() for c in df.columns]
    df = df.rename(columns={c: "GPA" for c in df.columns if c.lower() == "gpa"})
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(f"필수 열이 없습니다: {', '.join(missing)} (화면의 '양식 샘플 다운로드'를 참고하세요)")
    for c in OPTIONAL_COLS:
        if c not in df.columns:
            df[c] = ""

    students, warnings, seen = [], [], set()
    for idx, row in df.iterrows():
        line = idx + 2  # 헤더가 1행
        sid = str(row["학번"]).strip()
        if not sid and not str(row["이름"]).strip():
            continue  # 완전 빈 행
        if sid in seen:
            warnings.append(f"{line}행: 학번 {sid} 중복 → 제외")
            continue
        try:
            gpa = float(str(row["GPA"]).strip())
            birth = int(float(str(row["출생연도"]).strip()))
        except ValueError:
            warnings.append(f"{line}행: 학번 {sid}의 GPA 또는 출생연도가 숫자가 아님 → 제외")
            continue
        seen.add(sid)
        prefs_raw = [str(row[f"{i}지망병원"]).strip() for i in (1, 2, 3)]
        students.append({
            "student_id": sid,
            "name": str(row["이름"]).strip(),
            "gender": norm_gender(row["성별"]),
            "gpa": gpa,
            "birth_year": birth,
            "address": str(row["주소"]).strip(),
            "dong": dong_level(str(row["주소"]).strip()),
            "past_hospital": str(row["과거실습지"]).strip(),
            "prefs_raw": prefs_raw,
            "prefs_norm": [norm_h(p) for p in prefs_raw],
            "tb": tie_break(sid),
            "line": line,
        })
    if not students:
        raise ValueError("읽을 수 있는 학생 데이터가 없습니다.")
    if len(students) > MAX_STUDENTS:
        raise ValueError(f"학생 수가 너무 많습니다(최대 {MAX_STUDENTS}명).")
    return students, warnings


# ──────────────────────────────────────────────────────────────
# 점수 계산
# ──────────────────────────────────────────────────────────────
def pref_score(stu: dict, hosp_norm: str) -> Tuple[float, str]:
    for i, (p, sc) in enumerate(zip(stu["prefs_norm"], PREF_SCORES), start=1):
        if p and p == hosp_norm:
            return sc, f"{i}지망 일치"
    return NON_PREF_SCORE, "비지망(기본)"


TRAVEL_MIN_MIN, TRAVEL_MAX_MIN = 15, 90   # 통학 점수 환산 구간: 15분=100점 … 90분 이상=0점 (필요 시 조정)


def build_pair(stu: dict, hospital: str, dw: float, pw: float, real: bool = False) -> dict:
    info = lookup_cached(stu.get("dong", ""), hospital) if real else None
    if info:
        t, tr, w, source = info["t"], info["tr"], info["w"], info.get("src", "TMAP")
    else:
        (t, tr, w), source = estimate_transit(stu["address"], hospital), "추정"
    mfi = round(t + tr * 12.0 + w * 1.2, 1)
    dist_n = max(0.0, min(100.0, 100 - ((t - TRAVEL_MIN_MIN) / (TRAVEL_MAX_MIN - TRAVEL_MIN_MIN) * 100)))
    ps, pstat = pref_score(stu, norm_h(hospital))
    return {
        "travel_time_minutes": t,
        "transfers": tr,
        "walk_time_minutes": w,
        "fatigue_index": mfi,
        "pref_status": pstat,
        "source": source,
        "dist_score": round(dist_n, 1),
        "pref_score": ps,
        "composite_score": round(dist_n * dw + ps * pw, 1),
        "satisfaction_index": satisfaction_index(mfi),
    }


def static_ineligible_reasons(stu: dict, gender_criteria: str, min_gpa: Optional[float],
                              birth_year_after: Optional[int]) -> List[str]:
    reasons = []
    if gender_criteria == "남성만" and stu["gender"] != "남":
        reasons.append("성별 조건 불일치")
    elif gender_criteria == "여성만" and stu["gender"] != "여":
        reasons.append("성별 조건 불일치")
    if min_gpa is not None and stu["gpa"] < min_gpa:
        reasons.append("성적 미달")
    if birth_year_after is not None and stu["birth_year"] < birth_year_after:
        reasons.append("출생연도 기준 미충족")
    return reasons


def has_past(stu: dict, hospital: str) -> bool:
    return norm_h(hospital) in norm_h(stu["past_hospital"])


def build_alts(pair_map: Dict[str, dict], allowed: List[str], assigned_h: str) -> Tuple[List[dict], bool, str]:
    """학생이 다른 병원으로 갔다면 어땠는지(상위 3곳)와, 배정 병원이 그 학생의 최고점 병원인지."""
    ranked = sorted(((h, pair_map[h]) for h in allowed), key=lambda kv: -kv[1]["composite_score"])
    best_h, best_p = ranked[0]
    others = [{
        "hospital": h, "composite_score": p["composite_score"],
        "travel_time_minutes": p["travel_time_minutes"], "preference_status": p["pref_status"],
    } for h, p in ranked if h != assigned_h][:3]
    return others, pair_map[assigned_h]["composite_score"] >= best_p["composite_score"], best_h


def pair_stats(pairs: List[dict]) -> Optional[dict]:
    """pair 딕셔너리 목록의 평균 통학시간·지망 충족률·평균 종합점수."""
    n = len(pairs)
    if not n:
        return None
    return {
        "avg_travel_minutes": round(sum(p["travel_time_minutes"] for p in pairs) / n, 1),
        "pref_match_rate": round(sum(1 for p in pairs if "일치" in p["pref_status"]) / n * 100, 1),
        "avg_composite_score": round(sum(p["composite_score"] for p in pairs) / n, 1),
    }


def make_comparison(ours: Optional[dict], base: Optional[dict], label: str) -> Optional[dict]:
    if not ours or not base:
        return None
    saved = round(base["avg_travel_minutes"] - ours["avg_travel_minutes"], 1)
    return {
        "baseline_label": label,
        "ours": ours,
        "baseline": base,
        "travel_saved_minutes": saved,
        "travel_saved_pct": round(saved / base["avg_travel_minutes"] * 100, 1) if base["avg_travel_minutes"] else 0.0,
        "pref_rate_delta": round(ours["pref_match_rate"] - base["pref_match_rate"], 1),
        "score_delta": round(ours["avg_composite_score"] - base["avg_composite_score"], 1),
    }


def random_baseline(eligible: List[dict], pairs: List[Dict[str, dict]], allowed: List[List[str]],
                    valid: List[str], cap: int) -> Optional[dict]:
    """기준선: 정원은 지키되 학생을 무작위로 병원에 넣는 방식의 평균(고정 시드, 반복 시뮬레이션)."""
    n = len(eligible)
    if not n:
        return None
    rng = np.random.default_rng(20260101)
    trials = 100 if n <= 500 else 20
    picked: List[dict] = []
    for _ in range(trials):
        remaining = {h: cap for h in valid}
        for i in rng.permutation(n):
            opts = [h for h in allowed[i] if remaining[h] > 0]
            if not opts:
                continue
            h = opts[int(rng.integers(len(opts)))]
            remaining[h] -= 1
            picked.append(pairs[i][h])
    return pair_stats(picked)


def result_row(stu: dict, pair: dict, hospital: str, rank: Optional[int],
               alts: Optional[List[dict]] = None, is_top: Optional[bool] = None, best_h: str = "") -> dict:
    return {
        "dist_score": pair["dist_score"],
        "pref_score": pair["pref_score"],
        "transit_source": pair["source"],
        "alternatives": alts or [],
        "is_top_choice": is_top,
        "best_hospital": best_h,
        "rank": rank,
        "hospital": hospital,
        "student_id": stu["student_id"],
        "name": stu["name"],
        "gender": stu["gender"],
        "gpa": stu["gpa"],
        "birth_year": stu["birth_year"],
        "address": stu["address"],
        "preference_status": pair["pref_status"],
        "travel_time_minutes": pair["travel_time_minutes"],
        "transfers": pair["transfers"],
        "fatigue_index": pair["fatigue_index"],
        "composite_score": pair["composite_score"],
        "satisfaction_index": pair["satisfaction_index"],
        "report": (f"{stu['name']} 학생은 {hospital}에 배정되었습니다. "
                   f"(통학 약 {pair['travel_time_minutes']}분, 지망 반영: {pair['pref_status']}, "
                   f"종합 {pair['composite_score']}점)"),
    }


def unassigned_row(stu: dict, status: str, reason: str, hospital: str = "") -> dict:
    return {
        "student_id": stu["student_id"], "name": stu["name"], "gender": stu["gender"],
        "gpa": stu["gpa"], "birth_year": stu["birth_year"], "address": stu["address"],
        "status": status, "reason": reason, "hospital": hospital,
    }


# ──────────────────────────────────────────────────────────────
# 배정 핵심 로직 (웹 프레임워크와 분리: 단독 테스트 가능)
# ──────────────────────────────────────────────────────────────
def run_assignment(
    df: pd.DataFrame,
    mode: str,
    target_hospital: str,
    grade: str,
    semester: str,
    subject: str,
    start_date: str = "",
    end_date: str = "",
    max_period_capacity: int = 50,
    max_hospital_capacity: int = 10,
    gender_criteria: str = "무관",
    min_gpa: Optional[float] = None,
    birth_year_after: Optional[int] = None,
    exclude_past_hospital: bool = False,
    distance_weight: float = 0.7,
    preference_weight: float = 0.3,
    use_transit_api: bool = False,
) -> dict:
    valid = hospitals_for(grade, semester, subject)
    if not valid:
        raise ValueError("선택한 학년/학기/교과목에 해당하는 병원이 없습니다.")
    valid_norm = {norm_h(h) for h in valid}

    total_w = distance_weight + preference_weight
    if total_w <= 0:
        distance_weight, preference_weight, total_w = 0.7, 0.3, 1.0
    dw, pw = distance_weight / total_w, preference_weight / total_w

    cap = max(1, min(int(max_hospital_capacity), int(max_period_capacity)))

    students, warnings = parse_students(df)

    # 지망 병원명이 병원 목록과 안 맞는 경우 안내 (오타 투명성)
    bad = []
    for s in students:
        for i, (raw, nm) in enumerate(zip(s["prefs_raw"], s["prefs_norm"]), start=1):
            if raw and nm not in valid_norm:
                bad.append(f"{s['line']}행 {s['name']}: {i}지망 '{raw}'은(는) 이번 교과목 병원 목록에 없어 지망 점수에서 제외")
    warnings.extend(bad[:15])
    if len(bad) > 15:
        warnings.append(f"…지망 병원명 불일치 {len(bad) - 15}건 더 있음")

    eligible, unassigned = [], []
    for s in students:
        reasons = static_ineligible_reasons(s, gender_criteria, min_gpa, birth_year_after)
        if reasons:
            unassigned.append(unassigned_row(s, "자격미달", ", ".join(reasons)))
        else:
            eligible.append(s)

    real = bool(use_transit_api and TRANSIT_AVAILABLE)
    if real:
        warnings.extend(prefetch_transit(eligible, [target_hospital] if mode == "single" else valid))

    assigned: List[dict] = []

    if mode == "single":
        if not target_hospital or norm_h(target_hospital) not in valid_norm:
            raise ValueError("배정 대상 병원을 올바르게 선택해 주세요.")
        pool = []
        for s in eligible:
            if exclude_past_hospital and has_past(s, target_hospital):
                unassigned.append(unassigned_row(s, "자격미달", "과거 실습 이력 중복", target_hospital))
                continue
            pool.append((s, build_pair(s, target_hospital, dw, pw, real)))
        upload_order = list(pool)  # 학번순(선착순) 기준선용: 업로드 순서 보존
        pool.sort(key=lambda x: (-x[1]["composite_score"], x[0]["tb"]))
        for rank, (s, pair) in enumerate(pool, start=1):
            if rank <= cap:
                assigned.append(result_row(s, pair, target_hospital, rank))
            else:
                unassigned.append(unassigned_row(
                    s, "대기",
                    f"{target_hospital} 정원({cap}명) 마감 — 종합점수 {pair['composite_score']}점, 대기 {rank - cap}번",
                    target_hospital))
        method = f"단일 병원 순위 배정 (거리 {round(dw*100)}% · 희망도 {round(pw*100)}%, 정원 {cap}명)"
        target_label = target_hospital
        slots_total = cap
        hospital_cap_map = {target_hospital: cap}
        comparison = make_comparison(
            pair_stats([p for _, p in pool[:cap]]),
            pair_stats([p for _, p in upload_order[:cap]]),
            "학번순(선착순) 배정 — 명단 순서대로 정원까지")
    else:
        n = len(eligible)
        cap_eff = max(1, min(cap, n)) if n else cap
        slot_hosp = [h for h in valid for _ in range(cap_eff)]
        slots_total = len(slot_hosp)
        hospital_cap_map = {h: cap_eff for h in valid}
        comparison = None
        if n:
            pairs = [{h: build_pair(s, h, dw, pw, real) for h in valid} for s in eligible]
            allowed = [[h for h in valid if not (exclude_past_hospital and has_past(s, h))] for s in eligible]
            # 학생 x 병원 비용표를 만든 뒤, 병원별 정원만큼 열을 복제해 슬롯 행렬로 확장
            base_cost = np.full((n, len(valid)), BIG_COST)
            for i, s in enumerate(eligible):
                for hi, h in enumerate(valid):
                    if h in allowed[i]:
                        base_cost[i, hi] = 100.0 - pairs[i][h]["composite_score"] + s["tb"] * 1e-6
            slot_idx = np.repeat(np.arange(len(valid)), cap_eff)
            cost = base_cost[:, slot_idx]
            rows, cols = linear_sum_assignment(cost)
            matched = {}
            for i, j in zip(rows, cols):
                if cost[i, j] < BIG_COST / 2:
                    matched[i] = slot_hosp[j]
            by_hosp: Dict[str, List[tuple]] = {}
            for i, s in enumerate(eligible):
                if i in matched:
                    h = matched[i]
                    by_hosp.setdefault(h, []).append((s, pairs[i][h], i))
                else:
                    if slots_total < n:
                        unassigned.append(unassigned_row(s, "대기", "전체 병원 정원 마감 (총 정원 < 대상 인원)"))
                    else:
                        unassigned.append(unassigned_row(s, "대기", "배정 가능한 병원 없음 (과거 실습 이력 제한 등)"))
            for h, lst in by_hosp.items():
                lst.sort(key=lambda x: (-x[1]["composite_score"], x[0]["tb"]))
                for rank, (s, pair, i) in enumerate(lst, start=1):
                    alts, is_top, best_h = build_alts(pairs[i], allowed[i], h)
                    assigned.append(result_row(s, pair, h, rank, alts, is_top, best_h))
            assigned.sort(key=lambda r: (r["hospital"], r["rank"]))
            comparison = make_comparison(
                pair_stats([pairs[i][h] for i, h in matched.items()]),
                random_baseline(eligible, pairs, allowed, valid, cap_eff),
                "무작위 배정 — 정원은 지키되 학생을 임의로 배치 (100회 시뮬레이션 평균)")
        method = (f"전체 병원 동시 배정 · SciPy 헝가리안 전역 최적화 "
                  f"(거리 {round(dw*100)}% · 희망도 {round(pw*100)}%, 병원당 정원 {cap_eff}명, 병원 {len(valid)}곳)")
        target_label = "전체 병원"

    # 요약 통계
    na = len(assigned)
    avg_travel = round(sum(r["travel_time_minutes"] for r in assigned) / na, 1) if na else None
    avg_score = round(sum(r["composite_score"] for r in assigned) / na, 1) if na else None
    pref_hit = sum(1 for r in assigned if "일치" in r["preference_status"])
    pref_rate = round(pref_hit / na * 100, 1) if na else None

    hosp_summary = []
    for h, c in hospital_cap_map.items():
        rs = [r for r in assigned if r["hospital"] == h]
        if mode == "single" or rs:
            hosp_summary.append({
                "hospital": h, "assigned": len(rs), "capacity": c,
                "avg_travel": round(sum(r["travel_time_minutes"] for r in rs) / len(rs), 1) if rs else None,
                "pref_matched": sum(1 for r in rs if "일치" in r["preference_status"]),
            })

    return {
        "status": "success",
        "mode": mode,
        "target_hospital": target_label,
        "period": {"start": start_date, "end": end_date},
        "total_students": len(students),
        "eligible_count": len(eligible),
        "assigned_count": na,
        "waitlist_count": sum(1 for u in unassigned if u["status"] == "대기"),
        "ineligible_count": sum(1 for u in unassigned if u["status"] == "자격미달"),
        "total_slots": slots_total,
        "avg_travel_minutes": avg_travel,
        "avg_composite_score": avg_score,
        "pref_match_rate": pref_rate,
        "optimization_method": method,
        "weights": {"distance": round(dw * 100), "preference": round(pw * 100)},
        "transit": {
            "requested": bool(use_transit_api), "available": TRANSIT_AVAILABLE, "real": real, "mode": TRANSIT_MODE,
            "real_ratio": round(sum(1 for r in assigned if r["transit_source"] != "추정") / na * 100, 1) if na else 0.0,
        },
        "comparison": comparison,
        "hospital_summary": hosp_summary,
        "warnings": warnings,
        "results": assigned,
        "unassigned": unassigned,
    }


# ──────────────────────────────────────────────────────────────
# API
# ──────────────────────────────────────────────────────────────
class PasswordVerifyRequest(BaseModel):
    password: str


def _require_token(token: Optional[str]):
    if AUTH_REQUIRED and not _token_ok(token):
        raise HTTPException(status_code=401, detail="인증이 만료되었습니다. 다시 접속해 주세요.")


@app.post("/api/v1/verify-password")
async def verify_password(payload: PasswordVerifyRequest):
    if hmac.compare_digest(payload.password.encode("utf-8"), APP_PASSWORD.encode("utf-8")):
        return {"status": "success", "token": _make_token()}
    raise HTTPException(status_code=401, detail="암호가 올바르지 않습니다.")


@app.get("/api/v1/config")
async def get_config():
    missing = 0
    if TRANSIT_AVAILABLE:
        addr = load_hospital_addresses()
        missing = sum(1 for h in all_hospital_names() if norm_h(h) not in addr)
    return {"transit_available": TRANSIT_AVAILABLE, "transit_mode": TRANSIT_MODE, "hospitals_missing_address": missing, "auth_required": AUTH_REQUIRED, "version": APP_VERSION}


@app.get("/api/v1/curriculum")
async def get_curriculum():
    return {g: {s: list(subs.keys()) for s, subs in sems.items()} for g, sems in MASTER_HOSPITAL_DB.items()}


@app.get("/api/v1/hospitals")
async def get_hospitals(grade: str, semester: str, subject: str):
    return {"hospitals": hospitals_for(grade, semester, subject)}


@app.get("/api/v1/template")
async def download_template():
    sample = pd.DataFrame([
        {"학번": "20240001", "이름": "홍길동", "성별": "남", "GPA": 3.8, "출생연도": 2004,
         "주소": "경기도 안산시 단원구 고잔동", "과거실습지": "",
         "1지망병원": "고려대학교안산병원", "2지망병원": "가톨릭대학교 부천성모병원", "3지망병원": "한림대학교성심병원"},
        {"학번": "20240002", "이름": "김나은", "성별": "여", "GPA": 4.1, "출생연도": 2003,
         "주소": "경기도 안산시 상록구 본오동", "과거실습지": "고려대학교안산병원",
         "1지망병원": "순천향대학교부천병원", "2지망병원": "", "3지망병원": ""},
    ])
    buf = io.BytesIO()
    sample.to_excel(buf, index=False, sheet_name="학생명단")
    buf.seek(0)
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename*=UTF-8''%EC%8B%A4%EC%8A%B5%EC%A7%80%EB%B0%B0%EC%A0%95_%EC%96%91%EC%8B%9D.xlsx"},
    )


SAMPLE_SURNAMES = ["김", "이", "박", "최", "정", "강", "조", "윤", "장", "임", "한", "오", "서", "신", "권", "황"]
SAMPLE_FEMALE = ["서연", "지우", "하윤", "서윤", "지민", "수아", "하은", "지안", "유진", "채원", "다은", "소연", "나은", "예린", "민서", "서현", "아린", "예은"]
SAMPLE_MALE = ["민준", "도윤", "시우", "예준", "주원", "지호", "현우", "건우", "선우", "준서", "유찬", "태윤"]
SAMPLE_ADDRS = (
    ["경기도 안산시 단원구 고잔동", "경기도 안산시 단원구 와동", "경기도 안산시 단원구 선부동", "경기도 안산시 단원구 초지동",
     "경기도 안산시 단원구 원곡동", "경기도 안산시 상록구 본오동", "경기도 안산시 상록구 사동", "경기도 안산시 상록구 일동",
     "경기도 안산시 상록구 성포동"] * 2
    + ["경기도 군포시 산본동", "경기도 군포시 금정동", "경기도 수원시 팔달구", "경기도 수원시 영통구",
       "경기도 부천시 중동", "경기도 안양시 동안구"]
)


def make_sample_df(grade: str, semester: str, subject: str, n: int = 60, seed: int = 2026) -> pd.DataFrame:
    """데모용 가상 학생 명단 (고정 시드 → 항상 같은 데이터). 일부 병원에 지망이 몰리도록 만들어 경쟁 상황을 재현한다."""
    rnd = random.Random(seed)
    hosp = hospitals_for(grade, semester, subject) or hospitals_for("3학년", "1학기", "성인1")
    all_hosp = sorted({h for g in MASTER_HOSPITAL_DB.values() for s in g.values() for lst in s.values() for h in lst})
    order = hosp[:]
    rnd.shuffle(order)
    weights = [max(1.0, 3.5 - 0.6 * i) for i in range(len(order))]  # 앞쪽 병원일수록 인기
    year = "2024" if grade == "3학년" else "2023"
    rows = []
    for i in range(n):
        female = rnd.random() < 0.78
        name = rnd.choice(SAMPLE_SURNAMES) + rnd.choice(SAMPLE_FEMALE if female else SAMPLE_MALE)
        picks: List[str] = []
        while len(picks) < min(3, len(order)):
            h = rnd.choices(order, weights)[0]
            if h not in picks:
                picks.append(h)
        picks += [""] * (3 - len(picks))
        rows.append({
            "학번": f"{year}{i + 1:04d}", "이름": name, "성별": "여" if female else "남",
            "GPA": round(min(4.5, max(2.6, rnd.gauss(3.6, 0.35))), 2),
            "출생연도": rnd.choice([2002, 2003, 2003, 2004, 2004, 2005]),
            "주소": rnd.choice(SAMPLE_ADDRS),
            "과거실습지": rnd.choice(all_hosp) if rnd.random() < 0.3 else "",
            "1지망병원": picks[0], "2지망병원": picks[1], "3지망병원": picks[2],
        })
    return pd.DataFrame(rows)


def _df_to_xlsx_bytes(sheets: Dict[str, pd.DataFrame]) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for name, frame in sheets.items():
            frame.to_excel(writer, index=False, sheet_name=name[:31])
    return buf.getvalue()


def _safe_cell(v):
    """엑셀 수식 주입 방지: = + - @ 로 시작하는 문자열은 앞에 ' 를 붙여 텍스트로 저장."""
    if isinstance(v, str) and v[:1] in ("=", "+", "-", "@"):
        return "'" + v
    return v


def build_export_xlsx(data: dict) -> bytes:
    res = [{
        "병원": r.get("hospital"), "순위": r.get("rank"), "학번": r.get("student_id"), "이름": r.get("name"),
        "성별": r.get("gender"), "GPA": r.get("gpa"), "주소": r.get("address"), "지망 반영": r.get("preference_status"),
        "통학시간(분)": r.get("travel_time_minutes"), "통학시간 출처": r.get("transit_source"), "환승": r.get("transfers"),
        "통학 점수": r.get("dist_score"), "지망 점수": r.get("pref_score"),
        "종합 매칭점수": r.get("composite_score"), "예상 만족도 지수": r.get("satisfaction_index"),
    } for r in data.get("results", [])]
    un = [{
        "구분": u.get("status"), "학번": u.get("student_id"), "이름": u.get("name"), "GPA": u.get("gpa"),
        "주소": u.get("address"), "사유": u.get("reason"),
    } for u in data.get("unassigned", [])]
    hs = [{
        "병원": h.get("hospital"), "배정": h.get("assigned"), "정원": h.get("capacity"),
        "평균 통학시간(분)": h.get("avg_travel"), "지망 일치 인원": h.get("pref_matched"),
    } for h in data.get("hospital_summary", [])]
    cmp_ = data.get("comparison") or {}
    summary = [{"항목": "배정 방식", "값": data.get("optimization_method")},
               {"항목": "전체 학생", "값": data.get("total_students")},
               {"항목": "배정 인원", "값": data.get("assigned_count")},
               {"항목": "평균 통학시간(분, 추정)", "값": data.get("avg_travel_minutes")},
               {"항목": "지망 충족률(%)", "값": data.get("pref_match_rate")}]
    if cmp_:
        summary += [{"항목": "비교 기준", "값": cmp_.get("baseline_label")},
                    {"항목": "기준선 평균 통학시간(분)", "값": cmp_["baseline"]["avg_travel_minutes"]},
                    {"항목": "기준선 지망 충족률(%)", "값": cmp_["baseline"]["pref_match_rate"]},
                    {"항목": "통학시간 단축(분)", "값": cmp_.get("travel_saved_minutes")},
                    {"항목": "지망 충족률 변화(%p)", "값": cmp_.get("pref_rate_delta")}]
    frames = {"배정결과": pd.DataFrame(res), "미배정_사유": pd.DataFrame(un),
              "병원별현황": pd.DataFrame(hs), "요약": pd.DataFrame(summary)}
    frames = {k: v.apply(lambda col: col.map(_safe_cell)) for k, v in frames.items()}
    return _df_to_xlsx_bytes(frames)


@app.get("/api/v1/sample-data")
async def download_sample(grade: str = "3학년", semester: str = "1학기", subject: str = "성인1"):
    data = _df_to_xlsx_bytes({"학생명단": make_sample_df(grade, semester, subject)})
    return StreamingResponse(
        io.BytesIO(data),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=sample_students.xlsx"},
    )


@app.post("/api/v1/export-xlsx")
async def export_xlsx(payload: dict = Body(...), x_auth_token: Optional[str] = Header(None)):
    _require_token(x_auth_token)
    try:
        data = build_export_xlsx(payload)
    except Exception:
        raise HTTPException(status_code=400, detail="내보내기에 실패했습니다.")
    return StreamingResponse(
        io.BytesIO(data),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=result.xlsx"},
    )


@app.post("/api/v1/assign-file")
async def assign_hospital_from_file(
    mode: str = Form("global"),
    target_hospital: str = Form(""),
    grade: str = Form("3학년"),
    semester: str = Form("1학기"),
    subject: str = Form("성인1"),
    start_date: str = Form(""),
    end_date: str = Form(""),
    max_period_capacity: int = Form(50),
    max_hospital_capacity: int = Form(10),
    gender_criteria: str = Form("무관"),
    min_gpa: Optional[float] = Form(None),
    birth_year_after: Optional[int] = Form(None),
    exclude_past_hospital: bool = Form(False),
    distance_weight: float = Form(0.7),
    preference_weight: float = Form(0.3),
    use_transit_api: bool = Form(False),
    file: UploadFile = File(...),
    x_auth_token: Optional[str] = Header(None),
):
    _require_token(x_auth_token)
    contents = await file.read()
    if len(contents) > MAX_UPLOAD_BYTES:
        raise HTTPException(status_code=413, detail="파일이 너무 큽니다(최대 5MB).")
    try:
        df = read_table(contents, file.filename)
        # 외부 API 조회가 서버 전체를 막지 않도록 별도 스레드에서 실행
        return await run_in_threadpool(
            run_assignment,
            df, "single" if mode == "single" else "global", target_hospital, grade, semester, subject,
            start_date, end_date, max_period_capacity, max_hospital_capacity, gender_criteria,
            min_gpa, birth_year_after, exclude_past_hospital, distance_weight, preference_weight,
            use_transit_api=use_transit_api,
        )
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception:
        raise HTTPException(status_code=400, detail="데이터 처리에 실패했습니다. 양식 샘플과 비교해 파일을 확인해 주세요.")
    # 업로드 파일은 메모리에서만 처리하며 서버에 저장하지 않는다.


# ──────────────────────────────────────────────────────────────
# 화면
# ──────────────────────────────────────────────────────────────
HTML_PAGE = r"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>간호학과 스마트 실습지 최적 배정 시스템</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css" media="print" onload="this.media='all'">
<style>
  :root {
    --navy:#0f172a; --blue:#1e3a8a; --teal:#0d9488; --teal-d:#0f766e; --teal-l:#ccfbf1; --amber:#f59e0b;
    --bg:#f4f7fb; --card:#ffffff; --line:#e2e8f0; --muted:#64748b; --text:#1e293b; --good:#047857; --bad:#dc2626; --warn:#b45309;
    --shadow:0 1px 2px rgba(15,23,42,.04), 0 8px 24px rgba(15,23,42,.05);
  }
  * { box-sizing:border-box; }
  html { -webkit-text-size-adjust:100%; }
  body { margin:0; font-family:'Pretendard','Apple SD Gothic Neo','Malgun Gothic','Noto Sans KR',system-ui,sans-serif; background:var(--bg); color:var(--text); line-height:1.5; }
  .wrap { max-width:1180px; margin:0 auto; padding:0 20px; }
  button, input, select { font:inherit; color:inherit; }

  /* 상단 */
  .top { background:linear-gradient(135deg,var(--navy),var(--blue)); color:#fff; }
  .top-in { display:flex; align-items:center; justify-content:space-between; gap:12px; padding-top:14px; padding-bottom:14px; }
  .brand { font-size:17px; font-weight:700; letter-spacing:-.2px; }
  .brand span { font-weight:500; opacity:.8; margin-left:8px; }
  .ver { font-size:12px; background:rgba(255,255,255,.14); padding:4px 10px; border-radius:99px; white-space:nowrap; }

  /* 단계 표시 */
  .stepper { list-style:none; display:flex; gap:8px; padding:0; margin:22px 0 18px; }
  .stepper li { flex:1; display:flex; align-items:center; gap:10px; padding:12px 14px; background:var(--card); border:1px solid var(--line); border-radius:12px; color:var(--muted); font-weight:600; font-size:14px; cursor:default; }
  .stepper li i { font-style:normal; width:26px; height:26px; flex:none; border-radius:50%; background:#e2e8f0; color:#475569; display:grid; place-items:center; font-size:13px; font-weight:700; }
  .stepper li.done { cursor:pointer; color:var(--text); }
  .stepper li.done i { background:var(--teal); color:#fff; }
  .stepper li.active { border-color:var(--teal); color:var(--navy); box-shadow:0 0 0 3px rgba(13,148,136,.12); }
  .stepper li.active i { background:var(--navy); color:#fff; }

  /* 카드·공통 */
  .card { background:var(--card); border-radius:16px; box-shadow:var(--shadow); padding:26px; }
  .card h2 { margin:0 0 4px; font-size:19px; letter-spacing:-.3px; color:var(--navy); }
  .lead { margin:0 0 20px; color:var(--muted); font-size:14px; }
  .field > label, .lbl { display:block; font-size:13px; font-weight:700; color:#475569; margin-bottom:6px; }
  select, input[type=number], input[type=date], input[type=search], input[type=password] { width:100%; padding:11px 12px; border:1.5px solid var(--line); border-radius:10px; background:#fff; }
  select:focus, input:focus { outline:none; border-color:var(--teal); box-shadow:0 0 0 3px rgba(13,148,136,.15); }
  .grid { display:grid; gap:14px; }
  .g2 { grid-template-columns:1fr 1fr; } .g3 { grid-template-columns:repeat(3,1fr); }
  .hint { font-size:12px; color:var(--muted); margin-top:6px; }
  .btn { border:none; border-radius:12px; padding:13px 22px; font-weight:700; cursor:pointer; transition:transform .12s, filter .12s; display:inline-flex; align-items:center; justify-content:center; gap:8px; text-decoration:none; }
  .btn:active { transform:translateY(1px); }
  .btn-primary { background:linear-gradient(135deg,var(--teal),var(--teal-d)); color:#fff; box-shadow:0 6px 16px rgba(13,148,136,.28); }
  .btn-primary:hover { filter:brightness(1.05); }
  .btn-primary:disabled { opacity:.6; cursor:wait; }
  .btn-ghost { background:#fff; color:var(--blue); border:1.5px solid var(--line); }
  .btn-ghost:hover { border-color:var(--blue); }
  .btn-sm { padding:9px 14px; font-size:13px; border-radius:10px; }
  .actions { display:flex; justify-content:space-between; gap:10px; margin-top:24px; flex-wrap:wrap; }

  /* 모드 선택 */
  .mode { display:block; border:2px solid var(--line); border-radius:14px; padding:16px; cursor:pointer; transition:border-color .15s, background .15s; height:100%; }
  .mode input { position:absolute; opacity:0; pointer-events:none; }
  .mode b { display:block; font-size:15px; color:var(--navy); }
  .mode span { font-size:13px; color:var(--muted); }
  .mode.on { border-color:var(--teal); background:#f0fdfa; }
  .mode:focus-within { box-shadow:0 0 0 3px rgba(13,148,136,.2); }

  details.adv { margin-top:20px; border:1.5px solid var(--line); border-radius:12px; }
  details.adv > summary { cursor:pointer; font-weight:700; padding:14px 16px; list-style:none; color:var(--navy); }
  details.adv > summary::after { content:'＋'; float:right; color:var(--muted); }
  details.adv[open] > summary::after { content:'－'; }
  details.adv .adv-body { padding:4px 16px 18px; border-top:1px solid var(--line); }
  .sec { margin-top:16px; }

  /* 업로드 */
  .drop { border:2px dashed #94a3b8; border-radius:16px; background:#f8fafc; padding:34px 20px; text-align:center; transition:all .15s; }
  .drop.over { border-color:var(--teal); background:#f0fdfa; }
  .drop .big { font-size:17px; font-weight:700; color:var(--navy); margin-bottom:4px; }
  .drop .file-ok { display:none; margin-top:14px; font-weight:700; color:var(--good); }
  .chips { display:flex; flex-wrap:wrap; gap:8px; margin-bottom:18px; }
  .chip { background:#eef2ff; color:var(--blue); padding:5px 12px; border-radius:99px; font-size:13px; font-weight:600; }
  .or { display:flex; align-items:center; gap:12px; color:var(--muted); font-size:13px; margin:18px 0; }
  .or::before, .or::after { content:''; height:1px; background:var(--line); flex:1; }
  .tools { display:flex; flex-wrap:wrap; gap:10px; justify-content:center; }

  /* 결과: 히어로 */
  .res-top { display:flex; justify-content:space-between; align-items:center; gap:10px; flex-wrap:wrap; margin-bottom:14px; }
  .res-top .meta { font-size:14px; color:var(--muted); }
  .res-top .meta b { color:var(--navy); }
  .hero { background:linear-gradient(135deg,var(--navy),#1e3a8a); color:#fff; border-radius:20px; padding:26px; display:grid; grid-template-columns:1fr 1.25fr; gap:28px; box-shadow:var(--shadow); }
  .hero .eyebrow { font-size:12px; letter-spacing:.4px; opacity:.75; font-weight:600; margin-bottom:10px; }
  .bignums { display:flex; gap:26px; flex-wrap:wrap; }
  .bignum .n { font-size:40px; font-weight:800; line-height:1.05; letter-spacing:-1px; }
  .bignum .n small { font-size:16px; font-weight:700; margin-left:2px; }
  .bignum .t { font-size:13px; opacity:.8; margin-top:4px; }
  .pos { color:#5eead4; } .neg { color:#fda4af; }
  .cmp-row { margin-bottom:14px; }
  .cmp-row:last-of-type { margin-bottom:6px; }
  .cmp-row .cl { font-size:12px; opacity:.8; margin-bottom:5px; font-weight:600; }
  .pbar { display:flex; align-items:center; gap:10px; margin-top:4px; font-size:12px; }
  .pbar .tag { width:34px; opacity:.75; flex:none; }
  .pbar .track { flex:1; height:12px; border-radius:6px; background:rgba(255,255,255,.1); overflow:hidden; }
  .pbar .fill { height:100%; border-radius:6px; }
  .pbar .fill.base { background:#94a3b8; } .pbar .fill.ours { background:#2dd4bf; }
  .pbar .val { width:64px; text-align:right; font-weight:700; flex:none; }
  .cmp-note { font-size:12px; opacity:.7; margin-top:10px; }

  /* KPI */
  .kpis { display:grid; grid-template-columns:repeat(4,1fr); gap:14px; margin:16px 0; }
  .kpi { background:var(--card); border-radius:14px; padding:16px 18px; box-shadow:var(--shadow); border-left:4px solid var(--teal); }
  .kpi .k { font-size:12px; color:var(--muted); font-weight:600; }
  .kpi .v { font-size:26px; font-weight:800; color:var(--navy); letter-spacing:-.5px; }
  .kpi .v small { font-size:13px; font-weight:600; color:var(--muted); margin-left:3px; }

  /* 차트 */
  .charts { display:grid; grid-template-columns:repeat(auto-fit,minmax(300px,1fr)); gap:14px; margin-bottom:16px; }
  .chart { background:var(--card); border-radius:16px; box-shadow:var(--shadow); padding:20px; }
  .chart h3 { margin:0 0 14px; font-size:15px; color:var(--navy); }
  .hrow { display:grid; grid-template-columns:minmax(90px,38%) 1fr 44px; align-items:center; gap:8px; margin-bottom:9px; font-size:12px; }
  .hrow .nm { overflow:hidden; text-overflow:ellipsis; white-space:nowrap; color:#334155; }
  .hrow .tr { height:14px; background:#eef2f7; border-radius:7px; overflow:hidden; }
  .hrow .fl { height:100%; background:var(--teal); border-radius:7px; }
  .hrow .fl.full { background:var(--blue); }
  .hrow .ct { text-align:right; font-weight:700; color:var(--navy); }
  .vbars { display:flex; align-items:flex-end; gap:8px; height:150px; padding-top:6px; }
  .vb { flex:1; display:flex; flex-direction:column; justify-content:flex-end; align-items:center; height:100%; font-size:12px; }
  .vb .bar { width:100%; background:var(--teal); border-radius:6px 6px 0 0; min-height:2px; }
  .vb .n { font-weight:700; color:var(--navy); margin-bottom:3px; }
  .vlabels { display:flex; gap:8px; font-size:11px; color:var(--muted); margin-top:6px; }
  .vlabels span { flex:1; text-align:center; }
  .donut-wrap { display:flex; align-items:center; gap:22px; flex-wrap:wrap; justify-content:center; }
  .donut { width:150px; height:150px; border-radius:50%; position:relative; flex:none; }
  .donut-hole { position:absolute; inset:24px; background:#fff; border-radius:50%; display:flex; flex-direction:column; align-items:center; justify-content:center; }
  .donut-hole b { font-size:26px; color:var(--navy); line-height:1; }
  .donut-hole span { font-size:11px; color:var(--muted); margin-top:3px; }
  .legend { font-size:13px; } .legend div { display:flex; align-items:center; gap:8px; margin:5px 0; }
  .legend i { width:12px; height:12px; border-radius:3px; display:inline-block; }

  /* 표 */
  .alert { background:#fffbeb; border:1px solid #fde68a; color:#92400e; border-radius:12px; padding:12px 16px; margin-bottom:16px; font-size:13px; }
  .alert ul { margin:6px 0 0; padding-left:18px; }
  .tabs { display:flex; gap:4px; border-bottom:2px solid var(--line); padding:0 14px; flex-wrap:wrap; align-items:center; }
  .tab { border:none; background:none; padding:14px 16px; font-weight:700; color:var(--muted); border-bottom:3px solid transparent; margin-bottom:-2px; cursor:pointer; }
  .tab.on { color:var(--blue); border-bottom-color:var(--blue); }
  .tabs .search { margin-left:auto; width:210px; padding:8px 12px; }
  .pane { display:none; } .pane.on { display:block; }
  .tbl-wrap { overflow-x:auto; }
  table.tbl { width:100%; min-width:720px; border-collapse:collapse; font-size:14px; }
  .tbl td { white-space:nowrap; }
  .tbl th { background:#f8fafc; color:#475569; font-size:12px; text-align:center; padding:12px 10px; white-space:nowrap; border-bottom:1px solid var(--line); position:sticky; top:0; }
  .tbl td { padding:12px 10px; text-align:center; border-bottom:1px solid #f1f5f9; vertical-align:middle; }
  .tbl td.l, .tbl th.l { text-align:left; }
  .tbl tbody tr.click { cursor:pointer; } .tbl tbody tr.click:hover { background:#f0f9ff; }
  .tbl th[title] { cursor:help; text-decoration:underline dotted #94a3b8; text-underline-offset:3px; }
  .rk { display:inline-block; min-width:34px; padding:3px 8px; border-radius:99px; background:#fef3c7; color:#92400e; font-weight:700; font-size:12px; }
  .pill { display:inline-block; padding:3px 10px; border-radius:99px; font-size:12px; font-weight:700; white-space:nowrap; }
  .pill.hit { background:#ccfbf1; color:#115e59; } .pill.miss { background:#f1f5f9; color:#475569; }
  .pill.wait { background:#fef3c7; color:#92400e; } .pill.bad { background:#fee2e2; color:#991b1b; }
  .score { font-weight:800; color:var(--blue); }
  .empty { padding:36px; text-align:center; color:var(--muted); }
  .sub { color:var(--muted); font-size:12px; }

  /* 모달 */
  .modal { position:fixed; inset:0; background:rgba(15,23,42,.55); display:none; align-items:center; justify-content:center; z-index:50; padding:16px; }
  .modal.open { display:flex; }
  .sheet { background:#fff; border-radius:18px; width:100%; max-width:560px; max-height:92vh; overflow:auto; padding:26px; box-shadow:0 24px 60px rgba(0,0,0,.3); }
  .sheet h3 { margin:0; font-size:20px; color:var(--navy); }
  .sheet .x { float:right; border:none; background:#f1f5f9; width:34px; height:34px; border-radius:50%; font-size:18px; cursor:pointer; }
  .box { background:#f8fafc; border-radius:12px; padding:14px 16px; margin-top:14px; }
  .box h4 { margin:0 0 10px; font-size:13px; color:#475569; }
  .stack { display:flex; height:16px; border-radius:8px; overflow:hidden; background:#e2e8f0; margin:8px 0 10px; }
  .stack i { display:block; height:100%; }
  .eq { font-size:13px; display:grid; grid-template-columns:1fr auto; gap:4px 10px; }
  .eq b { text-align:right; }
  .alt { display:flex; justify-content:space-between; gap:10px; font-size:13px; padding:7px 0; border-bottom:1px solid #e2e8f0; }
  .alt:last-child { border-bottom:none; }

  /* 인증·토스트 */
  .auth { position:fixed; inset:0; background:radial-gradient(circle at 50% 30%,#1e293b 0%,#0f172a 100%); z-index:100; display:flex; align-items:center; justify-content:center; padding:16px; }
  .auth-card { background:rgba(30,41,59,.92); width:100%; max-width:400px; padding:38px 30px; border-radius:22px; border:1px solid rgba(255,255,255,.1); text-align:center; color:#fff; }
  .auth-card input { text-align:center; font-weight:700; margin:18px 0 8px; color:#111; background:#fff; }
  .auth-card input::placeholder { color:#6b7280; font-weight:500; }
  .toast { position:fixed; left:50%; bottom:26px; transform:translateX(-50%) translateY(20px); background:var(--navy); color:#fff; padding:12px 20px; border-radius:12px; font-size:14px; opacity:0; transition:all .25s; z-index:200; max-width:90vw; pointer-events:none; }
  .toast.show { opacity:1; transform:translateX(-50%); } .toast.err { background:#b91c1c; }

  @media (max-width:860px) {
    .hero { grid-template-columns:1fr; gap:20px; }
    .kpis { grid-template-columns:repeat(2,1fr); }
    .g3 { grid-template-columns:1fr; }
  }
  @media (max-width:600px) {
    .wrap { padding:0 14px; } .card { padding:20px 16px; }
    .stepper li { padding:10px 8px; font-size:12px; gap:6px; flex-direction:column; text-align:center; }
    .g2 { grid-template-columns:1fr; }
    .brand span { display:block; margin:0; font-size:12px; }
    .bignum .n { font-size:32px; }
    .tabs .search { width:100%; margin:8px 0; order:9; }
  }
  @media (prefers-reduced-motion:reduce) { * { transition:none !important; } }
</style>
</head>
<body>

<div id="auth" class="auth">
  <div class="auth-card">
    <h2 style="margin:0 0 4px;font-size:20px">보안 서버 인증</h2>
    <div style="opacity:.75;font-size:13px">간호학과 스마트 실습지 최적 배정 시스템</div>
    <input type="password" id="authPw" placeholder="접속 암호 입력" autocomplete="current-password" onkeyup="if(event.key==='Enter')verifyPassword()">
    <div id="authErr" style="color:#fbbf24;font-size:13px;min-height:20px;margin-bottom:8px"></div>
    <button class="btn btn-primary" style="width:100%" onclick="verifyPassword()">시스템 접속하기</button>
  </div>
</div>

<header class="top"><div class="wrap top-in">
  <div class="brand">🏥 로켓단<span>간호학과 실습지 최적 배정 시스템</span></div>
  <span class="ver">__APP_VERSION__</span>
</div></header>

<main class="wrap">
  <ol class="stepper" id="stepper">
    <li data-s="1" class="active"><i>1</i>조건 설정</li>
    <li data-s="2"><i>2</i>명단 업로드</li>
    <li data-s="3"><i>3</i>결과 확인</li>
  </ol>

  <!-- 1단계 -->
  <section id="step1" class="card step">
    <h2>어떤 실습을 배정할까요?</h2>
    <p class="lead">배정 방식과 교과목을 고르면 해당 교과목의 실습 병원이 자동으로 불러와집니다.</p>

    <div class="grid g2" style="margin-bottom:18px">
      <label class="mode on" id="modeG"><input type="radio" name="mode" value="global" checked onchange="onModeChange()">
        <b>전체 병원 동시 배정 <span class="pill hit" style="margin-left:4px">추천</span></b>
        <span>모든 학생과 모든 병원을 한 번에 최적화합니다. 학생 한 명은 병원 한 곳에만 배정됩니다.</span></label>
      <label class="mode" id="modeS"><input type="radio" name="mode" value="single" onchange="onModeChange()">
        <b>단일 병원 순위</b>
        <span>병원 하나를 골라 지원자의 종합점수 순위를 매기고 정원까지 배정합니다.</span></label>
    </div>

    <div class="grid g3">
      <div class="field"><label for="grade">학년</label>
        <select id="grade" onchange="updateSubjects()"><option>3학년</option><option>4학년</option></select></div>
      <div class="field"><label for="semester">학기</label>
        <select id="semester" onchange="updateSubjects()"><option>1학기</option><option>2학기</option></select></div>
      <div class="field"><label for="subject">실습 교과목</label>
        <select id="subject" onchange="updateHospitals()"></select></div>
    </div>
    <div class="field sec" id="hospWrap" style="display:none"><label for="hospital">배정 대상 병원 <span class="sub">(방중 = 방학중 운영)</span></label>
      <select id="hospital"></select></div>
    <div class="hint" id="hospCount"></div>

    <details class="adv">
      <summary>상세 설정 — 가중치 · 기간 · 정원 · 자격 조건</summary>
      <div class="adv-body">
        <div class="sec"><span class="lbl">배정 기준 가중치 (합계 100%)</span>
          <div class="grid g2">
            <div><span class="sub">통학 거리 (%)</span><input type="number" id="wDist" value="70" min="0" max="100" oninput="syncW('dist')"></div>
            <div><span class="sub">학생 희망도 (%)</span><input type="number" id="wPref" value="30" min="0" max="100" oninput="syncW('pref')"></div>
          </div></div>
        <div class="sec"><div class="grid g2">
          <div><span class="lbl">실습 시작일</span><input type="date" id="startDate" onchange="onStartChange()"></div>
          <div><span class="lbl">실습 종료일 (평일 5일 자동)</span><input type="date" id="endDate" onchange="checkDate(this)"></div>
        </div><div class="hint" id="dateErr" style="color:var(--bad);font-weight:700"></div></div>
        <div class="sec"><div class="grid g2">
          <div><span class="lbl">기간별 최대정원</span><input type="number" id="capPeriod" value="50" min="1"></div>
          <div><span class="lbl">기관별 최대정원</span><input type="number" id="capHosp" value="10" min="1"></div>
        </div><div class="hint">병원 한 곳의 실제 정원 = 두 값 중 작은 값</div></div>
        <div class="sec"><div class="grid g3">
          <div><span class="lbl">성별 조건</span><select id="gender"><option>무관</option><option>남성만</option><option>여성만</option></select></div>
          <div><span class="lbl">최소 GPA</span><input type="number" step="0.1" id="minGpa" placeholder="예: 3.5"></div>
          <div><span class="lbl">출생연도 이후</span><input type="number" id="birthYear" placeholder="예: 2003"></div>
        </div></div>
        <label class="sec" style="display:flex;gap:8px;align-items:center;font-size:14px;font-weight:600;cursor:pointer">
          <input type="checkbox" id="exPast" style="width:auto"> 과거 실습 기관 중복 배정 자동 제외</label>
        <label class="sec" style="display:flex;gap:8px;align-items:flex-start;font-size:14px;font-weight:600;cursor:pointer">
          <input type="checkbox" id="useTransit" style="width:auto;margin-top:4px">
          <span>🚌 지도 기반 통학시간 사용 (거리 반영)<br><span class="sub" id="transitHint" style="font-weight:400"></span></span></label>
      </div>
    </details>

    <div class="actions"><span></span><button class="btn btn-primary" onclick="goStep(2)">다음: 명단 업로드 →</button></div>
  </section>

  <!-- 2단계 -->
  <section id="step2" class="card step" style="display:none">
    <h2>학생 명단을 올려 주세요</h2>
    <p class="lead">파일은 메모리에서만 처리되며 서버에 저장되지 않습니다.</p>
    <div class="chips" id="condChips"></div>

    <div class="drop" id="drop">
      <div class="big">엑셀(.xlsx) 또는 CSV 파일을 여기에 끌어다 놓으세요</div>
      <div class="sub" style="margin-bottom:14px">필수: 학번 · 이름 · 성별 · GPA · 출생연도 &nbsp;/&nbsp; 선택: 주소 · 과거실습지 · 1~3지망병원</div>
      <label class="btn btn-ghost btn-sm" style="cursor:pointer">파일 선택<input type="file" id="file" accept=".csv,.xlsx" style="display:none"></label>
      <div class="file-ok" id="fileOk"></div>
    </div>
    <div class="or">또는</div>
    <div class="tools">
      <button class="btn btn-ghost" onclick="loadSample()">✨ 샘플 데이터로 체험 (가상 학생 60명)</button>
      <a class="btn btn-ghost" href="/api/v1/template">📄 양식 샘플 다운로드</a>
    </div>

    <div class="actions">
      <button class="btn btn-ghost" onclick="goStep(1)">← 이전</button>
      <button class="btn btn-primary" id="runBtn" onclick="run()">🚀 최적 배정 실행</button>
    </div>
  </section>

  <!-- 3단계 -->
  <section id="step3" class="step" style="display:none">
    <div class="res-top">
      <div class="meta" id="resMeta"></div>
      <div style="display:flex;gap:8px;flex-wrap:wrap">
        <button class="btn btn-ghost btn-sm" onclick="goStep(1)">⚙️ 조건 바꿔 다시 배정</button>
        <button class="btn btn-primary btn-sm" onclick="exportXlsx()">📥 결과 엑셀 다운로드</button>
      </div>
    </div>

    <div id="warn"></div>
    <div id="hero"></div>
    <div class="kpis" id="kpis"></div>
    <div class="charts" id="charts"></div>

    <div class="card" style="padding:0;overflow:hidden">
      <div class="tabs">
        <button class="tab on" data-p="pA" onclick="showPane('pA')">✅ 배정 결과</button>
        <button class="tab" data-p="pH" onclick="showPane('pH')">🏥 병원별 현황</button>
        <button class="tab" data-p="pU" onclick="showPane('pU')">⏳ 미배정·사유</button>
        <input type="search" class="search" id="q" placeholder="이름·학번·병원 검색" oninput="renderAssigned()">
      </div>
      <div id="pA" class="pane on tbl-wrap">
        <table class="tbl"><thead><tr>
          <th>순위</th><th class="l">병원</th><th>학번</th><th>이름</th><th>GPA</th>
          <th title="학생이 적은 1~3지망과 배정 병원의 일치 여부">지망 반영</th>
          <th title="TMAP 대중교통 소요시간. ≈ 는 지도 좌표 직선거리 추정, * 는 주소 기반 규칙 추정">통학시간</th>
          <th title="통학 점수 × 거리 가중치 + 지망 점수 × 희망도 가중치">종합점수</th>
          <th title="통학 부담(소요시간·환승·도보) 기반 규칙형 지수. 학습 모델이 아닙니다">만족도 지수</th>
        </tr></thead><tbody id="bodyA"></tbody></table>
        <div class="sub" style="padding:10px 16px">행을 누르면 점수가 어떻게 계산됐는지 볼 수 있습니다. 통학시간의 ≈ 는 지도 좌표 직선거리 기반, * 는 규칙 기반 추정치입니다.</div>
      </div>
      <div id="pH" class="pane tbl-wrap">
        <table class="tbl"><thead><tr><th class="l">병원</th><th>배정 / 정원</th><th>평균 통학</th><th>지망 일치 인원</th></tr></thead><tbody id="bodyH"></tbody></table>
      </div>
      <div id="pU" class="pane tbl-wrap">
        <table class="tbl"><thead><tr><th>구분</th><th>학번</th><th>이름</th><th>GPA</th><th class="l">사유</th></tr></thead><tbody id="bodyU"></tbody></table>
      </div>
    </div>
  </section>
</main>
<div style="height:60px"></div>

<div class="modal" id="modal" onclick="if(event.target===this)closeModal()"><div class="sheet" id="sheet" role="dialog" aria-modal="true"></div></div>
<div class="toast" id="toast" role="status"></div>

<script>
const $ = id => document.getElementById(id);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const num = (v, d = 1) => (v === null || v === undefined || isNaN(v)) ? '-' : Number(v).toFixed(d).replace(/\.0+$/, '');
function store(k, v) { try { v === null ? sessionStorage.removeItem(k) : sessionStorage.setItem(k, v); } catch (e) {} }
function load(k) { try { return sessionStorage.getItem(k) || ''; } catch (e) { return ''; } }
let TOKEN = load('token'), CURRICULUM = {}, CONFIG = {}, DATA = null, FILE = null, STEP = 1;
let toastTimer;
function toast(msg, err) {
  const t = $('toast'); t.textContent = msg; t.className = 'toast show' + (err ? ' err' : '');
  clearTimeout(toastTimer); toastTimer = setTimeout(() => t.className = 'toast', err ? 6000 : 3200);
}

/* ===== 인증 ===== */
$('auth').style.display = 'none';   // 로그인 창은 서버가 암호를 요구할 때만 init()에서 띄운다.
async function verifyPassword() {
  $('authErr').textContent = '';
  try {
    const r = await fetch('/api/v1/verify-password', {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({password: $('authPw').value})});
    if (r.ok) { TOKEN = (await r.json()).token; store('token', TOKEN); $('authPw').value = ''; $('auth').style.display = 'none'; }
    else $('authErr').textContent = '암호가 올바르지 않습니다.';
  } catch (e) { $('authErr').textContent = '서버에 연결할 수 없습니다.'; }
}
function authExpired() { store('token', null); TOKEN = ''; if (CONFIG.auth_required !== false) $('auth').style.display = 'flex'; }

/* ===== 날짜 ===== */
const HOLIDAYS = new Set([
  '2026-01-01','2026-02-16','2026-02-17','2026-02-18','2026-03-01','2026-03-02','2026-05-05','2026-05-24','2026-05-25',
  '2026-06-03','2026-06-06','2026-08-15','2026-08-17','2026-09-24','2026-09-25','2026-09-26','2026-10-03','2026-10-05','2026-10-09','2026-12-25',
  '2027-01-01','2027-02-06','2027-02-07','2027-02-08','2027-02-09','2027-03-01','2027-05-05','2027-05-13','2027-06-06',
  '2027-08-15','2027-08-16','2027-09-14','2027-09-15','2027-09-16','2027-10-03','2027-10-04','2027-10-09','2027-10-11','2027-12-25','2027-12-27']);
const FIXED = new Set(['01-01','03-01','05-05','06-06','08-15','10-03','10-09','12-25']);
const pd_ = s => { const [y, m, d] = s.split('-').map(Number); return new Date(y, m - 1, d); };
const fmt = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
function dateProblem(s) {
  const w = pd_(s).getDay();
  if (w === 0 || w === 6) return '주말(토·일요일)은 실습일로 지정할 수 없습니다.';
  if (HOLIDAYS.has(s) || FIXED.has(s.slice(5))) return '공휴일·명절 연휴는 실습일로 지정할 수 없습니다.';
  return null;
}
function addBiz(s, n) { const d = pd_(s); let c = 0; while (c < n) { d.setDate(d.getDate() + 1); if (!dateProblem(fmt(d))) c++; } return fmt(d); }
function initDates() {
  const d = new Date(); d.setDate(d.getDate() + ((8 - d.getDay()) % 7 || 7));
  let s = fmt(d); while (dateProblem(s)) { d.setDate(d.getDate() + 1); s = fmt(d); }
  $('startDate').value = s; $('endDate').value = addBiz(s, 4);
}
function checkDate(inp) { const m = inp.value ? dateProblem(inp.value) : null; $('dateErr').textContent = m ? '⚠️ ' + m : ''; if (m) inp.value = ''; return !m; }
function onStartChange() { const s = $('startDate'); if (checkDate(s) && s.value) $('endDate').value = addBiz(s.value, 4); }
function syncW(ch) {
  const a = $('wDist'), b = $('wPref'), cl = v => Math.max(0, Math.min(100, parseInt(v) || 0));
  if (ch === 'dist') { a.value = cl(a.value); b.value = 100 - a.value; } else { b.value = cl(b.value); a.value = 100 - b.value; }
}

/* ===== 교과목·병원 ===== */
const mode = () => document.querySelector('input[name=mode]:checked').value;
function onModeChange() {
  $('modeG').classList.toggle('on', mode() === 'global'); $('modeS').classList.toggle('on', mode() === 'single');
  $('hospWrap').style.display = mode() === 'single' ? 'block' : 'none';
}
function updateSubjects() {
  const subs = (CURRICULUM[$('grade').value] || {})[$('semester').value] || [];
  const box = $('subject'); box.innerHTML = '';
  subs.forEach(x => box.add(new Option(x + ' 실습', x)));
  box.add(new Option('전체 교과목 병원 통합', 'ALL'));
  updateHospitals();
}
async function updateHospitals() {
  const q = new URLSearchParams({grade: $('grade').value, semester: $('semester').value, subject: $('subject').value});
  const box = $('hospital'); box.innerHTML = '';
  try {
    const d = await (await fetch('/api/v1/hospitals?' + q)).json();
    (d.hospitals || []).forEach(h => box.add(new Option(h.includes('방중') ? h + ' 🏖️' : h, h)));
    $('hospCount').textContent = `이 교과목의 실습 병원 ${(d.hospitals || []).length}곳`;
  } catch (e) { $('hospCount').textContent = '병원 목록을 불러오지 못했습니다.'; }
}
window.addEventListener('load', async () => {
  initDates();
  try { CURRICULUM = await (await fetch('/api/v1/curriculum')).json(); } catch (e) {}
  try { CONFIG = await (await fetch('/api/v1/config')).json(); } catch (e) { CONFIG = {}; }
  if (CONFIG.auth_required && !TOKEN) $('auth').style.display = 'flex';
  const ut = $('useTransit');
  if (CONFIG.transit_available) {
    const m = CONFIG.hospitals_missing_address;
    ut.checked = true;
    $('transitHint').textContent = (CONFIG.transit_mode === 'tmap' ? 'TMAP 대중교통 소요시간을 사용합니다. ' : '지도 좌표의 직선거리로 통학시간을 추정합니다. ')
      + (m ? `병원 주소 미등록 ${m}곳은 규칙 기반 추정치로 대체됩니다. ` : '') + '첫 실행은 시간이 걸릴 수 있고 이후엔 저장된 값을 사용합니다. (동 단위 주소만 전송)';
  } else {
    ut.disabled = true; $('transitHint').textContent = 'API 키가 설정되지 않아 규칙 기반 추정치를 사용합니다.';
  }
  updateSubjects(); onModeChange();
});

/* ===== 단계 이동 ===== */
function goStep(n) {
  if (n === 3 && !DATA) return;
  STEP = n; $('toast').className = 'toast';
  document.querySelectorAll('.step').forEach(s => s.style.display = s.id === 'step' + n ? 'block' : 'none');
  document.querySelectorAll('#stepper li').forEach(li => {
    const s = +li.dataset.s;
    li.className = s === n ? 'active' : ((s < n || (s === 3 && DATA) || (s === 2)) ? 'done' : '');
  });
  if (n === 2) renderChips();
  window.scrollTo({top: 0, behavior: 'smooth'});
}
document.querySelectorAll('#stepper li').forEach(li => li.addEventListener('click', () => { if (li.classList.contains('done')) goStep(+li.dataset.s); }));
function renderChips() {
  const chips = [mode() === 'global' ? '전체 병원 동시 배정' : '단일 병원: ' + ($('hospital').value || '미선택'),
    `${$('grade').value} ${$('semester').value}`, $('subject').selectedOptions[0]?.textContent || '',
    `거리 ${$('wDist').value} : 희망 ${$('wPref').value}`, `병원당 정원 ${Math.min(+$('capHosp').value || 10, +$('capPeriod').value || 50)}명`,
    `${$('startDate').value} ~ ${$('endDate').value}`];
  $('condChips').innerHTML = chips.filter(Boolean).map(c => `<span class="chip">${esc(c)}</span>`).join('');
}

/* ===== 파일 ===== */
function setFile(f) {
  FILE = f; const ok = $('fileOk');
  ok.style.display = 'block'; ok.textContent = `✅ ${f.name} (${Math.max(1, Math.round(f.size / 1024))}KB)`;
}
$('file').addEventListener('change', e => { if (e.target.files[0]) setFile(e.target.files[0]); });
const dz = $('drop');
['dragenter', 'dragover'].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.add('over'); }));
['dragleave', 'drop'].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.remove('over'); }));
dz.addEventListener('drop', e => { const f = e.dataTransfer.files[0]; if (f) setFile(f); });
async function loadSample() {
  try {
    const q = new URLSearchParams({grade: $('grade').value, semester: $('semester').value, subject: $('subject').value});
    const r = await fetch('/api/v1/sample-data?' + q);
    if (!r.ok) throw 0;
    setFile(new File([await r.blob()], '샘플_가상학생_60명.xlsx'));
    toast('가상 학생 60명을 불러왔습니다. 이제 실행하세요.');
  } catch (e) { toast('샘플 데이터를 불러오지 못했습니다.', true); }
}

/* ===== 실행 ===== */
async function run() {
  if (!FILE) { toast('학생 명단 파일을 선택하거나 샘플 데이터를 불러오세요.', true); return; }
  if (mode() === 'single' && !$('hospital').value) { toast('배정 대상 병원을 선택해 주세요.', true); goStep(1); return; }
  if (!$('startDate').value || !$('endDate').value) { toast('실습 시작일과 종료일을 확인해 주세요.', true); goStep(1); return; }
  const f = new FormData();
  f.append('mode', mode()); f.append('target_hospital', $('hospital').value || '');
  f.append('grade', $('grade').value); f.append('semester', $('semester').value); f.append('subject', $('subject').value);
  f.append('start_date', $('startDate').value); f.append('end_date', $('endDate').value);
  f.append('max_period_capacity', $('capPeriod').value || 50); f.append('max_hospital_capacity', $('capHosp').value || 10);
  f.append('gender_criteria', $('gender').value); f.append('exclude_past_hospital', $('exPast').checked);
  f.append('use_transit_api', $('useTransit').checked && !$('useTransit').disabled);
  f.append('distance_weight', (parseInt($('wDist').value) || 0) / 100); f.append('preference_weight', (parseInt($('wPref').value) || 0) / 100);
  if ($('minGpa').value !== '') f.append('min_gpa', $('minGpa').value);
  if ($('birthYear').value !== '') f.append('birth_year_after', $('birthYear').value);
  f.append('file', FILE);
  const b = $('runBtn'); b.disabled = true; b.textContent = '⏳ 최적 배정 계산 중...';
  try {
    const r = await fetch('/api/v1/assign-file', {method: 'POST', body: f, headers: {'X-Auth-Token': TOKEN}});
    let d = {}; try { d = await r.json(); } catch (e) {}
    if (r.status === 401) { authExpired(); return; }
    if (!r.ok) { toast(d.detail || '요청 처리 중 오류가 발생했습니다.', true); return; }
    DATA = d; DATA._ctx = `${$('grade').value} ${$('semester').value}`; render(d); goStep(3);
  } catch (e) { toast('서버에 연결할 수 없습니다.', true); }
  finally { b.disabled = false; b.textContent = '🚀 최적 배정 실행'; }
}

/* ===== 결과 렌더 ===== */
function showPane(id) {
  document.querySelectorAll('.pane').forEach(p => p.classList.toggle('on', p.id === id));
  document.querySelectorAll('.tab').forEach(t => t.classList.toggle('on', t.dataset.p === id));
  $('q').style.display = id === 'pA' ? '' : 'none';
}
function pbar(tag, cls, val, max, label) {
  const w = max > 0 ? Math.max(2, Math.min(100, val / max * 100)) : 0;
  return `<div class="pbar"><span class="tag">${tag}</span><div class="track"><div class="fill ${cls}" style="width:${w}%"></div></div><span class="val">${label}</span></div>`;
}
function renderHero(d) {
  const c = d.comparison;
  if (!c) { $('hero').innerHTML = ''; return; }
  const sv = c.travel_saved_minutes, pr = c.pref_rate_delta;
  const sgn = v => v > 0 ? '+' : (v < 0 ? '−' : '');
  const cls = v => v > 0 ? 'pos' : (v < 0 ? 'neg' : '');
  const tmax = Math.max(c.baseline.avg_travel_minutes, c.ours.avg_travel_minutes);
  $('hero').innerHTML = `<div class="hero">
    <div><div class="eyebrow">기존 방식 대비 개선 효과</div>
      <div class="bignums">
        <div class="bignum"><div class="n ${cls(pr)}">${sgn(pr)}${num(Math.abs(pr))}<small>%p</small></div><div class="t">지망 충족률</div></div>
        <div class="bignum"><div class="n ${cls(c.score_delta)}">${sgn(c.score_delta)}${num(Math.abs(c.score_delta))}<small>점</small></div><div class="t">평균 종합 매칭점수</div></div>
        <div class="bignum"><div class="n ${cls(sv)}" style="font-size:26px;margin-top:14px">${sv > 0 ? '−' : sgn(-sv)}${num(Math.abs(sv))}<small>분</small></div><div class="t">평균 통학시간 ${sv > 0 ? '단축' : (sv < 0 ? '증가' : '변화 없음')}</div></div>
      </div></div>
    <div>
      <div class="cmp-row"><div class="cl">평균 통학시간 (분, 낮을수록 좋음)</div>
        ${pbar('기존', 'base', c.baseline.avg_travel_minutes, tmax, num(c.baseline.avg_travel_minutes) + '분')}
        ${pbar('개선', 'ours', c.ours.avg_travel_minutes, tmax, num(c.ours.avg_travel_minutes) + '분')}</div>
      <div class="cmp-row"><div class="cl">지망 충족률 (1~3지망에 배정된 비율)</div>
        ${pbar('기존', 'base', c.baseline.pref_match_rate, 100, num(c.baseline.pref_match_rate) + '%')}
        ${pbar('개선', 'ours', c.ours.pref_match_rate, 100, num(c.ours.pref_match_rate) + '%')}</div>
      <div class="cmp-row"><div class="cl">평균 종합 매칭점수</div>
        ${pbar('기존', 'base', c.baseline.avg_composite_score, 100, num(c.baseline.avg_composite_score) + '점')}
        ${pbar('개선', 'ours', c.ours.avg_composite_score, 100, num(c.ours.avg_composite_score) + '점')}</div>
      <div class="cmp-note">비교 기준: ${esc(c.baseline_label)}<br>${d.transit && d.transit.real_ratio > 0 ? `통학시간: ${d.transit.mode === 'tmap' ? 'TMAP 대중교통 실측' : '지도 좌표 직선거리 추정'} ${num(d.transit.real_ratio)}% (나머지는 규칙 기반 추정)` : '통학시간은 동 단위 규칙 기반 추정치입니다.'}</div>
    </div></div>`;
}
function renderKpis(d) {
  const k = (t, v, s) => `<div class="kpi"><div class="k">${esc(t)}</div><div class="v">${esc(v)}<small>${esc(s || '')}</small></div></div>`;
  $('kpis').innerHTML = k('배정 / 전체 학생', d.assigned_count, `/ ${d.total_students}명`) + k('평균 통학시간(추정)', num(d.avg_travel_minutes), '분') +
    k('지망 충족률', num(d.pref_match_rate), '%') + k('대기 · 자격미달', `${d.waitlist_count} · ${d.ineligible_count}`, '명');
}
function renderCharts(d) {
  const rs = d.results; let html = '';
  if (d.mode === 'global' && d.hospital_summary.length) {
    html += `<div class="chart"><h3>병원별 배정 현황</h3>` + d.hospital_summary.map(h => {
      const pct = h.capacity ? h.assigned / h.capacity * 100 : 0;
      return `<div class="hrow"><span class="nm" title="${esc(h.hospital)}">${esc(h.hospital)}</span><div class="tr"><div class="fl ${pct >= 100 ? 'full' : ''}" style="width:${pct}%"></div></div><span class="ct">${h.assigned}/${h.capacity}</span></div>`;
    }).join('') + `<div class="sub" style="margin-top:8px">진한 파랑 = 정원 마감</div></div>`;
  }
  const bins = [['~24분', 0, 24], ['25–34', 25, 34], ['35–44', 35, 44], ['45–54', 45, 54], ['55–64', 55, 64], ['65분~', 65, 999]];
  const cnt = bins.map(([, lo, hi]) => rs.filter(r => r.travel_time_minutes >= lo && r.travel_time_minutes <= hi).length);
  const mx = Math.max(1, ...cnt);
  html += `<div class="chart"><h3>통학시간 분포 (추정)</h3><div class="vbars">` +
    cnt.map(c => `<div class="vb"><span class="n">${c}</span><div class="bar" style="height:${c / mx * 100}%"></div></div>`).join('') +
    `</div><div class="vlabels">${bins.map(b => `<span>${b[0]}</span>`).join('')}</div></div>`;
  const pc = [['1지망', '#0f766e'], ['2지망', '#2dd4bf'], ['3지망', '#99f6e4'], ['비지망', '#cbd5e1']];
  const pv = pc.map(([l]) => rs.filter(r => l === '비지망' ? r.preference_status.startsWith('비') : r.preference_status.startsWith(l)).length);
  const tot = pv.reduce((a, b) => a + b, 0); let acc = 0;
  const stops = tot ? pc.map(([, col], i) => { const a = acc / tot * 100; acc += pv[i]; return `${col} ${a}% ${acc / tot * 100}%`; }).join(',') : '#e2e8f0 0 100%';
  html += `<div class="chart"><h3>지망 충족 현황</h3><div class="donut-wrap"><div class="donut" style="background:conic-gradient(${stops})"><div class="donut-hole"><b>${num(d.pref_match_rate)}%</b><span>지망 충족</span></div></div>
    <div class="legend">${pc.map(([l, col], i) => `<div><i style="background:${col}"></i>${l} <b>${pv[i]}명</b></div>`).join('')}</div></div></div>`;
  $('charts').innerHTML = html;
}
function renderAssigned() {
  const q = $('q').value.trim().toLowerCase();
  const rows = DATA.results.map((r, i) => [r, i]).filter(([r]) => !q || [r.name, r.student_id, r.hospital].some(v => String(v).toLowerCase().includes(q)));
  $('bodyA').innerHTML = rows.map(([r, i]) => `<tr class="click" onclick="openDetail(${i})">
    <td><span class="rk">${esc(r.rank)}위</span></td><td class="l">${esc(r.hospital)}</td><td>${esc(r.student_id)}</td><td><b>${esc(r.name)}</b></td><td>${esc(r.gpa)}</td>
    <td><span class="pill ${r.preference_status.includes('일치') ? 'hit' : 'miss'}">${esc(r.preference_status)}</span></td>
    <td><b>${esc(r.travel_time_minutes)}분</b>${r.transit_source === 'TMAP' ? '' : `<span class="sub">${r.transit_source === '직선거리' ? '≈' : '*'}</span>`}</td><td><span class="score">${esc(r.composite_score)}</span></td><td>${esc(r.satisfaction_index)}</td></tr>`).join('')
    || `<tr><td colspan="9" class="empty">${q ? '검색 결과가 없습니다.' : '배정된 학생이 없습니다.'}</td></tr>`;
}
function render(d) {
  $('resMeta').innerHTML = `<b>[${esc(d._ctx)}]</b> ${esc(d.period.start)} ~ ${esc(d.period.end)} · 대상 <b>${esc(d.target_hospital)}</b><br><span class="sub">${esc(d.optimization_method)}</span>`;
  $('warn').innerHTML = d.warnings.length ? `<div class="alert"><b>⚠️ 데이터 확인 필요 (${d.warnings.length}건)</b><ul>${d.warnings.map(w => `<li>${esc(w)}</li>`).join('')}</ul></div>` : '';
  renderHero(d); renderKpis(d); renderCharts(d); renderAssigned();
  $('bodyH').innerHTML = d.hospital_summary.map(h => `<tr><td class="l">${esc(h.hospital)}</td><td><b>${h.assigned}</b> / ${h.capacity}</td><td>${h.avg_travel === null ? '-' : esc(h.avg_travel) + '분'}</td><td>${h.pref_matched}명</td></tr>`).join('');
  $('bodyU').innerHTML = d.unassigned.map(u => `<tr><td><span class="pill ${u.status === '대기' ? 'wait' : 'bad'}">${esc(u.status)}</span></td><td>${esc(u.student_id)}</td><td><b>${esc(u.name)}</b></td><td>${esc(u.gpa)}</td><td class="l">${esc(u.reason)}</td></tr>`).join('')
    || '<tr><td colspan="5" class="empty">미배정 학생이 없습니다. 🎉</td></tr>';
  $('q').value = ''; showPane('pA');
}

/* ===== 학생 상세 ===== */
function openDetail(i) {
  const r = DATA.results[i], w = DATA.weights;
  const dc = r.dist_score * w.distance / 100, pc = r.pref_score * w.preference / 100;
  const alts = (r.alternatives || []).map(a => `<div class="alt"><span>${esc(a.hospital)} <span class="sub">· ${esc(a.preference_status)} · 통학 ${esc(a.travel_time_minutes)}분</span></span><b>${esc(a.composite_score)}점</b></div>`).join('');
  const note = r.is_top_choice === false
    ? `이 학생의 최고점 병원은 <b>${esc(r.best_hospital)}</b>이지만, 병원 정원과 전체 최적화(모든 학생의 점수 합 최대화) 결과 <b>${esc(r.hospital)}</b>에 배정되었습니다.`
    : (r.is_top_choice === true ? '이 학생에게 가장 점수가 높은 병원에 배정되었습니다. 👍' : '');
  $('sheet').innerHTML = `<button class="x" onclick="closeModal()" aria-label="닫기">×</button>
    <h3>${esc(r.name)} <span class="sub" style="font-weight:500">${esc(r.student_id)}</span></h3>
    <div class="sub">${esc(r.gender)} · GPA ${esc(r.gpa)} · ${esc(r.birth_year)}년생 · ${esc(r.address || '주소 없음')}</div>
    <div class="box" style="background:#f0fdfa"><b style="color:var(--teal-d)">${esc(r.hospital)}</b> · 병원 내 ${esc(r.rank)}순위 · 종합 <b>${esc(r.composite_score)}점</b></div>
    <div class="box"><h4>종합 매칭점수는 이렇게 계산됐어요</h4>
      <div class="stack"><i style="width:${dc}%;background:var(--teal)"></i><i style="width:${pc}%;background:var(--amber)"></i></div>
      <div class="eq"><span>🟩 통학 점수 ${num(r.dist_score)} × ${w.distance}%</span><b>${num(dc)}</b>
        <span>🟨 지망 점수 ${num(r.pref_score)} (${esc(r.preference_status)}) × ${w.preference}%</span><b>${num(pc)}</b>
        <span><b style="text-align:left">종합 매칭점수</b></span><b>${esc(r.composite_score)}</b></div></div>
    <div class="box"><h4>통학 (${r.transit_source === 'TMAP' ? 'TMAP 대중교통 · 평일 오전 8시 기준' : (r.transit_source === '직선거리' ? '지도 좌표 직선거리 기반 추정' : '규칙 기반 추정')})</h4><div class="eq"><span>소요시간</span><b>${esc(r.travel_time_minutes)}분</b><span>환승</span><b>${esc(r.transfers)}회</b><span>피로지수(MFI)</span><b>${esc(r.fatigue_index)}</b><span>예상 만족도 지수</span><b>${esc(r.satisfaction_index)}</b></div></div>
    ${alts ? `<div class="box"><h4>다른 병원이었다면</h4>${alts}</div>` : ''}
    ${note ? `<div class="hint" style="margin-top:12px">${note}</div>` : ''}`;
  $('modal').classList.add('open');
}
function closeModal() { $('modal').classList.remove('open'); }
document.addEventListener('keydown', e => { if (e.key === 'Escape') closeModal(); });

/* ===== 엑셀 내보내기 (서버에서 생성 → 외부 라이브러리 불필요) ===== */
async function exportXlsx() {
  if (!DATA) return;
  try {
    const r = await fetch('/api/v1/export-xlsx', {method: 'POST', headers: {'Content-Type': 'application/json', 'X-Auth-Token': TOKEN}, body: JSON.stringify(DATA)});
    if (r.status === 401) { authExpired(); return; }
    if (!r.ok) throw 0;
    const a = document.createElement('a'); a.href = URL.createObjectURL(await r.blob());
    a.download = `실습지배정결과_${DATA.target_hospital}.xlsx`; document.body.appendChild(a); a.click(); a.remove();
  } catch (e) { toast('엑셀 내보내기에 실패했습니다.', true); }
}
</script>
</body>
</html>
"""


@app.get("/", response_class=HTMLResponse)
def render_ui():
    return HTMLResponse(content=HTML_PAGE.replace("__APP_VERSION__", APP_VERSION))


def run_check() -> None:
    """python main.py check — 키·병원 주소·지도 조회가 되는지 한국어로 점검."""
    print("=== 지도 연동 점검 ===")
    print("네이버 키(ID/Secret):", "설정됨" if TRANSIT_AVAILABLE else "없음 → keys.env 파일을 확인하세요")
    print("TMAP 키:", "설정됨" if TMAP_APP_KEY else "없음 (없어도 됩니다: 좌표 직선거리로 통학시간을 추정)")
    if not TRANSIT_AVAILABLE:
        print("→ 지금은 규칙 기반 추정치로 동작합니다.")
        return
    print("통학시간 방식:", "TMAP 대중교통" if TMAP_APP_KEY else "지도 좌표 직선거리 추정")
    addr = load_hospital_addresses()
    print(f"병원 주소 입력: {len(addr)}곳 (hospital_addresses.json)")
    tests = ["경기도 안산시", "경기도 안산시 단원구", "서울시 구로구", "경기도 안산시 단원구 고잔동"]
    tests += list(addr.values())[:5]
    for q in tests:
        try:
            a = _geocode_raw(q)
            if a:
                print(f"  OK   {q} → {a[0].get('roadAddress') or a[0].get('jibunAddress')} ({a[0]['x']}, {a[0]['y']})")
            else:
                print(f"  없음 {q} → 이 주소는 좌표로 바꿀 수 없습니다(더 자세한 주소가 필요)")
        except urllib.error.HTTPError as e:
            hint = " → 키가 틀렸거나 Maps(Geocoding)가 선택되지 않았습니다" if e.code in (401, 403) else ""
            print(f"  오류 {q} → HTTP {e.code}{hint}")
        except Exception as e:
            print(f"  오류 {q} → {type(e).__name__}: 인터넷 연결을 확인하세요")


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "check":
        run_check()
    else:
        uvicorn.run(app, host="0.0.0.0", port=8000)
