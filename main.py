"""
AI기반 간호실습 배치 의사결정 프로그램

버전 기록 (수정할 때마다 APP_VERSION 을 올립니다)
  v13.3 소개 페이지 디자인을 v12.2로 되돌리고 "만든 사람들"(팀 캐릭터)과 맨 위 버전 표시만 유지
  v13.2 대중교통 조회 기준을 평일 오전 6시 출발로 변경(07시 실습 시작 기준, 환경변수 COMMUTE_TIME으로 조절)
  v13.1 소개 페이지 맨 위에 버전 표시
  v13.0 소개 페이지 새 디자인(팀 캐릭터·실제 화면·개발 이유, 포스터와 같은 파란색)
  v12.2 소개 페이지 정리(업데이트 내역 제거, 예시 표 단순화, 버튼 문구 통일, 개선 효과·팀 정보 표시)
  v12.1 화면 문구에서 "직선거리" 표현 정리(통학은 "지도 기반 예상 소요시간"으로 표시)
  v12.0 프로그램 이름을 "AI기반 간호실습 배치 의사결정 프로그램"으로 변경(팀 이름 로켓단은 유지)
  v11.9 소개 페이지 구역마다 화면 높이에 맞게 꽉 채우고 가운데 정렬(영상 촬영용)
  v11.8 소개 페이지 숫자 키(1~9)로 구역 이동
  v11.7 소개 페이지 오른쪽 점 모양 구역 안내(PC 화면)
  v11.6 소개 페이지 순서 변경(소개 → 서비스 시작하기 → 사용법 → 순서), 자주 묻는 질문 추가
  v11.5 소개 페이지: 화면 예시, 거리·희망 비중 체험, 수작업과 비교, 숫자 올라가는 효과
  v11.4 소식(업데이트) 맨 위에 현재 버전 표시
  v11.3 소개 페이지 첫 문장을 짧게(한 줄로 이어지도록)
  v11.2 소개 페이지 줄바꿈 정리(단어 중간에서 끊기지 않고 한 줄에 가깝게)
  v11.1 소개 페이지를 홈페이지 형태로 개편(상단 메뉴·Why·숫자·서비스·소식), 사용설명서 링크(GUIDE_URL 환경변수)
  v11.0 소개 페이지(/) 추가, 배정 도구는 /app 으로 이동, 검색용 정보(robots.txt·sitemap.xml)
  v10.9 지도 문구 정리(거주지 표시, 동 기준 위치 안내)
  v10.8 화면에서 이름(가운데 글자)·학번(가운데 3자리) 가림 처리(엑셀 다운로드는 원본 유지)
  v10.7 지도에 거주지-병원 직선과 직선거리 다시 표시(v10.5와 같은 화면)
  v10.6 지도에서 직선·거리 표시 제거(거주지와 병원 두 위치만 표시)
  v10.5 거주지-병원 지도를 네이버 지도로 변경(Dynamic Map 미설정 시 기존 지도로 대체)
  v10.4 자문 반영: 문구·이모지 정리, 엑셀 버튼 위치, 20명 단위 페이지, 열 정렬, 거주지-병원 지도(상세·다른 병원)
  v10.3 실습일수 표시(8~10일 규칙): 8일 미만이면 경고, 오늘이 목·금이면 시작일을 다음 월요일로
  v10.2 종료일 기본값: 시작일이 속한 주의 다음 주 금요일(휴일이면 그 전 평일)
  v10.1 실습 기간 기본값: 시작일=오늘(휴일이면 다음 평일)
  v10.0 화면 정리(상세 설정 재구성, 지도 거리 항상 사용, 로고 아이콘·샘플 버튼 정리, 통학시간 "약" 표기)
  v9.9  병원별 정원을 화면에서 직접 입력(기본값: 2025-1 실제 배정 인원)
  v9.8  병원별 실제 정원(2025-1 실제 배정 인원) 사용 옵션
  v9.7  지도가 못 찾은 주소를 결과 경고에 표시
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
from fastapi import Body, FastAPI, File, Form, Header, HTTPException, Request, UploadFile
from fastapi.concurrency import run_in_threadpool
from fastapi.responses import HTMLResponse, PlainTextResponse, Response, StreamingResponse
from pydantic import BaseModel
from scipy.optimize import linear_sum_assignment

APP_VERSION = "v13.3"
app = FastAPI(title=f"AI기반 간호실습 배치 의사결정 프로그램 ({APP_VERSION})")

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
# 학과의 실제 배정 인원(2025-1 4학년 관리 실습). 병원마다 받는 인원이 다르므로 '병원별 실제 정원'으로 쓸 수 있다.
ACTUAL_CAPS = {
    ("4학년", "1학기", "관리"): {
        "순천향대학교부천병원": 36, "순천향대학교서울병원": 36, "단원병원": 28, "아주대학교병원": 24,
        "고려대학교구로병원": 24, "시화병원": 23, "한림대학교동탄성심병원": 23, "서울특별시서남병원": 19,
        "중앙대학교광명병원": 8,
    },
}


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
# 대중교통 조회 출발 시각(HHMM). 실습 07:00 시작 → 06:00 출발 기준. Render 환경변수 COMMUTE_TIME 으로 변경 가능
COMMUTE_TIME = (os.environ.get("COMMUTE_TIME", "0600").strip().replace(":", "") or "0600").zfill(4)[:4]
if not (COMMUTE_TIME.isdigit() and int(COMMUTE_TIME[:2]) < 24 and int(COMMUTE_TIME[2:]) < 60):
    COMMUTE_TIME = "0600"
COMMUTE_LABEL = f"평일 오전 {int(COMMUTE_TIME[:2])}시" + (f" {int(COMMUTE_TIME[2:])}분" if COMMUTE_TIME[2:] != "00" else "") + " 출발"
LINE_BASE_MIN, LINE_MIN_PER_KM = 12, 2.0   # 직선거리 추정식: 통학(분) = 12 + 2.0 × 직선거리(km)  (실측과 비교해 조정 가능)


def _rkey(dong: str, hosp_addr: str) -> str:
    tag = f"{TRANSIT_MODE}@{COMMUTE_TIME}" if TRANSIT_MODE == "tmap" else TRANSIT_MODE   # 출발 시각이 바뀌면 다시 조회
    return f"{tag}|{dong}||{hosp_addr}"


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
    """출근 시간대 기준(다음 평일 COMMUTE_TIME, 한국시간). TMAP searchDttm 형식 yyyymmddhhmi."""
    kst = datetime.timezone(datetime.timedelta(hours=9))
    d = datetime.datetime.now(kst).date() + datetime.timedelta(days=1)
    while d.weekday() >= 5:
        d += datetime.timedelta(days=1)
    return d.strftime("%Y%m%d") + COMMUTE_TIME


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


def cached_ll(query: str) -> Optional[List[float]]:
    """이미 조회해 둔 주소의 [경도, 위도]. 새로 조회하지 않는다(없으면 None)."""
    if not query:
        return None
    v = _load_cache()["geo"].get(query)
    return [v[0], v[1]] if v else None


def hosp_ll(hospital: str) -> Optional[List[float]]:
    return cached_ll(_ADDR_MAP.get(norm_h(hospital), ""))


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
        with _cache_lock:
            bad = sorted({q for d, a in todo for q in (d, a) if q in c["geo"] and c["geo"][q] is None})
        if bad:
            notes.append(f"지도가 위치를 찾지 못한 주소 {len(bad)}개: " + " / ".join(bad[:8]) + (" 외" if len(bad) > 8 else ""))
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
        "h_ll": hosp_ll(h),
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
                    valid: List[str], cap) -> Optional[dict]:
    """기준선: 정원은 지키되 학생을 무작위로 병원에 넣는 방식의 평균(고정 시드, 반복 시뮬레이션)."""
    n = len(eligible)
    if not n:
        return None
    rng = np.random.default_rng(20260101)
    trials = 100 if n <= 500 else 20
    picked: List[dict] = []
    for _ in range(trials):
        remaining = dict(cap) if isinstance(cap, dict) else {h: cap for h in valid}
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
        "s_ll": cached_ll(stu.get("dong", "")),
        "h_ll": hosp_ll(hospital),
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
    use_actual_caps: bool = False,
    hospital_caps: Optional[Dict[str, int]] = None,
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
        cap_by_h = {h: cap_eff for h in valid}
        actual_used = False
        custom_used = False
        if hospital_caps:
            hn = {}
            for k, v in hospital_caps.items():
                try:
                    if int(v) >= 1:
                        hn[norm_h(k)] = int(v)
                except (TypeError, ValueError):
                    pass
            if hn:
                cap_by_h = {h: hn.get(norm_h(h), cap_eff) for h in valid}
                custom_used = True
        elif use_actual_caps:
            table = ACTUAL_CAPS.get((grade, semester, subject))
            if table:
                tn = {norm_h(k): v for k, v in table.items()}
                cap_by_h = {h: tn.get(norm_h(h), cap_eff) for h in valid}
                actual_used = True
            else:
                warnings.append("이 교과목은 실제 정원 자료가 없어 '기관별 최대정원' 값을 모든 병원에 똑같이 적용했습니다.")
        slot_hosp = [h for h in valid for _ in range(cap_by_h[h])]
        slots_total = len(slot_hosp)
        hospital_cap_map = dict(cap_by_h)
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
            slot_idx = np.repeat(np.arange(len(valid)), [cap_by_h[h] for h in valid])
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
                random_baseline(eligible, pairs, allowed, valid, cap_by_h),
                "무작위 배정 — 정원은 지키되 학생을 임의로 배치 (100회 시뮬레이션 평균)")
        method = (f"전체 병원 동시 배정 · SciPy 헝가리안 전역 최적화 "
                  f"(거리 {round(dw*100)}% · 희망도 {round(pw*100)}%, "
                  + ("병원별 지정 정원" if custom_used else "병원별 실제 정원(2025-1 배정 인원)" if actual_used else f"병원당 정원 {cap_eff}명")
                  + f", 병원 {len(valid)}곳)")
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
    return {"transit_available": TRANSIT_AVAILABLE, "transit_mode": TRANSIT_MODE, "hospitals_missing_address": missing, "auth_required": AUTH_REQUIRED, "version": APP_VERSION, "map_key_id": NAVER_KEY_ID or "", "commute_time": COMMUTE_TIME, "commute_label": COMMUTE_LABEL}


@app.get("/api/v1/curriculum")
async def get_curriculum():
    return {g: {s: list(subs.keys()) for s, subs in sems.items()} for g, sems in MASTER_HOSPITAL_DB.items()}


def _parse_caps(raw: str) -> Optional[Dict[str, int]]:
    try:
        d = json.loads(raw) if raw else None
        return {str(k): int(v) for k, v in d.items()} if isinstance(d, dict) else None
    except Exception:
        return None


@app.get("/api/v1/hospitals")
async def get_hospitals(grade: str, semester: str, subject: str):
    hs = hospitals_for(grade, semester, subject)
    table = ACTUAL_CAPS.get((grade, semester, subject)) or {}
    tn = {norm_h(k): v for k, v in table.items()}
    return {"hospitals": hs, "actual_caps": {h: tn[norm_h(h)] for h in hs if norm_h(h) in tn}}


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
    use_actual_caps: bool = Form(False),
    hospital_caps: str = Form(""),
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
            use_transit_api=use_transit_api, use_actual_caps=use_actual_caps,
            hospital_caps=_parse_caps(hospital_caps),
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
<title>AI기반 간호실습 배치 의사결정 프로그램 | 로켓단</title>
<meta name="robots" content="noindex">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css" media="print" onload="this.media='all'">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.css">
<script src="https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.js"></script>
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
  .adv-body h4 { margin:22px 0 8px; font-size:14px; color:var(--navy); }
  .adv-body h4:first-child { margin-top:14px; }
  details.more { margin-top:22px; border:1px dashed var(--line); border-radius:10px; }
  details.more > summary { cursor:pointer; padding:11px 14px; font-weight:600; font-size:14px; color:var(--muted); list-style:none; }
  details.more > summary::after { content:'＋'; float:right; }
  details.more[open] > summary::after { content:'－'; }
  details.more .more-body { padding:2px 14px 16px; }
  .wbar { display:flex; justify-content:space-between; font-weight:700; font-size:14px; margin-bottom:6px; }
  input[type=range] { width:100%; accent-color:var(--navy); padding:0; border:0; background:transparent; }
  .caps { display:grid; grid-template-columns:repeat(auto-fill,minmax(250px,1fr)); gap:8px 18px; }
  .caprow { display:flex; align-items:center; justify-content:space-between; gap:10px; font-size:14px; }
  .caprow input { width:84px; padding:8px 10px; text-align:right; }

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
  .tabs .xl { margin-left:auto; }
  .tabs .search { width:210px; padding:8px 12px; margin-left:8px; }
  th.sortable { cursor:pointer; user-select:none; }
  th.sortable:hover { color:var(--navy); }
  .ar { display:inline-block; width:12px; color:var(--blue); font-size:10px; }
  .pager { display:flex; align-items:center; justify-content:space-between; gap:10px; flex-wrap:wrap; padding:12px 16px; border-top:1px solid var(--line); font-size:13px; color:var(--muted); }
  .pager .pg { display:flex; gap:4px; flex-wrap:wrap; }
  .pager button { min-width:34px; height:34px; border:1px solid var(--line); background:#fff; border-radius:8px; cursor:pointer; font-size:13px; color:#334155; padding:0 8px; }
  .pager button.on { background:var(--blue); border-color:var(--blue); color:#fff; font-weight:700; }
  .pager button:disabled { opacity:.4; cursor:default; }
  .map { height:220px; border-radius:10px; overflow:hidden; border:1px solid var(--line); margin-top:4px; }
  .mapnone { font-size:13px; color:var(--muted); padding:6px 0; }
  .alt.click { cursor:pointer; } .alt.click:hover { color:var(--blue); }
  .altmap { display:none; margin:4px 0 10px; } .altmap.open { display:block; }
  .alt .lnk { color:var(--blue); font-size:12px; margin-left:6px; }
  .dot { display:inline-block; width:10px; height:10px; border-radius:2px; margin-right:6px; vertical-align:baseline; }
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
    <div style="opacity:.75;font-size:13px">AI기반 간호실습 배치 의사결정 프로그램</div>
    <input type="password" id="authPw" placeholder="접속 암호 입력" autocomplete="current-password" onkeyup="if(event.key==='Enter')verifyPassword()">
    <div id="authErr" style="color:#fbbf24;font-size:13px;min-height:20px;margin-bottom:8px"></div>
    <button class="btn btn-primary" style="width:100%" onclick="verifyPassword()">시스템 접속하기</button>
  </div>
</div>

<header class="top"><div class="wrap top-in">
  <a href="/" style="text-decoration:none;color:inherit" title="소개 페이지로"><div class="brand">로켓단<span>AI기반 간호실습 배치 의사결정 프로그램</span></div></a>
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
    <h2>배정 조건</h2>
    <p class="lead">배정 방식과 교과목을 선택하면 해당 교과목의 실습 병원 목록을 불러옵니다.</p>

    <div class="grid g2" style="margin-bottom:18px">
      <label class="mode on" id="modeG"><input type="radio" name="mode" value="global" checked onchange="onModeChange()">
        <b>전체 병원 동시 배정 <span class="pill hit" style="margin-left:4px">추천</span></b>
        <span>모든 학생과 병원을 한 번에 계산합니다. 학생은 병원 한 곳에만 배정됩니다.</span></label>
      <label class="mode" id="modeS"><input type="radio" name="mode" value="single" onchange="onModeChange()">
        <b>단일 병원 순위</b>
        <span>병원 한 곳을 선택해 지원자의 종합점수 순위를 매기고 정원까지 배정합니다.</span></label>
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
    <input type="checkbox" id="useTransit" checked hidden>

    <div class="hint" id="transitHint"></div>

    <details class="adv">
      <summary>상세 설정 — 배정 기준 · 실습 기간 · 병원별 정원 · 추가 조건</summary>
      <div class="adv-body">
        <h4>배정 기준</h4>
        <div class="wbar"><span id="wPrefLbl">희망 병원 30%</span><span id="wDistLbl">통학 거리 70%</span></div>
        <input type="range" id="wDist" min="0" max="100" step="5" value="70" oninput="syncW()" aria-label="통학 거리와 희망 병원의 비중">
        <input type="hidden" id="wPref" value="30">
        <div class="hint">왼쪽은 희망 병원, 오른쪽은 통학 거리의 비중이 커집니다.</div>

        <h4>실습 기간</h4>
        <div class="grid g2">
          <div><span class="lbl">실습 시작일</span><input type="date" id="startDate" onchange="onStartChange()"></div>
          <div><span class="lbl">실습 종료일 (2주 안 · 금요일 자동)</span><input type="date" id="endDate" onchange="checkDate(this);updateDayCount()"></div>
        </div><div class="hint" id="dateErr" style="color:var(--bad);font-weight:700"></div>
        <div class="hint" id="dayCount" style="font-weight:700"></div>

        <h4>병원별 정원</h4>
        <div class="hint" style="margin:0 0 10px">병원별 정원을 입력합니다. 비워 두면 아래 기본 정원을 적용합니다.</div>
        <div id="capTable" class="caps"></div>
        <div style="display:flex;gap:8px;margin-top:12px;flex-wrap:wrap;align-items:center">
          <button type="button" class="btn btn-ghost btn-sm" id="capActualBtn" onclick="fillCaps('actual')" style="display:none">2025-1 실제 인원으로 채우기</button>
          <button type="button" class="btn btn-ghost btn-sm" onclick="fillCaps('clear')">모두 비우기</button>
          <span class="sub" id="capSum"></span>
        </div>
        <div style="display:flex;gap:10px;align-items:center;margin-top:12px;font-size:14px">
          <span class="lbl" style="margin:0">기본 정원</span><input type="number" id="capHosp" value="40" min="1" style="width:84px;text-align:right">
          <span class="sub">명 (칸을 비운 병원에 적용)</span>
        </div>
        <input type="hidden" id="capPeriod" value="50">

        <details class="more">
          <summary>추가 조건 (성별 · GPA · 출생연도 · 과거 실습 제외)</summary>
          <div class="more-body">
            <div class="grid g3">
              <div><span class="lbl">성별 조건</span><select id="gender"><option>무관</option><option>남성만</option><option>여성만</option></select></div>
              <div><span class="lbl">최소 GPA</span><input type="number" step="0.1" id="minGpa" placeholder="예: 3.5"></div>
              <div><span class="lbl">출생연도 이후</span><input type="number" id="birthYear" placeholder="예: 2003"></div>
            </div>
            <label style="display:flex;gap:8px;align-items:center;font-size:14px;font-weight:600;cursor:pointer;margin-top:14px">
              <input type="checkbox" id="exPast" style="width:auto"> 과거 실습 기관 중복 배정 자동 제외</label>
          </div>
        </details>
      </div>
    </details>

    <div class="actions"><span></span><button class="btn btn-primary" onclick="goStep(2)">다음: 명단 업로드 →</button></div>
  </section>

  <!-- 2단계 -->
  <section id="step2" class="card step" style="display:none">
    <h2>학생 명단 업로드</h2>
    <p class="lead">파일은 메모리에서만 처리하며 서버에 저장하지 않습니다.</p>
    <div class="chips" id="condChips"></div>

    <div class="drop" id="drop">
      <div class="big">엑셀(.xlsx) 또는 CSV 파일을 이곳에 놓거나 선택하세요</div>
      <div class="sub" style="margin-bottom:14px">필수 열: 학번 · 이름 · 성별 · GPA · 출생연도 (나머지 열은 양식 샘플 참고)</div>
      <label class="btn btn-ghost btn-sm" style="cursor:pointer">파일 선택<input type="file" id="file" accept=".csv,.xlsx" style="display:none"></label>
      <div class="file-ok" id="fileOk"></div>
    </div>
    <div class="or">또는</div>
    <div class="tools">
      <button class="btn btn-ghost btn-sm" onclick="loadSample()">샘플 명단 불러오기</button>
      <a class="btn btn-ghost btn-sm" href="/api/v1/template">양식 샘플 다운로드</a>
    </div>

    <div class="actions">
      <button class="btn btn-ghost" onclick="goStep(1)">← 이전</button>
      <button class="btn btn-primary" id="runBtn" onclick="run()">배정 실행</button>
    </div>
  </section>

  <!-- 3단계 -->
  <section id="step3" class="step" style="display:none">
    <div class="res-top">
      <div class="meta" id="resMeta"></div>
      <div style="display:flex;gap:8px;flex-wrap:wrap">
        <button class="btn btn-ghost btn-sm" onclick="goStep(1)">조건 바꿔 다시 배정</button>
      </div>
    </div>

    <div id="warn"></div>
    <div id="hero"></div>
    <div class="kpis" id="kpis"></div>
    <div class="charts" id="charts"></div>

    <div class="card" style="padding:0;overflow:hidden">
      <div class="tabs">
        <button class="tab on" data-p="pA" onclick="showPane('pA')">배정 결과</button>
        <button class="tab" data-p="pH" onclick="showPane('pH')">병원별 현황</button>
        <button class="tab" data-p="pU" onclick="showPane('pU')">미배정·사유</button>
        <button class="btn btn-primary btn-sm xl" onclick="exportXlsx()">엑셀 다운로드</button>
        <input type="search" class="search" id="q" placeholder="이름·학번·병원 검색" oninput="PG=1;renderAssigned()">
      </div>
      <div id="pA" class="pane on tbl-wrap">
        <table class="tbl"><thead><tr>
          <th class="sortable" onclick="sortBy('rank')">순위<span class="ar" id="ar_rank"></span></th><th class="l sortable" onclick="sortBy('hospital')">병원<span class="ar" id="ar_hospital"></span></th><th class="sortable" onclick="sortBy('student_id')">학번<span class="ar" id="ar_student_id"></span></th><th class="sortable" onclick="sortBy('name')">이름<span class="ar" id="ar_name"></span></th><th class="sortable" onclick="sortBy('gpa')">GPA<span class="ar" id="ar_gpa"></span></th>
          <th class="sortable" onclick="sortBy('pref')" title="학생이 적은 1~3지망과 배정 병원의 일치 여부">지망 반영<span class="ar" id="ar_pref"></span></th>
          <th class="sortable" onclick="sortBy('time')" title="TMAP 대중교통 소요시간. '약'은 지도 좌표 기반 추정, '추정'은 주소 기반 규칙 추정">통학시간<span class="ar" id="ar_time"></span></th>
          <th class="sortable" onclick="sortBy('score')" title="통학 점수 × 거리 가중치 + 지망 점수 × 희망도 가중치">종합점수<span class="ar" id="ar_score"></span></th>
          <th class="sortable" onclick="sortBy('sat')" title="통학 부담(소요시간·환승·도보) 기반 규칙형 지수. 학습 모델이 아닙니다">만족도 지수<span class="ar" id="ar_sat"></span></th>
        </tr></thead><tbody id="bodyA"></tbody></table>
        <div class="pager" id="pager"></div>
        <div class="sub" style="padding:0 16px 12px">행을 누르면 점수 산출 내역을 볼 수 있습니다. 통학시간의 '약'은 지도 기반 예상 소요시간, '추정'은 규칙 기반 추정입니다. 열 제목을 누르면 정렬됩니다.</div>
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
function bizDays(a, b) {   // 시작일~종료일 사이 평일(공휴일 제외) 일수, 양끝 포함
  const d = pd_(a), end = pd_(b); let n = 0;
  while (d <= end) { if (!dateProblem(fmt(d))) n++; d.setDate(d.getDate() + 1); }
  return n;
}
function updateDayCount() {
  const a = $('startDate').value, b = $('endDate').value, el = $('dayCount');
  if (!a || !b) { el.textContent = ''; return; }
  const n = bizDays(a, b);
  if (n < 8) { el.style.color = 'var(--bad)'; el.textContent = `실습일수 ${n}일: 8일 미만이면 실습 이수가 되지 않습니다.`; }
  else if (n > 10) { el.style.color = 'var(--bad)'; el.textContent = `실습일수 ${n}일: 최대 10일을 넘습니다.`; }
  else { el.style.color = 'var(--ok, #0f766e)'; el.textContent = `실습일수 ${n}일 (8~10일 충족)`; }
}
function defaultEnd(s) {   // 실습은 2주 안에 끝나고 보통 금요일: 시작일이 속한 주의 '다음 주 금요일'
  const d = pd_(s); d.setDate(d.getDate() - ((d.getDay() + 6) % 7) + 11);
  let e = fmt(d); while (dateProblem(e)) { d.setDate(d.getDate() - 1); e = fmt(d); }   // 그날이 휴일이면 그 전 평일
  return e;
}
function addBiz(s, n) { const d = pd_(s); let c = 0; while (c < n) { d.setDate(d.getDate() + 1); if (!dateProblem(fmt(d))) c++; } return fmt(d); }
function initDates() {
  const d = new Date();   // 시작일 기본값 = 오늘 (주말·공휴일이면 다음 평일)
  let s = fmt(d); while (dateProblem(s)) { d.setDate(d.getDate() + 1); s = fmt(d); }
  if (bizDays(s, defaultEnd(s)) < 8) {   // 오늘 시작하면 8일을 못 채우는 요일(목·금 등)이면 다음 월요일부터
    d.setDate(d.getDate() + ((8 - d.getDay()) % 7 || 7)); s = fmt(d); while (dateProblem(s)) { d.setDate(d.getDate() + 1); s = fmt(d); }
  }
  $('startDate').value = s; $('endDate').value = defaultEnd(s); updateDayCount();
}
function checkDate(inp) { const m = inp.value ? dateProblem(inp.value) : null; $('dateErr').textContent = m ? m : ''; if (m) inp.value = ''; return !m; }
function onStartChange() { const s = $('startDate'); if (checkDate(s) && s.value) $('endDate').value = defaultEnd(s.value); updateDayCount(); }
function syncW() {
  const a = $('wDist'), b = $('wPref'), d = Math.max(0, Math.min(100, parseInt(a.value) || 0));
  a.value = d; b.value = 100 - d;
  $('wDistLbl').textContent = `통학 거리 ${d}%`; $('wPrefLbl').textContent = `희망 병원 ${100 - d}%`;
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
let ACTUAL_CAPS = {};
function buildCapTable(hs) {
  const box = $('capTable'); box.innerHTML = '';
  hs.forEach(h => {
    const w = document.createElement('div'); w.className = 'caprow';
    w.innerHTML = `<span>${esc(h.replace('(방중)', ''))}</span><input type="number" min="1" class="capIn" data-h="${esc(h)}" placeholder="기본" oninput="capSum()">`;
    box.appendChild(w);
  });
  $('capActualBtn').style.display = Object.keys(ACTUAL_CAPS).length ? '' : 'none';
  fillCaps(Object.keys(ACTUAL_CAPS).length ? 'actual' : 'clear');
}
function fillCaps(kind) {
  document.querySelectorAll('.capIn').forEach(i => { i.value = kind === 'actual' ? (ACTUAL_CAPS[i.dataset.h] ?? '') : ''; });
  capSum();
}
function capSum() {
  let t = 0, n = 0;
  document.querySelectorAll('.capIn').forEach(i => { if (i.value !== '') { t += parseInt(i.value) || 0; n++; } });
  $('capSum').textContent = n ? `입력된 ${n}곳 합계 ${t}명` : '';
}
function capsJson() {
  const o = {};
  document.querySelectorAll('.capIn').forEach(i => { if (i.value !== '' && parseInt(i.value) >= 1) o[i.dataset.h] = parseInt(i.value); });
  return Object.keys(o).length ? JSON.stringify(o) : '';
}
async function updateHospitals() {
  const q = new URLSearchParams({grade: $('grade').value, semester: $('semester').value, subject: $('subject').value});
  const box = $('hospital'); box.innerHTML = '';
  try {
    const d = await (await fetch('/api/v1/hospitals?' + q)).json();
    (d.hospitals || []).forEach(h => box.add(new Option(h, h)));
    $('hospCount').textContent = `이 교과목의 실습 병원 ${(d.hospitals || []).length}곳`;
    ACTUAL_CAPS = d.actual_caps || {};
    buildCapTable(d.hospitals || []);
  } catch (e) { $('hospCount').textContent = '병원 목록을 불러오지 못했습니다.'; }
}
window.addEventListener('load', async () => {
  initDates();
  try { CURRICULUM = await (await fetch('/api/v1/curriculum')).json(); } catch (e) {}
  try { CONFIG = await (await fetch('/api/v1/config')).json(); } catch (e) { CONFIG = {}; }
  loadNaverMap();
  if (CONFIG.auth_required && !TOKEN) $('auth').style.display = 'flex';
  if (CONFIG.transit_available) {
    const m = CONFIG.hospitals_missing_address;
    $('transitHint').textContent = '통학시간: ' + (CONFIG.transit_mode === 'tmap' ? 'TMAP 대중교통 소요시간 사용(' + (CONFIG.commute_label || '평일 오전 6시 출발') + ' 기준). ' : '지도 기반 예상 소요시간. ')
      + (m ? `병원 주소가 없는 ${m}곳은 규칙 기반 추정치를 씁니다. ` : '') + '첫 실행은 시간이 걸리며, 이후에는 저장된 값을 사용합니다.';
  } else {
    $('transitHint').textContent = '지도 키가 설정되지 않아 통학시간은 규칙 기반 추정치로 계산합니다.';
  }
  syncW();
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
  ok.style.display = 'block'; ok.textContent = `${f.name} (${Math.max(1, Math.round(f.size / 1024))}KB)`;
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
    toast('가상 학생 60명을 불러왔습니다.');
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
  f.append('max_period_capacity', $('capPeriod').value || 50); f.append('max_hospital_capacity',"""
AI기반 간호실습 배치 의사결정 프로그램

버전 기록 (수정할 때마다 APP_VERSION 을 올립니다)
  v13.4 소개 페이지·배정 도구 전체 새 디자인(병원 라벨 시트 디자인 체계). 계산·기능·문구·팀 캐릭터는 그대로
  v13.3 소개 페이지 디자인을 v12.2로 되돌리고 "만든 사람들"(팀 캐릭터)과 맨 위 버전 표시만 유지
  v13.2 대중교통 조회 기준을 평일 오전 6시 출발로 변경(07시 실습 시작 기준, 환경변수 COMMUTE_TIME으로 조절)
  v13.1 소개 페이지 맨 위에 버전 표시
  v13.0 소개 페이지 새 디자인(팀 캐릭터·실제 화면·개발 이유, 포스터와 같은 파란색)
  v12.2 소개 페이지 정리(업데이트 내역 제거, 예시 표 단순화, 버튼 문구 통일, 개선 효과·팀 정보 표시)
  v12.1 화면 문구에서 "직선거리" 표현 정리(통학은 "지도 기반 예상 소요시간"으로 표시)
  v12.0 프로그램 이름을 "AI기반 간호실습 배치 의사결정 프로그램"으로 변경(팀 이름 로켓단은 유지)
  v11.9 소개 페이지 구역마다 화면 높이에 맞게 꽉 채우고 가운데 정렬(영상 촬영용)
  v11.8 소개 페이지 숫자 키(1~9)로 구역 이동
  v11.7 소개 페이지 오른쪽 점 모양 구역 안내(PC 화면)
  v11.6 소개 페이지 순서 변경(소개 → 서비스 시작하기 → 사용법 → 순서), 자주 묻는 질문 추가
  v11.5 소개 페이지: 화면 예시, 거리·희망 비중 체험, 수작업과 비교, 숫자 올라가는 효과
  v11.4 소식(업데이트) 맨 위에 현재 버전 표시
  v11.3 소개 페이지 첫 문장을 짧게(한 줄로 이어지도록)
  v11.2 소개 페이지 줄바꿈 정리(단어 중간에서 끊기지 않고 한 줄에 가깝게)
  v11.1 소개 페이지를 홈페이지 형태로 개편(상단 메뉴·Why·숫자·서비스·소식), 사용설명서 링크(GUIDE_URL 환경변수)
  v11.0 소개 페이지(/) 추가, 배정 도구는 /app 으로 이동, 검색용 정보(robots.txt·sitemap.xml)
  v10.9 지도 문구 정리(거주지 표시, 동 기준 위치 안내)
  v10.8 화면에서 이름(가운데 글자)·학번(가운데 3자리) 가림 처리(엑셀 다운로드는 원본 유지)
  v10.7 지도에 거주지-병원 직선과 직선거리 다시 표시(v10.5와 같은 화면)
  v10.6 지도에서 직선·거리 표시 제거(거주지와 병원 두 위치만 표시)
  v10.5 거주지-병원 지도를 네이버 지도로 변경(Dynamic Map 미설정 시 기존 지도로 대체)
  v10.4 자문 반영: 문구·이모지 정리, 엑셀 버튼 위치, 20명 단위 페이지, 열 정렬, 거주지-병원 지도(상세·다른 병원)
  v10.3 실습일수 표시(8~10일 규칙): 8일 미만이면 경고, 오늘이 목·금이면 시작일을 다음 월요일로
  v10.2 종료일 기본값: 시작일이 속한 주의 다음 주 금요일(휴일이면 그 전 평일)
  v10.1 실습 기간 기본값: 시작일=오늘(휴일이면 다음 평일)
  v10.0 화면 정리(상세 설정 재구성, 지도 거리 항상 사용, 로고 아이콘·샘플 버튼 정리, 통학시간 "약" 표기)
  v9.9  병원별 정원을 화면에서 직접 입력(기본값: 2025-1 실제 배정 인원)
  v9.8  병원별 실제 정원(2025-1 실제 배정 인원) 사용 옵션
  v9.7  지도가 못 찾은 주소를 결과 경고에 표시
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
import base64
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
from fastapi import Body, FastAPI, File, Form, Header, HTTPException, Request, UploadFile
from fastapi.concurrency import run_in_threadpool
from fastapi.responses import HTMLResponse, PlainTextResponse, Response, StreamingResponse
from pydantic import BaseModel
from scipy.optimize import linear_sum_assignment

APP_VERSION = "v13.4"
app = FastAPI(title=f"AI기반 간호실습 배치 의사결정 프로그램 ({APP_VERSION})")

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
# 학과의 실제 배정 인원(2025-1 4학년 관리 실습). 병원마다 받는 인원이 다르므로 '병원별 실제 정원'으로 쓸 수 있다.
ACTUAL_CAPS = {
    ("4학년", "1학기", "관리"): {
        "순천향대학교부천병원": 36, "순천향대학교서울병원": 36, "단원병원": 28, "아주대학교병원": 24,
        "고려대학교구로병원": 24, "시화병원": 23, "한림대학교동탄성심병원": 23, "서울특별시서남병원": 19,
        "중앙대학교광명병원": 8,
    },
}


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
# 대중교통 조회 출발 시각(HHMM). 실습 07:00 시작 → 06:00 출발 기준. Render 환경변수 COMMUTE_TIME 으로 변경 가능
COMMUTE_TIME = (os.environ.get("COMMUTE_TIME", "0600").strip().replace(":", "") or "0600").zfill(4)[:4]
if not (COMMUTE_TIME.isdigit() and int(COMMUTE_TIME[:2]) < 24 and int(COMMUTE_TIME[2:]) < 60):
    COMMUTE_TIME = "0600"
COMMUTE_LABEL = f"평일 오전 {int(COMMUTE_TIME[:2])}시" + (f" {int(COMMUTE_TIME[2:])}분" if COMMUTE_TIME[2:] != "00" else "") + " 출발"
LINE_BASE_MIN, LINE_MIN_PER_KM = 12, 2.0   # 직선거리 추정식: 통학(분) = 12 + 2.0 × 직선거리(km)  (실측과 비교해 조정 가능)


def _rkey(dong: str, hosp_addr: str) -> str:
    tag = f"{TRANSIT_MODE}@{COMMUTE_TIME}" if TRANSIT_MODE == "tmap" else TRANSIT_MODE   # 출발 시각이 바뀌면 다시 조회
    return f"{tag}|{dong}||{hosp_addr}"


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
    """출근 시간대 기준(다음 평일 COMMUTE_TIME, 한국시간). TMAP searchDttm 형식 yyyymmddhhmi."""
    kst = datetime.timezone(datetime.timedelta(hours=9))
    d = datetime.datetime.now(kst).date() + datetime.timedelta(days=1)
    while d.weekday() >= 5:
        d += datetime.timedelta(days=1)
    return d.strftime("%Y%m%d") + COMMUTE_TIME


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


def cached_ll(query: str) -> Optional[List[float]]:
    """이미 조회해 둔 주소의 [경도, 위도]. 새로 조회하지 않는다(없으면 None)."""
    if not query:
        return None
    v = _load_cache()["geo"].get(query)
    return [v[0], v[1]] if v else None


def hosp_ll(hospital: str) -> Optional[List[float]]:
    return cached_ll(_ADDR_MAP.get(norm_h(hospital), ""))


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
        with _cache_lock:
            bad = sorted({q for d, a in todo for q in (d, a) if q in c["geo"] and c["geo"][q] is None})
        if bad:
            notes.append(f"지도가 위치를 찾지 못한 주소 {len(bad)}개: " + " / ".join(bad[:8]) + (" 외" if len(bad) > 8 else ""))
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
        "h_ll": hosp_ll(h),
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
                    valid: List[str], cap) -> Optional[dict]:
    """기준선: 정원은 지키되 학생을 무작위로 병원에 넣는 방식의 평균(고정 시드, 반복 시뮬레이션)."""
    n = len(eligible)
    if not n:
        return None
    rng = np.random.default_rng(20260101)
    trials = 100 if n <= 500 else 20
    picked: List[dict] = []
    for _ in range(trials):
        remaining = dict(cap) if isinstance(cap, dict) else {h: cap for h in valid}
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
        "s_ll": cached_ll(stu.get("dong", "")),
        "h_ll": hosp_ll(hospital),
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
    use_actual_caps: bool = False,
    hospital_caps: Optional[Dict[str, int]] = None,
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
        cap_by_h = {h: cap_eff for h in valid}
        actual_used = False
        custom_used = False
        if hospital_caps:
            hn = {}
            for k, v in hospital_caps.items():
                try:
                    if int(v) >= 1:
                        hn[norm_h(k)] = int(v)
                except (TypeError, ValueError):
                    pass
            if hn:
                cap_by_h = {h: hn.get(norm_h(h), cap_eff) for h in valid}
                custom_used = True
        elif use_actual_caps:
            table = ACTUAL_CAPS.get((grade, semester, subject))
            if table:
                tn = {norm_h(k): v for k, v in table.items()}
                cap_by_h = {h: tn.get(norm_h(h), cap_eff) for h in valid}
                actual_used = True
            else:
                warnings.append("이 교과목은 실제 정원 자료가 없어 '기관별 최대정원' 값을 모든 병원에 똑같이 적용했습니다.")
        slot_hosp = [h for h in valid for _ in range(cap_by_h[h])]
        slots_total = len(slot_hosp)
        hospital_cap_map = dict(cap_by_h)
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
            slot_idx = np.repeat(np.arange(len(valid)), [cap_by_h[h] for h in valid])
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
                random_baseline(eligible, pairs, allowed, valid, cap_by_h),
                "무작위 배정 — 정원은 지키되 학생을 임의로 배치 (100회 시뮬레이션 평균)")
        method = (f"전체 병원 동시 배정 · SciPy 헝가리안 전역 최적화 "
                  f"(거리 {round(dw*100)}% · 희망도 {round(pw*100)}%, "
                  + ("병원별 지정 정원" if custom_used else "병원별 실제 정원(2025-1 배정 인원)" if actual_used else f"병원당 정원 {cap_eff}명")
                  + f", 병원 {len(valid)}곳)")
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
    return {"transit_available": TRANSIT_AVAILABLE, "transit_mode": TRANSIT_MODE, "hospitals_missing_address": missing, "auth_required": AUTH_REQUIRED, "version": APP_VERSION, "map_key_id": NAVER_KEY_ID or "", "commute_time": COMMUTE_TIME, "commute_label": COMMUTE_LABEL}


@app.get("/api/v1/curriculum")
async def get_curriculum():
    return {g: {s: list(subs.keys()) for s, subs in sems.items()} for g, sems in MASTER_HOSPITAL_DB.items()}


def _parse_caps(raw: str) -> Optional[Dict[str, int]]:
    try:
        d = json.loads(raw) if raw else None
        return {str(k): int(v) for k, v in d.items()} if isinstance(d, dict) else None
    except Exception:
        return None


@app.get("/api/v1/hospitals")
async def get_hospitals(grade: str, semester: str, subject: str):
    hs = hospitals_for(grade, semester, subject)
    table = ACTUAL_CAPS.get((grade, semester, subject)) or {}
    tn = {norm_h(k): v for k, v in table.items()}
    return {"hospitals": hs, "actual_caps": {h: tn[norm_h(h)] for h in hs if norm_h(h) in tn}}


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
    use_actual_caps: bool = Form(False),
    hospital_caps: str = Form(""),
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
            use_transit_api=use_transit_api, use_actual_caps=use_actual_caps,
            hospital_caps=_parse_caps(hospital_caps),
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
<title>AI기반 간호실습 배치 의사결정 프로그램 | 로켓단</title>
<meta name="robots" content="noindex">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css" media="print" onload="this.media='all'">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.css">
<script src="https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.js"></script>
<style>
  @font-face { font-family:'BMJUA'; src:url(/assets/jua.woff2) format('woff2'); font-weight:400; font-style:normal; font-display:swap; }
  :root {
    --ink:#0C1A3A; --ink-2:#2B3A5E; --muted:#52617F;
    --blue:#3E5BB8; --blue-d:#2D4590; --blue-l:#EAEFFB; --blue-m:#8EA2E0;
    --liner:#F1F4F9; --hair:rgba(12,26,58,.12);
    --gold:#FFD34D; --gold-d:#6B5000; --gold-l:#FFF6D6;
    --green:#0B7A5A; --green-l:#DCF4EA; --red:#C62A44; --red-l:#FDE3E8;
    --r:14px; --r-s:8px;
    --lift:0 1px 0 rgba(12,26,58,.06), 0 14px 30px rgba(20,32,80,.22);
    --lift-s:0 1px 0 rgba(12,26,58,.06), 0 8px 20px rgba(20,32,80,.09);
    --ease:cubic-bezier(.16,1,.3,1);
    --font:'BMJUA','Pretendard Variable','Pretendard',-apple-system,'Apple SD Gothic Neo','Noto Sans KR','Malgun Gothic',system-ui,sans-serif;
    /* 예전 이름(화면 스크립트가 쓰는 변수)을 새 색에 연결 */
    --navy:var(--ink); --teal:var(--blue); --teal-d:var(--blue-d); --amber:var(--gold);
    --bg:var(--liner); --card:#fff; --line:var(--hair); --text:var(--ink); --good:var(--green); --bad:var(--red); --warn:var(--gold-d);
  }
  * { box-sizing:border-box; }
  html { font-synthesis:none; -webkit-text-size-adjust:100%; scrollbar-color:var(--blue) var(--liner); }
  body { margin:0; font-family:var(--font); background:var(--liner); color:var(--ink); line-height:1.55; font-size:15px; word-break:keep-all; overflow-wrap:break-word; }
  ::selection { background:var(--gold); color:var(--ink); }
  :focus-visible { outline:3px solid var(--gold); outline-offset:2px; box-shadow:0 0 0 5px var(--ink); border-radius:6px; }
  button, input, select { font:inherit; color:inherit; caret-color:var(--blue); }
  input[type=checkbox], input[type=radio], input[type=range] { accent-color:var(--blue); }
  h2, h3 { text-wrap:balance; }
  .wrap { max-width:1240px; margin:0 auto; padding:0 28px; }
  .num, .tbl td, .kpi .v, .bignum .n { font-variant-numeric:tabular-nums; }

  /* 상단: 파란 라벨 시트 위에 단계가 놓임 */
  .top { background:var(--blue-l); color:var(--ink); }
  .top-in { display:flex; align-items:center; justify-content:space-between; gap:14px; padding-top:18px; padding-bottom:16px; }
  .brand { display:flex; align-items:center; gap:14px; text-decoration:none; color:var(--ink); min-width:0; }
  .brand b { background:var(--blue); color:#fff; font-size:20px; font-weight:900; letter-spacing:-.01em; padding:7px 15px 8px; border-radius:10px; box-shadow:0 3px 0 var(--blue-d); white-space:nowrap; }
  .brand span { font-size:14.5px; font-weight:650; color:var(--muted); letter-spacing:-.01em; }
  .ver { font-size:12.5px; font-weight:800; color:var(--blue-d); background:#fff; padding:4px 11px; border-radius:99px; white-space:nowrap; }

  /* 단계: 상태는 색이 아니라 표시(완료=체크, 현재=노란 번호) */
  .stepper { list-style:none; display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:12px; padding:0 0 26px; margin:0; }
  .stepper li { display:flex; align-items:center; gap:12px; padding:13px 18px; border-radius:12px; border:2px dashed rgba(12,26,58,.28); color:var(--muted); font-weight:750; font-size:15.5px; letter-spacing:-.01em; cursor:default; transition:background .2s, transform .2s var(--ease); }
  .stepper li i { font-style:normal; width:30px; height:30px; flex:none; border-radius:50%; border:2px solid currentColor; display:grid; place-items:center; font-size:14px; font-weight:900; }
  .stepper li.done { cursor:pointer; background:rgba(255,255,255,.6); border:2px solid rgba(62,91,184,.35); color:var(--ink); }
  .stepper li.done:hover { background:#fff; transform:translateY(-1px); }
  .stepper li.done i { background:var(--blue); border-color:var(--blue); color:transparent; font-size:0; position:relative; }
  .stepper li.done i::after { content:''; width:11px; height:6px; border-left:3px solid #fff; border-bottom:3px solid #fff; transform:translateY(-1px) rotate(-45deg); }
  .stepper li.active { background:#fff; color:var(--ink); border:2px solid var(--blue); box-shadow:var(--lift-s); }
  .stepper li.active i { background:var(--gold); border-color:var(--ink); color:var(--ink); }

  main.wrap { padding-top:26px; }

  /* 카드(라벨) */
  .card { background:#fff; border-radius:var(--r); box-shadow:var(--lift-s); padding:34px 36px 32px; }
  .card h2 { margin:0 0 6px; font-size:25px; font-weight:850; letter-spacing:-.01em; color:var(--ink); }
  .lead { margin:0 0 24px; color:var(--muted); font-size:15.5px; }
  .field > label, .lbl { display:block; font-size:13.5px; font-weight:750; color:var(--ink-2); margin-bottom:7px; letter-spacing:-.01em; }
  select, input[type=number], input[type=date], input[type=search], input[type=password] { width:100%; min-height:46px; padding:10px 14px; border:1.5px solid rgba(12,26,58,.24); border-radius:10px; background:#fff; font-size:15.5px; transition:border-color .15s, box-shadow .15s; }
  select { appearance:none; padding-right:40px; background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='14' height='9' viewBox='0 0 14 9'%3E%3Cpath d='M1.5 1.5l5.5 5.5 5.5-5.5' fill='none' stroke='%230C1A3A' stroke-width='2.2' stroke-linecap='round' stroke-linejoin='round'/%3E%3C/svg%3E"); background-repeat:no-repeat; background-position:right 15px center; cursor:pointer; }
  select:hover, input:hover { border-color:rgba(12,26,58,.5); }
  select:focus, input:focus { outline:none; border-color:var(--blue); box-shadow:0 0 0 4px var(--blue-l); }
  input::placeholder { color:#7C8AA8; }
  .grid { display:grid; gap:16px; }
  .g2 { grid-template-columns:1fr 1fr; } .g3 { grid-template-columns:repeat(3,1fr); }
  .hint { font-size:13px; color:var(--muted); margin-top:7px; }
  .sub { color:var(--muted); font-size:12.5px; }

  /* 버튼: 노란 라벨(주요) / 흰 윤곽(보조) */
  .btn { border:2px solid var(--ink); border-radius:12px; padding:0 24px; min-height:48px; font-weight:800; font-size:15.5px; letter-spacing:-.01em; cursor:pointer; display:inline-flex; align-items:center; justify-content:center; gap:8px; text-decoration:none; box-shadow:0 3px 0 var(--ink); transition:transform .14s var(--ease), box-shadow .14s var(--ease), background .14s; }
  .btn:hover { transform:translateY(-2px); box-shadow:0 5px 0 var(--ink); }
  .btn:active { transform:translateY(2px); box-shadow:0 1px 0 var(--ink); }
  .btn-primary { background:var(--gold); color:var(--ink); }
  .btn-primary:disabled { opacity:.75; cursor:wait; transform:none; box-shadow:0 3px 0 var(--ink); }
  .btn-ghost { background:#fff; color:var(--ink); }
  .btn-ghost:hover { background:var(--blue-l); }
  .btn-sm { min-height:38px; padding:0 16px; font-size:13.5px; border-radius:10px; box-shadow:0 2px 0 var(--ink); }
  .btn-sm:hover { box-shadow:0 4px 0 var(--ink); }
  .btn.busy::before { content:''; width:16px; height:16px; border-radius:50%; border:3px solid rgba(12,26,58,.25); border-top-color:var(--ink); animation:spin .8s linear infinite; }
  @keyframes sheetUp { from { transform:translateY(40px); opacity:0; } to { transform:none; opacity:1; } }
  @keyframes spin { to { transform:rotate(360deg); } }
  .actions { display:flex; justify-content:space-between; gap:12px; margin-top:30px; flex-wrap:wrap; }

  /* 배정 방식 */
  .mode { display:block; position:relative; border:2px solid rgba(12,26,58,.18); border-radius:var(--r); padding:20px 56px 20px 22px; cursor:pointer; transition:border-color .15s, background .15s, transform .15s var(--ease); height:100%; background:#fff; }
  .mode:hover { border-color:var(--blue); transform:translateY(-1px); }
  .mode input { position:absolute; opacity:0; pointer-events:none; }
  .mode b { display:block; font-size:17px; font-weight:800; letter-spacing:-.01em; color:var(--ink); margin-bottom:4px; }
  .mode > span { font-size:14px; color:var(--muted); display:block; }
  .mode::after { content:''; position:absolute; right:20px; top:20px; width:26px; height:26px; border-radius:50%; border:2px solid rgba(12,26,58,.28); background:#fff; }
  .mode.on { border-color:var(--blue); background:var(--blue-l); box-shadow:0 0 0 3px rgba(62,91,184,.14); }
  .mode.on::after { border-color:var(--blue); background:var(--blue); }
  .mode.on::before { content:''; position:absolute; right:29px; top:28px; width:9px; height:5px; border-left:3px solid #fff; border-bottom:3px solid #fff; transform:rotate(-45deg); z-index:1; }
  .mode:focus-within { box-shadow:0 0 0 4px var(--gold), 0 0 0 6px var(--ink); }
  .mode b .pill { vertical-align:2px; }

  details.adv { margin-top:24px; border:2px solid rgba(12,26,58,.14); border-radius:var(--r); background:#fff; }
  details.adv > summary { cursor:pointer; font-weight:800; font-size:16px; letter-spacing:-.01em; padding:17px 22px; list-style:none; color:var(--ink); display:flex; justify-content:space-between; align-items:center; gap:14px; border-radius:var(--r); }
  details.adv > summary:hover { background:var(--blue-l); }
  details.adv > summary::-webkit-details-marker, details.more > summary::-webkit-details-marker { display:none; }
  details.adv > summary::after, details.more > summary::after { content:''; flex:none; width:26px; height:26px; border-radius:50%; background:linear-gradient(var(--ink),var(--ink)) center/11px 2px no-repeat, linear-gradient(var(--ink),var(--ink)) center/2px 11px no-repeat, var(--blue-l); transition:transform .3s var(--ease); float:none; }
  details.adv[open] > summary::after, details.more[open] > summary::after { transform:rotate(45deg); }
  details.adv .adv-body { padding:6px 22px 24px; border-top:1px solid var(--hair); }
  .sec { margin-top:18px; }
  .adv-body h4 { margin:28px 0 10px; font-size:16px; font-weight:800; letter-spacing:-.01em; color:var(--ink); }
  .adv-body h4:first-child { margin-top:20px; }
  details.more { margin-top:26px; border:2px dashed rgba(12,26,58,.22); border-radius:12px; }
  details.more > summary { cursor:pointer; padding:13px 18px; font-weight:700; font-size:14.5px; color:var(--ink-2); list-style:none; display:flex; justify-content:space-between; align-items:center; gap:12px; border-radius:12px; }
  details.more .more-body { padding:4px 18px 20px; }
  .wbar { display:flex; justify-content:space-between; font-weight:800; font-size:16px; margin-bottom:2px; letter-spacing:-.01em; }
  .wbar span:first-child { color:var(--blue-d); }
  .wbar span:last-child { color:var(--gold-d); }
  input[type=range] { --v:.7; -webkit-appearance:none; appearance:none; width:100%; height:38px; background:transparent; cursor:pointer; margin:6px 0 0; padding:0; border:0; min-height:0; }
  input[type=range]:hover { border:0; }
  input[type=range]::-webkit-slider-runnable-track { height:12px; border-radius:99px; background:linear-gradient(90deg,var(--blue) 0 calc(14px + (100% - 28px) * var(--v)),var(--gold) calc(14px + (100% - 28px) * var(--v)) 100%); }
  input[type=range]::-moz-range-track { height:12px; border-radius:99px; background:linear-gradient(90deg,var(--blue) 0 calc(14px + (100% - 28px) * var(--v)),var(--gold) calc(14px + (100% - 28px) * var(--v)) 100%); }
  input[type=range]::-webkit-slider-thumb { -webkit-appearance:none; width:30px; height:30px; margin-top:-9px; border-radius:50%; background:#fff; border:4px solid var(--ink); box-shadow:0 3px 8px rgba(12,26,58,.3); transition:transform .15s var(--ease); }
  input[type=range]::-moz-range-thumb { width:22px; height:22px; border-radius:50%; background:#fff; border:4px solid var(--ink); box-shadow:0 3px 8px rgba(12,26,58,.3); }
  input[type=range]:hover::-webkit-slider-thumb { transform:scale(1.1); }
  input[type=range]:active::-webkit-slider-thumb { transform:scale(.94); }
  input[type=range]:focus, input[type=range]:focus-visible { outline:none; box-shadow:none; border:0; }
  input[type=range]:focus-visible::-webkit-slider-thumb { box-shadow:0 0 0 4px var(--gold), 0 0 0 7px var(--ink); }
  .caps { display:grid; grid-template-columns:repeat(auto-fill,minmax(270px,1fr)); gap:10px 22px; }
  .caprow { display:flex; align-items:center; justify-content:space-between; gap:12px; font-size:14.5px; font-weight:600; color:var(--ink-2); }
  .caprow input { width:92px; min-height:40px; padding:6px 10px; text-align:right; font-variant-numeric:tabular-nums; }

  /* 업로드 */
  .drop { border:2.5px dashed rgba(12,26,58,.3); border-radius:var(--r); background:var(--liner); padding:44px 22px; text-align:center; transition:border-color .15s, background .15s; }
  .drop:hover { border-color:var(--blue); }
  .drop.over { border-color:var(--blue); background:var(--blue-l); }
  .drop .big { font-size:19px; font-weight:800; letter-spacing:-.01em; color:var(--ink); margin-bottom:6px; }
  .drop .file-ok { display:none; margin-top:16px; font-weight:800; color:var(--green); background:var(--green-l); border-radius:99px; padding:6px 16px; width:fit-content; margin-left:auto; margin-right:auto; }
  .chips { display:flex; flex-wrap:wrap; gap:8px; margin-bottom:20px; }
  .chip { background:var(--blue-l); color:var(--blue-d); padding:6px 14px; border-radius:99px; font-size:13.5px; font-weight:750; letter-spacing:-.01em; }
  .or { display:flex; align-items:center; gap:14px; color:var(--muted); font-size:13.5px; font-weight:700; margin:22px 0; }
  .or::before, .or::after { content:''; height:0; border-top:2px dashed rgba(12,26,58,.18); flex:1; }
  .tools { display:flex; flex-wrap:wrap; gap:12px; justify-content:center; }

  /* 결과 */
  .res-top { display:flex; justify-content:space-between; align-items:flex-end; gap:12px; flex-wrap:wrap; margin-bottom:16px; }
  .res-top .meta { font-size:15px; color:var(--ink-2); }
  .res-top .meta b { color:var(--ink); }
  .hero { background:var(--blue); color:#fff; border-radius:var(--r); padding:32px 34px; display:grid; grid-template-columns:1fr 1.25fr; gap:36px; align-items:center; box-shadow:var(--lift); }
  .hero-t { margin:0 0 18px; font-size:20px; font-weight:850; letter-spacing:-.01em; color:#fff; }
  .bignums { display:flex; gap:30px; flex-wrap:wrap; }
  .bignum .n { font-size:46px; font-weight:900; line-height:1.05; letter-spacing:-.01em; }
  .bignum .n small { font-size:17px; font-weight:800; margin-left:2px; letter-spacing:-.01em; }
  .bignum .t { font-size:13.5px; color:#E3E9FA; margin-top:6px; font-weight:650; }
  .pos { color:var(--gold); } .neg { color:#FFC2CC; }
  .cmp-row { margin-bottom:16px; }
  .cmp-row:last-of-type { margin-bottom:6px; }
  .cmp-row .cl { font-size:13px; color:#E3E9FA; margin-bottom:6px; font-weight:700; }
  .pbar { display:flex; align-items:center; gap:12px; margin-top:5px; font-size:13px; }
  .pbar .tag { width:36px; color:#E3E9FA; flex:none; font-weight:700; }
  .pbar .track { flex:1; height:14px; border-radius:99px; background:rgba(255,255,255,.2); overflow:hidden; }
  .pbar .fill { height:100%; border-radius:99px; transition:width .9s var(--ease); }
  .pbar .fill.base { background:rgba(255,255,255,.55); } .pbar .fill.ours { background:var(--gold); }
  .pbar .val { width:68px; text-align:right; font-weight:800; flex:none; font-variant-numeric:tabular-nums; }
  .cmp-note { font-size:12.5px; color:#E3E9FA; margin-top:12px; }

  .kpis { display:grid; grid-template-columns:repeat(4,1fr); gap:12px; margin:12px 0; }
  .kpi { background:#fff; border-radius:var(--r); padding:20px 22px 18px; box-shadow:var(--lift-s); }
  .kpi .k { font-size:13.5px; color:var(--muted); font-weight:700; }
  .kpi .v { font-size:32px; font-weight:900; color:var(--ink); letter-spacing:-.01em; margin-top:2px; }
  .kpi .v small { font-size:14px; font-weight:700; color:var(--muted); margin-left:4px; letter-spacing:-.01em; }

  .charts { display:grid; grid-template-columns:repeat(auto-fit,minmax(310px,1fr)); gap:12px; margin-bottom:12px; }
  .chart { background:#fff; border-radius:var(--r); box-shadow:var(--lift-s); padding:24px 26px; }
  .chart h3 { margin:0 0 18px; font-size:17px; font-weight:850; letter-spacing:-.01em; color:var(--ink); }
  .hrow { display:grid; grid-template-columns:minmax(96px,38%) 1fr 48px; align-items:center; gap:10px; margin-bottom:10px; font-size:12.5px; }
  .hrow .nm { overflow:hidden; text-overflow:ellipsis; white-space:nowrap; color:var(--ink-2); font-weight:600; }
  .hrow .tr { height:14px; background:var(--liner); border-radius:99px; overflow:hidden; }
  .hrow .fl { height:100%; background:var(--blue); border-radius:99px; }
  .hrow .fl.full { background:var(--ink); }
  .hrow .ct { text-align:right; font-weight:800; color:var(--ink); font-variant-numeric:tabular-nums; }
  .vbars { display:flex; align-items:flex-end; gap:10px; height:156px; padding-top:6px; }
  .vb { flex:1; display:flex; flex-direction:column; justify-content:flex-end; align-items:center; height:100%; font-size:12.5px; }
  .vb .bar { width:100%; background:var(--blue); border-radius:8px 8px 0 0; min-height:3px; }
  .vb .n { font-weight:800; color:var(--ink); margin-bottom:4px; font-variant-numeric:tabular-nums; }
  .vlabels { display:flex; gap:10px; font-size:11.5px; color:var(--muted); margin-top:8px; font-weight:600; }
  .vlabels span { flex:1; text-align:center; }
  .donut-wrap { display:flex; align-items:center; gap:26px; flex-wrap:wrap; justify-content:center; }
  .donut { width:156px; height:156px; border-radius:50%; position:relative; flex:none; }
  .donut-hole { position:absolute; inset:26px; background:#fff; border-radius:50%; display:flex; flex-direction:column; align-items:center; justify-content:center; }
  .donut-hole b { font-size:28px; font-weight:900; letter-spacing:-.01em; color:var(--ink); line-height:1; }
  .donut-hole span { font-size:11.5px; color:var(--muted); margin-top:4px; font-weight:650; }
  .legend { font-size:13.5px; font-weight:600; color:var(--ink-2); } .legend div { display:flex; align-items:center; gap:9px; margin:7px 0; }
  .legend i { width:13px; height:13px; border-radius:4px; display:inline-block; }

  .alert { background:var(--gold-l); border:1.5px solid #F0D06A; color:var(--gold-d); border-radius:12px; padding:14px 18px; margin-bottom:14px; font-size:13.5px; font-weight:600; }
  .alert ul { margin:6px 0 0; padding-left:18px; }
  .tabs { display:flex; gap:2px; border-bottom:2px solid var(--hair); padding:0 18px; flex-wrap:wrap; align-items:center; }
  .tab { border:none; background:none; padding:17px 18px; font-weight:800; font-size:15px; letter-spacing:-.01em; color:var(--muted); border-bottom:4px solid transparent; margin-bottom:-2px; cursor:pointer; transition:color .15s, border-color .15s; }
  .tab:hover { color:var(--ink); }
  .tab.on { color:var(--blue); border-bottom-color:var(--blue); }
  .tabs .xl { margin-left:auto; }
  .tabs .search { width:230px; min-height:40px; padding:6px 14px; margin-left:10px; }
  th.sortable { cursor:pointer; user-select:none; }
  th.sortable:hover { color:var(--ink); background:var(--blue-l); }
  .ar { display:inline-block; width:12px; color:var(--blue); font-size:10px; }
  .pager { display:flex; align-items:center; justify-content:space-between; gap:12px; flex-wrap:wrap; padding:14px 20px; border-top:1px solid var(--hair); font-size:13.5px; color:var(--muted); font-weight:600; }
  .pager .pg { display:flex; gap:5px; flex-wrap:wrap; }
  .pager button { min-width:36px; height:36px; border:1.5px solid rgba(12,26,58,.2); background:#fff; border-radius:9px; cursor:pointer; font-size:13.5px; font-weight:700; color:var(--ink-2); padding:0 10px; transition:background .12s, border-color .12s; }
  .pager button:hover:not(:disabled) { border-color:var(--blue); background:var(--blue-l); }
  .pager button.on { background:var(--blue); border-color:var(--blue); color:#fff; }
  .pager button:disabled { opacity:.4; cursor:default; }
  .map { height:230px; border-radius:10px; overflow:hidden; border:1px solid var(--hair); margin-top:4px; }
  .mapnone { font-size:13.5px; color:var(--muted); padding:6px 0; }
  .alt.click { cursor:pointer; } .alt.click:hover { color:var(--blue); }
  .altmap { display:none; margin:4px 0 10px; } .altmap.open { display:block; }
  .alt .lnk { color:var(--blue); font-size:12.5px; margin-left:6px; font-weight:750; }
  .dot { display:inline-block; width:11px; height:11px; border-radius:3px; margin-right:7px; vertical-align:baseline; }
  .pane { display:none; } .pane.on { display:block; }
  .tbl-wrap { overflow-x:auto; }
  table.tbl { width:100%; min-width:760px; border-collapse:collapse; font-size:14.5px; }
  .tbl td { white-space:nowrap; }
  .tbl th { background:var(--liner); color:var(--ink-2); font-size:12.5px; font-weight:800; text-align:center; padding:13px 12px; white-space:nowrap; border-bottom:1px solid var(--hair); position:sticky; top:0; letter-spacing:-.01em; }
  .tbl td { padding:14px 12px; text-align:center; border-bottom:1px solid rgba(12,26,58,.07); vertical-align:middle; }
  .tbl td.l, .tbl th.l { text-align:left; }
  .tbl tbody tr { transition:background .12s; }
  .tbl tbody tr.click { cursor:pointer; } .tbl tbody tr.click:hover { background:var(--blue-l); }
  .tbl th[title] { cursor:help; text-decoration:underline dotted rgba(12,26,58,.4); text-underline-offset:4px; }
  .rk { display:inline-block; min-width:38px; padding:3px 10px; border-radius:99px; background:var(--gold-l); color:var(--gold-d); font-weight:800; font-size:12.5px; }
  .pill { display:inline-block; padding:3px 11px; border-radius:99px; font-size:12.5px; font-weight:800; white-space:nowrap; letter-spacing:-.01em; }
  .pill.hit { background:var(--green-l); color:var(--green); } .pill.miss { background:var(--liner); color:var(--ink-2); }
  .pill.wait { background:var(--gold-l); color:var(--gold-d); } .pill.bad { background:var(--red-l); color:var(--red); }
  .score { font-weight:900; color:var(--blue); font-size:15.5px; letter-spacing:-.01em; }
  .empty { padding:40px; text-align:center; color:var(--muted); }

  /* 학생 상세: 인쇄된 배정 라벨 */
  .modal { position:fixed; inset:0; background:rgba(12,26,58,.62); display:none; align-items:center; justify-content:center; z-index:50; padding:16px; }
  .modal.open { display:flex; }
  .sheet { background:#fff; border-radius:18px; width:100%; max-width:580px; max-height:92vh; overflow:auto; padding:0 28px 28px; box-shadow:0 28px 70px rgba(0,0,0,.35); position:relative; }
  .sheet-top { display:flex; align-items:center; gap:12px; padding:24px 0 16px; margin-bottom:16px; border-bottom:3px dashed rgba(12,26,58,.2); }
  .sheet-top .bcs { height:26px; width:132px; color:var(--ink); }
  .sheet h3 { margin:30px 0 0; font-size:23px; font-weight:850; letter-spacing:-.01em; color:var(--ink); }
  .sheet .x { position:absolute; right:18px; top:16px; border:none; background:var(--liner); width:38px; height:38px; border-radius:50%; font-size:22px; line-height:1; cursor:pointer; transition:background .15s, transform .15s var(--ease); }
  .sheet .x:hover { background:var(--blue-l); transform:rotate(90deg); }
  .box { background:var(--liner); border-radius:12px; padding:16px 18px; margin-top:14px; }
  .box h4 { margin:0 0 10px; font-size:13.5px; font-weight:800; color:var(--ink-2); letter-spacing:-.01em; }
  .stack { display:flex; height:16px; border-radius:99px; overflow:hidden; background:#fff; margin:8px 0 12px; }
  .stack i { display:block; height:100%; }
  .eq { font-size:14px; display:grid; grid-template-columns:1fr auto; gap:5px 12px; }
  .eq b { text-align:right; font-variant-numeric:tabular-nums; }
  .alt { display:flex; justify-content:space-between; gap:12px; font-size:13.5px; padding:9px 0; border-bottom:1px solid var(--hair); }
  .alt:last-child { border-bottom:none; }

  /* 인증·알림 */
  .auth { position:fixed; inset:0; background:var(--blue-l); z-index:100; display:flex; align-items:center; justify-content:center; padding:16px; }
  .auth-card { background:#fff; color:var(--ink); width:100%; max-width:420px; padding:42px 34px 34px; border-radius:var(--r); box-shadow:var(--lift); text-align:center; }
  .auth-card h2 { color:var(--ink); font-weight:850; letter-spacing:-.01em; }
  .auth-card input { text-align:center; font-weight:800; margin:20px 0 8px; font-size:17px; min-height:52px; }
  .auth-card .btn { width:100%; }
  .toast { position:fixed; left:50%; bottom:28px; transform:translateX(-50%) translateY(20px); background:var(--ink); color:#fff; padding:14px 22px; border-radius:12px; font-size:14.5px; font-weight:650; opacity:0; transition:opacity .25s, transform .3s var(--ease); z-index:200; max-width:90vw; pointer-events:none; box-shadow:var(--lift); }
  .toast.show { opacity:1; transform:translateX(-50%); } .toast.err { background:var(--red); }

  @media (max-width:900px) {
    .hero { grid-template-columns:1fr; gap:24px; padding:26px 24px; }
    .kpis { grid-template-columns:repeat(2,1fr); }
    .g3 { grid-template-columns:1fr; }
    .brand span { display:none; }
  }
  @media (max-width:600px) {
    .wrap { padding:0 14px; } .card { padding:24px 18px 24px; }
    .stepper { gap:8px; }
    .stepper li { padding:10px 8px; font-size:12.5px; gap:6px; flex-direction:column; text-align:center; }
    .stepper li i { width:26px; height:26px; font-size:12.5px; }
    .g2 { grid-template-columns:1fr; }
    .bignum .n { font-size:36px; }
    .tabs .search { width:100%; margin:8px 0; order:9; }
    .tabs { padding:0 10px; }
    .tab { padding:14px 12px; font-size:14px; }
    .actions .btn { flex:1; }
    .modal { align-items:flex-end; padding:0; }
    .sheet { padding:0 18px 22px; border-radius:20px 20px 0 0; max-height:88vh; animation:sheetUp .32s var(--ease); }
    .sheet::before { content:''; display:block; width:44px; height:5px; border-radius:99px; background:rgba(12,26,58,.22); margin:10px auto 0; }
    .sheet h3 { margin-top:18px; }
  }
  @media (prefers-reduced-motion:reduce) { *, *::before, *::after { transition:none !important; animation:none !important; } .btn.busy::before { border-top-color:var(--ink); } }

  /* ===== 유리(글래스모피즘) 층 ===== */
  :root {
    --ink:#0B1B3A; --ink-2:#2A3F66; --muted:#4F6285;
    --blue:#2F5BEA; --blue-d:#1E3FA8; --blue-l:rgba(47,91,234,.13); --blue-m:#86A5F5;
    --liner:#E6EEFB; --hair:rgba(11,27,58,.12);
    --gold:#FF9A52; --gold-d:#8A3D00; --gold-l:rgba(255,154,82,.26);
    --green:#0A7A55; --green-l:rgba(16,185,129,.17); --red:#C2334D; --red-l:rgba(194,51,77,.13);
    --r:22px; --r-s:12px;
    --glass:rgba(255,255,255,.52); --glass-hi:rgba(255,255,255,.78); --glass-edge:rgba(255,255,255,.88);
    --lift:0 18px 50px rgba(30,63,168,.22), inset 0 1px 0 rgba(255,255,255,.9);
    --lift-s:0 10px 34px rgba(30,63,168,.13), inset 0 1px 0 rgba(255,255,255,.9), inset 0 -1px 0 rgba(255,255,255,.3);
    --bg:#E6EEFB; --teal:#2F5BEA; --teal-d:#1E3FA8; --amber:#FF9A52;
  }
  html { scrollbar-color:#8FA9E8 #E6EEFB; }
  body { background:#E6EEFB; position:relative; isolation:isolate; }
  body::before { content:''; position:fixed; inset:-18%; z-index:-2; pointer-events:none;
    background:
      radial-gradient(38% 34% at 14% 16%, #38D6F0 0%, transparent 70%),
      radial-gradient(34% 32% at 86% 14%, #6A93FF 0%, transparent 70%),
      radial-gradient(40% 36% at 80% 84%, #FFB48F 0%, transparent 72%),
      radial-gradient(34% 34% at 10% 86%, #7DF0D0 0%, transparent 70%),
      radial-gradient(30% 28% at 50% 48%, #D3E2FF 0%, transparent 75%),
      linear-gradient(180deg,#F0F5FF,#E1ECFB);
    animation:aur 38s ease-in-out infinite alternate; }
  body::after { content:''; position:fixed; inset:-10%; z-index:-1; pointer-events:none; opacity:.9;
    background:
      radial-gradient(14% 16% at 30% 30%, rgba(255,255,255,.75), transparent 70%),
      radial-gradient(12% 14% at 72% 58%, rgba(255,255,255,.55), transparent 70%),
      radial-gradient(16% 14% at 40% 82%, rgba(255,180,143,.55), transparent 70%),
      radial-gradient(12% 12% at 90% 40%, rgba(106,147,255,.55), transparent 70%);
    animation:aur2 46s ease-in-out infinite alternate; }
  @keyframes aur { from { transform:translate3d(-2%,-1%,0) scale(1); } to { transform:translate3d(3%,2%,0) scale(1.08); } }
  @keyframes aur2 { from { transform:translate3d(2%,2%,0); } to { transform:translate3d(-3%,-2%,0) scale(1.06); } }
  ::selection { background:#BCD0FF; color:var(--ink); }
  :focus-visible { outline:3px solid #2F5BEA; outline-offset:2px; box-shadow:0 0 0 5px rgba(255,255,255,.9); }

  /* 유리 면 */
  .card, .kpi, .chart, .mode, details.adv, .stepper li, .brand b, .pager button, .tabs .search {
    -webkit-backdrop-filter:blur(20px) saturate(170%); backdrop-filter:blur(20px) saturate(170%); }
  .card, .kpi, .chart { background:var(--glass); border:1px solid var(--glass-edge); box-shadow:var(--lift-s); }
  .card { background:var(--glass-hi); }
  #step3 .card { background:rgba(255,255,255,.78); }

  /* 상단 */
  .top { background:transparent; color:var(--ink); }
  .brand { color:var(--ink); }
  .brand b { background:linear-gradient(135deg,#3A6AF3,#2347C9); color:#fff; border:1px solid rgba(255,255,255,.55); box-shadow:0 8px 20px rgba(35,71,201,.3), inset 0 1px 0 rgba(255,255,255,.45); border-radius:14px; }
  .brand span { color:var(--ink-2); }
  .ver { background:rgba(255,255,255,.65); color:var(--blue-d); border:1px solid var(--glass-edge); }
  .stepper li { border:1px dashed rgba(11,27,58,.3); background:rgba(255,255,255,.3); color:var(--muted); border-radius:16px; }
  .stepper li.done { background:rgba(255,255,255,.55); border:1px solid var(--glass-edge); color:var(--ink); }
  .stepper li.done:hover { background:rgba(255,255,255,.85); }
  .stepper li.done i { background:var(--blue); border-color:var(--blue); }
  .stepper li.active { background:rgba(255,255,255,.88); border:1px solid #fff; box-shadow:var(--lift-s); color:var(--ink); }
  .stepper li.active i { background:linear-gradient(135deg,#3A6AF3,#2347C9); border-color:transparent; color:#fff; }

  /* 입력 */
  select, input[type=number], input[type=date], input[type=search], input[type=password] { background:rgba(255,255,255,.75); border:1px solid rgba(11,27,58,.22); border-radius:14px; }
  select:hover, input:hover { border-color:rgba(11,27,58,.5); }
  select:focus, input:focus { border-color:var(--blue); box-shadow:0 0 0 4px rgba(47,91,234,.18); }

  /* 버튼 */
  .btn { border:1px solid rgba(255,255,255,.55); border-radius:14px; box-shadow:0 10px 24px rgba(35,71,201,.3), inset 0 1px 0 rgba(255,255,255,.45); }
  .btn:hover { box-shadow:0 14px 30px rgba(35,71,201,.38), inset 0 1px 0 rgba(255,255,255,.5); }
  .btn:active { box-shadow:0 4px 10px rgba(35,71,201,.3), inset 0 1px 0 rgba(255,255,255,.4); }
  .btn-primary { background:linear-gradient(135deg,#3A6AF3 0%,#2347C9 100%); color:#fff; }
  .btn-ghost { background:rgba(255,255,255,.7); color:var(--ink); box-shadow:0 8px 20px rgba(30,63,168,.12), inset 0 1px 0 rgba(255,255,255,.9); -webkit-backdrop-filter:blur(14px); backdrop-filter:blur(14px); }
  .btn-ghost:hover { background:rgba(255,255,255,.95); }
  .btn.busy::before { border-color:rgba(255,255,255,.35); border-top-color:#fff; }
  .btn-sm { box-shadow:0 6px 14px rgba(35,71,201,.22), inset 0 1px 0 rgba(255,255,255,.45); }

  /* 선택 카드·상세 설정 */
  .mode { background:rgba(255,255,255,.55); border:1px solid var(--glass-edge); box-shadow:var(--lift-s); }
  .mode:hover { border-color:var(--blue); }
  .mode.on { border:1px solid var(--blue); background:rgba(255,255,255,.9); box-shadow:0 0 0 3px rgba(47,91,234,.2), var(--lift-s); }
  .mode.on::after { background:var(--blue); border-color:var(--blue); }
  .mode:focus-within { box-shadow:0 0 0 3px rgba(255,255,255,.95), 0 0 0 5px var(--blue); }
  details.adv { background:rgba(255,255,255,.5); border:1px solid var(--glass-edge); box-shadow:var(--lift-s); }
  details.adv > summary:hover { background:rgba(255,255,255,.55); }
  details.more { border:1px dashed rgba(11,27,58,.28); background:rgba(255,255,255,.3); }
  input[type=range]::-webkit-slider-runnable-track { background:linear-gradient(90deg,#2F5BEA 0 calc(14px + (100% - 28px) * var(--v)),var(--gold) calc(14px + (100% - 28px) * var(--v)) 100%); }
  input[type=range]::-moz-range-track { background:linear-gradient(90deg,#2F5BEA 0 calc(14px + (100% - 28px) * var(--v)),var(--gold) calc(14px + (100% - 28px) * var(--v)) 100%); }
  input[type=range]::-webkit-slider-thumb { background:var(--blue); border-color:#fff; box-shadow:0 4px 12px rgba(30,63,168,.4); }
  input[type=range]::-moz-range-thumb { background:var(--blue); border-color:#fff; box-shadow:0 4px 12px rgba(30,63,168,.4); }
  input[type=range]:focus-visible::-webkit-slider-thumb { box-shadow:0 0 0 4px rgba(255,255,255,.95), 0 0 0 7px var(--blue); }
  .wbar span:last-child { color:var(--gold-d); }

  /* 업로드 */
  .drop { background:rgba(255,255,255,.4); border:2px dashed rgba(11,27,58,.3); }
  .drop.over { background:rgba(255,255,255,.75); border-color:var(--blue); }
  .chip { background:rgba(255,255,255,.7); border:1px solid var(--glass-edge); color:var(--blue-d); }

  /* 결과: 청록 유리 요약 블록이 이 화면의 유일한 진한 면 */
  .hero { background:linear-gradient(135deg, rgba(30,63,168,.95) 0%, rgba(47,91,234,.85) 100%); border:1px solid rgba(255,255,255,.5); color:#fff;
    -webkit-backdrop-filter:blur(20px); backdrop-filter:blur(20px); box-shadow:0 20px 50px rgba(30,63,168,.3), inset 0 1px 0 rgba(255,255,255,.5); }
  .pos { color:#FFD27A; } .neg { color:#FFC2CC; }
  .pbar .fill.ours { background:linear-gradient(90deg,#FFD27A,#FF9A52); }
  .pbar .fill.base { background:rgba(255,255,255,.6); }
  .hrow .fl { background:linear-gradient(90deg,#2F5BEA,#22C7F0); }
  .hrow .fl.full { background:var(--ink); }
  .hrow .tr { background:rgba(11,27,58,.09); }
  .vb .bar { background:linear-gradient(180deg,#22C7F0,#2F5BEA); }
  .rk { background:var(--gold-l); color:var(--gold-d); }
  .alert { background:rgba(255,154,82,.2); border:1px solid rgba(255,154,82,.55); color:var(--gold-d); border-radius:16px; }
  .tabs { border-bottom:1px solid rgba(11,27,58,.14); }
  .tab.on { color:var(--blue-d); border-bottom-color:var(--blue); }
  .tbl th { background:rgba(255,255,255,.7); backdrop-filter:blur(10px); -webkit-backdrop-filter:blur(10px); }
  .tbl tbody tr.click:hover { background:rgba(255,255,255,.7); }
  .tbl td { border-bottom-color:rgba(11,27,58,.08); }
  .pager button { background:rgba(255,255,255,.7); border:1px solid var(--glass-edge); }
  .pager button.on { background:var(--blue); border-color:var(--blue); color:#fff; }
  .tabs .search { background:rgba(255,255,255,.75); }

  /* 학생 상세·인증·알림 */
  .modal { background:rgba(11,27,58,.38); -webkit-backdrop-filter:blur(8px); backdrop-filter:blur(8px); }
  .sheet { background:rgba(255,255,255,.82); border:1px solid #fff; -webkit-backdrop-filter:blur(28px) saturate(170%); backdrop-filter:blur(28px) saturate(170%); box-shadow:0 30px 80px rgba(11,27,58,.4), inset 0 1px 0 #fff; border-radius:26px; }
  .sheet .x { background:rgba(11,27,58,.08); }
  .box { background:rgba(255,255,255,.6); border:1px solid rgba(255,255,255,.9); }
  .stack { background:rgba(11,27,58,.09); }
  .auth { background:#E6EEFB; }
  .auth::before { content:''; position:absolute; inset:0; z-index:-1;
    background:radial-gradient(38% 34% at 14% 16%, #38D6F0, transparent 70%), radial-gradient(34% 32% at 86% 14%, #6A93FF, transparent 70%), radial-gradient(40% 36% at 80% 84%, #FFB48F, transparent 72%), radial-gradient(34% 34% at 10% 86%, #7DF0D0, transparent 70%); }
  .auth-card { background:rgba(255,255,255,.7); border:1px solid #fff; -webkit-backdrop-filter:blur(26px) saturate(170%); backdrop-filter:blur(26px) saturate(170%); box-shadow:var(--lift); border-radius:26px; }
  .toast { background:rgba(11,27,58,.92); -webkit-backdrop-filter:blur(12px); backdrop-filter:blur(12px); border:1px solid rgba(255,255,255,.2); border-radius:16px; }
  .toast.err { background:rgba(194,51,77,.95); }

  @media (max-width:600px) {
    .modal { align-items:flex-end; padding:0; }
    .sheet { border-radius:26px 26px 0 0; }
    .card, .kpi, .chart, .mode, details.adv, .stepper li { -webkit-backdrop-filter:blur(14px) saturate(150%); backdrop-filter:blur(14px) saturate(150%); }
  }
  @media (prefers-reduced-motion:reduce) { body::before, body::after { animation:none !important; } }

  /* ===== 이동·전환 ===== */
  main.wrap { overflow-x:clip; }
  .stepper li { transition:background .25s, border-color .25s, color .25s, transform .25s var(--ease); }
  .stepper li i { transition:background .3s, border-color .3s, color .3s, transform .35s var(--ease); }
  .stepper li.active i { transform:scale(1.08); }
  .steprog { height:3px; background:rgba(11,27,58,.12); border-radius:3px; overflow:hidden; margin:0 0 -1px; }
  .steprog i { display:block; height:100%; width:33.3%; background:var(--blue); border-radius:3px; transition:width .6s var(--ease); }
  .mode { transition:border-color .2s, background .2s, transform .2s var(--ease), box-shadow .2s; }
  .mode.on::after { animation:tickBg .3s var(--ease); }
  .mode.on::before { animation:tick .35s var(--ease); }
  @keyframes tick { from { transform:rotate(-45deg) scale(0); } to { transform:rotate(-45deg) scale(1); } }
  @keyframes tickBg { from { transform:scale(.6); } to { transform:scale(1); } }
  .tabs { position:relative; }
  .tab { border-bottom:0; margin-bottom:0; padding-bottom:16px; }
  .tab.on { border-bottom-color:transparent; }
  .tab-ind { position:absolute; left:0; bottom:-1px; height:3px; width:0; border-radius:3px; background:var(--blue); transition:transform .45s var(--ease), width .45s var(--ease); pointer-events:none; }
  .pane.on { animation:paneIn .35s var(--ease); }
  @keyframes paneIn { from { opacity:0; transform:translateY(8px); } to { opacity:1; transform:none; } }
  .step.in-fwd { animation:inFwd .45s var(--ease); }
  .step.in-back { animation:inBack .45s var(--ease); }
  @keyframes inFwd { from { opacity:0; transform:translateX(36px); } to { opacity:1; transform:none; } }
  @keyframes inBack { from { opacity:0; transform:translateX(-36px); } to { opacity:1; transform:none; } }
  .modal.open { animation:fadeIn .22s ease; }
  .modal.open .sheet { animation:popIn .34s var(--ease); }
  @keyframes fadeIn { from { opacity:0; } to { opacity:1; } }
  @keyframes popIn { from { opacity:0; transform:translateY(18px) scale(.97); } to { opacity:1; transform:none; } }
  .brand b { view-transition-name:brandmark; }
  @view-transition { navigation:auto; }
  ::view-transition-old(root) { animation:vtOut .22s ease both; }
  ::view-transition-new(root) { animation:vtIn .42s var(--ease) both; }
  ::view-transition-group(brandmark) { animation-duration:.45s; animation-timing-function:var(--ease); }
  @keyframes vtOut { to { opacity:0; transform:translateY(-8px); } }
  @keyframes vtIn { from { opacity:0; transform:translateY(14px); } }
  @media (max-width:600px) { .modal.open .sheet { animation:sheetUp .36s var(--ease); } }
  @media (prefers-reduced-motion:reduce) { .step.in-fwd, .step.in-back, .pane.on, .modal.open, .modal.open .sheet { animation:none !important; } ::view-transition-old(root), ::view-transition-new(root) { animation:none !important; } }
</style>
</head>
<body>

<div id="auth" class="auth">
  <div class="auth-card">
    <h2 style="margin:0 0 6px;font-size:24px">보안 서버 인증</h2>
    <div class="sub" style="font-size:14px">AI기반 간호실습 배치 의사결정 프로그램</div>
    <input type="password" id="authPw" placeholder="접속 암호 입력" autocomplete="current-password" onkeyup="if(event.key==='Enter')verifyPassword()">
    <div id="authErr" style="color:var(--red);font-weight:700;font-size:13.5px;min-height:22px;margin-bottom:12px"></div>
    <button class="btn btn-primary" onclick="verifyPassword()">시스템 접속하기</button>
  </div>
</div>

<header class="top"><div class="wrap top-in">
  <a href="/" style="text-decoration:none;color:inherit" title="소개 페이지로"><div class="brand"><b>로켓단</b><span>AI기반 간호실습 배치 의사결정 프로그램</span></div></a>
  <span class="ver">__APP_VERSION__</span>
</div>
<div class="wrap">  <ol class="stepper" id="stepper">
    <li data-s="1" class="active"><i>1</i>조건 설정</li>
    <li data-s="2"><i>2</i>명단 업로드</li>
    <li data-s="3"><i>3</i>결과 확인</li>
  </ol><div class="steprog" aria-hidden="true"><i id="stepFill"></i></div></div></header>

<main class="wrap">

  <!-- 1단계 -->
  <section id="step1" class="card step">
    <h2>배정 조건</h2>
    <p class="lead">배정 방식과 교과목을 선택하면 해당 교과목의 실습 병원 목록을 불러옵니다.</p>

    <div class="grid g2" style="margin-bottom:18px">
      <label class="mode on" id="modeG"><input type="radio" name="mode" value="global" checked onchange="onModeChange()">
        <b>전체 병원 동시 배정 <span class="pill hit" style="margin-left:4px">추천</span></b>
        <span>모든 학생과 병원을 한 번에 계산합니다. 학생은 병원 한 곳에만 배정됩니다.</span></label>
      <label class="mode" id="modeS"><input type="radio" name="mode" value="single" onchange="onModeChange()">
        <b>단일 병원 순위</b>
        <span>병원 한 곳을 선택해 지원자의 종합점수 순위를 매기고 정원까지 배정합니다.</span></label>
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
    <input type="checkbox" id="useTransit" checked hidden>

    <div class="hint" id="transitHint"></div>

    <details class="adv">
      <summary>상세 설정 — 배정 기준 · 실습 기간 · 병원별 정원 · 추가 조건</summary>
      <div class="adv-body">
        <h4>배정 기준</h4>
        <div class="wbar"><span id="wDistLbl">통학 거리 70%</span><span id="wPrefLbl">희망 병원 30%</span></div>
        <input type="range" id="wDist" min="0" max="100" step="5" value="70" style="--v:.7" oninput="syncW()" aria-label="통학 거리와 희망 병원의 비중">
        <input type="hidden" id="wPref" value="30">
        <div class="hint">파란 쪽은 통학 거리, 노란 쪽은 희망 병원의 비중입니다. 손잡이를 밀어 조절하세요.</div>

        <h4>실습 기간</h4>
        <div class="grid g2">
          <div><span class="lbl">실습 시작일</span><input type="date" id="startDate" onchange="onStartChange()"></div>
          <div><span class="lbl">실습 종료일 (2주 안 · 금요일 자동)</span><input type="date" id="endDate" onchange="checkDate(this);updateDayCount()"></div>
        </div><div class="hint" id="dateErr" style="color:var(--bad);font-weight:700"></div>
        <div class="hint" id="dayCount" style="font-weight:700"></div>

        <h4>병원별 정원</h4>
        <div class="hint" style="margin:0 0 10px">병원별 정원을 입력합니다. 비워 두면 아래 기본 정원을 적용합니다.</div>
        <div id="capTable" class="caps"></div>
        <div style="display:flex;gap:8px;margin-top:12px;flex-wrap:wrap;align-items:center">
          <button type="button" class="btn btn-ghost btn-sm" id="capActualBtn" onclick="fillCaps('actual')" style="display:none">2025-1 실제 인원으로 채우기</button>
          <button type="button" class="btn btn-ghost btn-sm" onclick="fillCaps('clear')">모두 비우기</button>
          <span class="sub" id="capSum"></span>
        </div>
        <div style="display:flex;gap:10px;align-items:center;margin-top:12px;font-size:14px">
          <span class="lbl" style="margin:0">기본 정원</span><input type="number" id="capHosp" value="40" min="1" style="width:84px;text-align:right">
          <span class="sub">명 (칸을 비운 병원에 적용)</span>
        </div>
        <input type="hidden" id="capPeriod" value="50">

        <details class="more">
          <summary>추가 조건 (성별 · GPA · 출생연도 · 과거 실습 제외)</summary>
          <div class="more-body">
            <div class="grid g3">
              <div><span class="lbl">성별 조건</span><select id="gender"><option>무관</option><option>남성만</option><option>여성만</option></select></div>
              <div><span class="lbl">최소 GPA</span><input type="number" step="0.1" id="minGpa" placeholder="예: 3.5"></div>
              <div><span class="lbl">출생연도 이후</span><input type="number" id="birthYear" placeholder="예: 2003"></div>
            </div>
            <label style="display:flex;gap:8px;align-items:center;font-size:14px;font-weight:600;cursor:pointer;margin-top:14px">
              <input type="checkbox" id="exPast" style="width:auto"> 과거 실습 기관 중복 배정 자동 제외</label>
          </div>
        </details>
      </div>
    </details>

    <div class="actions"><span></span><button class="btn btn-primary" onclick="goStep(2)">다음: 명단 업로드</button></div>
  </section>

  <!-- 2단계 -->
  <section id="step2" class="card step" style="display:none">
    <h2>학생 명단 업로드</h2>
    <p class="lead">파일은 메모리에서만 처리하며 서버에 저장하지 않습니다.</p>
    <div class="chips" id="condChips"></div>

    <div class="drop" id="drop">
      <div class="big">엑셀(.xlsx) 또는 CSV 파일을 이곳에 놓거나 선택하세요</div>
      <div class="sub" style="margin-bottom:14px">필수 열: 학번 · 이름 · 성별 · GPA · 출생연도 (나머지 열은 양식 샘플 참고)</div>
      <label class="btn btn-ghost btn-sm" style="cursor:pointer">파일 선택<input type="file" id="file" accept=".csv,.xlsx" style="display:none"></label>
      <div class="file-ok" id="fileOk"></div>
    </div>
    <div class="or">또는</div>
    <div class="tools">
      <button class="btn btn-ghost btn-sm" onclick="loadSample()">샘플 명단 불러오기</button>
      <a class="btn btn-ghost btn-sm" href="/api/v1/template">양식 샘플 다운로드</a>
    </div>

    <div class="actions">
      <button class="btn btn-ghost" onclick="goStep(1)">이전</button>
      <button class="btn btn-primary" id="runBtn" onclick="run()">배정 실행</button>
    </div>
  </section>

  <!-- 3단계 -->
  <section id="step3" class="step" style="display:none">
    <div class="res-top">
      <div class="meta" id="resMeta"></div>
      <div style="display:flex;gap:8px;flex-wrap:wrap">
        <button class="btn btn-ghost btn-sm" onclick="goStep(1)">조건 바꿔 다시 배정</button>
      </div>
    </div>

    <div id="warn"></div>
    <div id="hero"></div>
    <div class="kpis" id="kpis"></div>
    <div class="charts" id="charts"></div>

    <div class="card" style="padding:0;overflow:hidden">
      <div class="tabs"><span class="tab-ind" id="tabInd"></span>
        <button class="tab on" data-p="pA" onclick="showPane('pA')">배정 결과</button>
        <button class="tab" data-p="pH" onclick="showPane('pH')">병원별 현황</button>
        <button class="tab" data-p="pU" onclick="showPane('pU')">미배정·사유</button>
        <button class="btn btn-primary btn-sm xl" onclick="exportXlsx()">엑셀 다운로드</button>
        <input type="search" class="search" id="q" placeholder="이름·학번·병원 검색" oninput="PG=1;renderAssigned()">
      </div>
      <div id="pA" class="pane on tbl-wrap">
        <table class="tbl"><thead><tr>
          <th class="sortable" onclick="sortBy('rank')">순위<span class="ar" id="ar_rank"></span></th><th class="l sortable" onclick="sortBy('hospital')">병원<span class="ar" id="ar_hospital"></span></th><th class="sortable" onclick="sortBy('student_id')">학번<span class="ar" id="ar_student_id"></span></th><th class="sortable" onclick="sortBy('name')">이름<span class="ar" id="ar_name"></span></th><th class="sortable" onclick="sortBy('gpa')">GPA<span class="ar" id="ar_gpa"></span></th>
          <th class="sortable" onclick="sortBy('pref')" title="학생이 적은 1~3지망과 배정 병원의 일치 여부">지망 반영<span class="ar" id="ar_pref"></span></th>
          <th class="sortable" onclick="sortBy('time')" title="TMAP 대중교통 소요시간. '약'은 지도 좌표 기반 추정, '추정'은 주소 기반 규칙 추정">통학시간<span class="ar" id="ar_time"></span></th>
          <th class="sortable" onclick="sortBy('score')" title="통학 점수 × 거리 가중치 + 지망 점수 × 희망도 가중치">종합점수<span class="ar" id="ar_score"></span></th>
          <th class="sortable" onclick="sortBy('sat')" title="통학 부담(소요시간·환승·도보) 기반 규칙형 지수. 학습 모델이 아닙니다">만족도 지수<span class="ar" id="ar_sat"></span></th>
        </tr></thead><tbody id="bodyA"></tbody></table>
        <div class="pager" id="pager"></div>
        <div class="sub" style="padding:0 16px 12px">행을 누르면 점수 산출 내역을 볼 수 있습니다. 통학시간의 '약'은 지도 기반 예상 소요시간, '추정'은 규칙 기반 추정입니다. 열 제목을 누르면 정렬됩니다.</div>
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
function bizDays(a, b) {   // 시작일~종료일 사이 평일(공휴일 제외) 일수, 양끝 포함
  const d = pd_(a), end = pd_(b); let n = 0;
  while (d <= end) { if (!dateProblem(fmt(d))) n++; d.setDate(d.getDate() + 1); }
  return n;
}
function updateDayCount() {
  const a = $('startDate').value, b = $('endDate').value, el = $('dayCount');
  if (!a || !b) { el.textContent = ''; return; }
  const n = bizDays(a, b);
  if (n < 8) { el.style.color = 'var(--bad)'; el.textContent = `실습일수 ${n}일: 8일 미만이면 실습 이수가 되지 않습니다.`; }
  else if (n > 10) { el.style.color = 'var(--bad)'; el.textContent = `실습일수 ${n}일: 최대 10일을 넘습니다.`; }
  else { el.style.color = 'var(--green)'; el.textContent = `실습일수 ${n}일 (8~10일 충족)`; }
}
function defaultEnd(s) {   // 실습은 2주 안에 끝나고 보통 금요일: 시작일이 속한 주의 '다음 주 금요일'
  const d = pd_(s); d.setDate(d.getDate() - ((d.getDay() + 6) % 7) + 11);
  let e = fmt(d); while (dateProblem(e)) { d.setDate(d.getDate() - 1); e = fmt(d); }   // 그날이 휴일이면 그 전 평일
  return e;
}
function addBiz(s, n) { const d = pd_(s); let c = 0; while (c < n) { d.setDate(d.getDate() + 1); if (!dateProblem(fmt(d))) c++; } return fmt(d); }
function initDates() {
  const d = new Date();   // 시작일 기본값 = 오늘 (주말·공휴일이면 다음 평일)
  let s = fmt(d); while (dateProblem(s)) { d.setDate(d.getDate() + 1); s = fmt(d); }
  if (bizDays(s, defaultEnd(s)) < 8) {   // 오늘 시작하면 8일을 못 채우는 요일(목·금 등)이면 다음 월요일부터
    d.setDate(d.getDate() + ((8 - d.getDay()) % 7 || 7)); s = fmt(d); while (dateProblem(s)) { d.setDate(d.getDate() + 1); s = fmt(d); }
  }
  $('startDate').value = s; $('endDate').value = defaultEnd(s); updateDayCount();
}
function checkDate(inp) { const m = inp.value ? dateProblem(inp.value) : null; $('dateErr').textContent = m ? m : ''; if (m) inp.value = ''; return !m; }
function onStartChange() { const s = $('startDate'); if (checkDate(s) && s.value) $('endDate').value = defaultEnd(s.value); updateDayCount(); }
function syncW() {
  const a = $('wDist'), b = $('wPref'), d = Math.max(0, Math.min(100, parseInt(a.value) || 0));
  a.value = d; b.value = 100 - d; a.style.setProperty('--v', d / 100);
  $('wDistLbl').textContent = `통학 거리 ${d}%`; $('wPrefLbl').textContent = `희망 병원 ${100 - d}%`;
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
let ACTUAL_CAPS = {};
function buildCapTable(hs) {
  const box = $('capTable'); box.innerHTML = '';
  hs.forEach(h => {
    const w = document.createElement('div'); w.className = 'caprow';
    w.innerHTML = `<span>${esc(h.replace('(방중)', ''))}</span><input type="number" min="1" class="capIn" data-h="${esc(h)}" placeholder="기본" oninput="capSum()">`;
    box.appendChild(w);
  });
  $('capActualBtn').style.display = Object.keys(ACTUAL_CAPS).length ? '' : 'none';
  fillCaps(Object.keys(ACTUAL_CAPS).length ? 'actual' : 'clear');
}
function fillCaps(kind) {
  document.querySelectorAll('.capIn').forEach(i => { i.value = kind === 'actual' ? (ACTUAL_CAPS[i.dataset.h] ?? '') : ''; });
  capSum();
}
function capSum() {
  let t = 0, n = 0;
  document.querySelectorAll('.capIn').forEach(i => { if (i.value !== '') { t += parseInt(i.value) || 0; n++; } });
  $('capSum').textContent = n ? `입력된 ${n}곳 합계 ${t}명` : '';
}
function capsJson() {
  const o = {};
  document.querySelectorAll('.capIn').forEach(i => { if (i.value !== '' && parseInt(i.value) >= 1) o[i.dataset.h] = parseInt(i.value); });
  return Object.keys(o).length ? JSON.stringify(o) : '';
}
async function updateHospitals() {
  const q = new URLSearchParams({grade: $('grade').value, semester: $('semester').value, subject: $('subject').value});
  const box = $('hospital'); box.innerHTML = '';
  try {
    const d = await (await fetch('/api/v1/hospitals?' + q)).json();
    (d.hospitals || []).forEach(h => box.add(new Option(h, h)));
    $('hospCount').textContent = `이 교과목의 실습 병원 ${(d.hospitals || []).length}곳`;
    ACTUAL_CAPS = d.actual_caps || {};
    buildCapTable(d.hospitals || []);
  } catch (e) { $('hospCount').textContent = '병원 목록을 불러오지 못했습니다.'; }
}
window.addEventListener('load', async () => {
  initDates();
  try { CURRICULUM = await (await fetch('/api/v1/curriculum')).json(); } catch (e) {}
  try { CONFIG = await (await fetch('/api/v1/config')).json(); } catch (e) { CONFIG = {}; }
  loadNaverMap();
  if (CONFIG.auth_required && !TOKEN) $('auth').style.display = 'flex';
  if (CONFIG.transit_available) {
    const m = CONFIG.hospitals_missing_address;
    $('transitHint').textContent = '통학시간: ' + (CONFIG.transit_mode === 'tmap' ? 'TMAP 대중교통 소요시간 사용(' + (CONFIG.commute_label || '평일 오전 6시 출발') + ' 기준). ' : '지도 기반 예상 소요시간. ')
      + (m ? `병원 주소가 없는 ${m}곳은 규칙 기반 추정치를 씁니다. ` : '') + '첫 실행은 시간이 걸리며, 이후에는 저장된 값을 사용합니다.';
  } else {
    $('transitHint').textContent = '지도 키가 설정되지 않아 통학시간은 규칙 기반 추정치로 계산합니다.';
  }
  syncW();
  updateSubjects(); onModeChange();
});

/* ===== 단계 이동 ===== */
function goStep(n) {
  if (n === 3 && !DATA) return;
  const _prev = STEP; STEP = n; $('toast').className = 'toast';
  document.querySelectorAll('.step').forEach(s => s.style.display = s.id === 'step' + n ? 'block' : 'none');
  { const el = $('step' + n); el.classList.remove('in-fwd', 'in-back'); void el.offsetWidth; el.classList.add(n >= _prev ? 'in-fwd' : 'in-back'); }
  const sf = $('stepFill'); if (sf) sf.style.width = (n / 3 * 100) + '%';
  document.querySelectorAll('#stepper li').forEach(li => {
    const s = +li.dataset.s;
    li.className = s === n ? 'active' : ((s < n || (s === 3 && DATA) || (s === 2)) ? 'done' : '');
  });
  if (n === 2) renderChips();
  if (n === 3) requestAnimationFrame(posTabInd);
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
  ok.style.display = 'block'; ok.textContent = `${f.name} (${Math.max(1, Math.round(f.size / 1024))}KB)`;
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
    toast('가상 학생 60명을 불러왔습니다.');
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
  f.append('use_transit_api', true);
  f.append('hospital_caps', capsJson());
  f.append('distance_weight', (parseInt($('wDist').value) || 0) / 100); f.append('preference_weight', (parseInt($('wPref').value) || 0) / 100);
  if ($('minGpa').value !== '') f.append('min_gpa', $('minGpa').value);
  if ($('birthYear').value !== '') f.append('birth_year_after', $('birthYear').value);
  f.append('file', FILE);
  const b = $('runBtn'); b.disabled = true; b.classList.add('busy'); b.textContent = '배정 계산 중…';
  try {
    const r = await fetch('/api/v1/assign-file', {method: 'POST', body: f, headers: {'X-Auth-Token': TOKEN}});
    let d = {}; try { d = await r.json(); } catch (e) {}
    if (r.status === 401) { authExpired(); return; }
    if (!r.ok) { toast(d.detail || '요청 처리 중 오류가 발생했습니다.', true); return; }
    DATA = d; DATA._ctx = `${$('grade').value} ${$('semester').value}`; render(d); goStep(3);
  } catch (e) { toast('서버에 연결할 수 없습니다.', true); }
  finally { b.disabled = false; b.classList.remove('busy'); b.textContent = '배정 실행'; }
}

/* ===== 결과 렌더 ===== */
function showPane(id) {
  document.querySelectorAll('.pane').forEach(p => p.classList.toggle('on', p.id === id));
  document.querySelectorAll('.tab').forEach(t => t.classList.toggle('on', t.dataset.p === id));
  $('q').style.display = id === 'pA' ? '' : 'none';
  posTabInd();
}
function posTabInd() {
  const on = document.querySelector('.tab.on'), ind = $('tabInd');
  if (!on || !ind || !on.offsetWidth) return;
  ind.style.width = on.offsetWidth + 'px'; ind.style.transform = 'translateX(' + on.offsetLeft + 'px)';
}
window.addEventListener('resize', posTabInd);
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
      <div class="cmp-note">비교 기준: ${esc(c.baseline_label)}<br>${d.transit && d.transit.real_ratio > 0 ? `통학시간: ${d.transit.mode === 'tmap' ? 'TMAP 대중교통 실측' : '지도 기반 예상 소요시간'} ${num(d.transit.real_ratio)}% (나머지는 규칙 기반 추정)` : '통학시간은 동 단위 규칙 기반 추정치입니다.'}</div>
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
  const pc = [['1지망', '#2F5BEA'], ['2지망', '#22C7F0'], ['3지망', '#A9C4FA'], ['비지망', '#D5DEEF']];
  const pv = pc.map(([l]) => rs.filter(r => l === '비지망' ? r.preference_status.startsWith('비') : r.preference_status.startsWith(l)).length);
  const tot = pv.reduce((a, b) => a + b, 0); let acc = 0;
  const stops = tot ? pc.map(([, col], i) => { const a = acc / tot * 100; acc += pv[i]; return `${col} ${a}% ${acc / tot * 100}%`; }).join(',') : '#E6EEFB 0 100%';
  html += `<div class="chart"><h3>지망 충족 현황</h3><div class="donut-wrap"><div class="donut" style="background:conic-gradient(${stops})"><div class="donut-hole"><b>${num(d.pref_match_rate)}%</b><span>지망 충족</span></div></div>
    <div class="legend">${pc.map(([l, col], i) => `<div><i style="background:${col}"></i>${l} <b>${pv[i]}명</b></div>`).join('')}</div></div></div>`;
  $('charts').innerHTML = html;
}
/* 개인정보 가림: 이름은 첫·끝 글자만, 학번은 가운데 3자리를 가린다 (화면 표시에만 적용) */
function mName(n) {
  const a = Array.from(String(n ?? '')); if (a.length < 2) return a.join('');
  return a.length === 2 ? a[0] + '○' : a[0] + '○'.repeat(a.length - 2) + a[a.length - 1];
}
function mId(v) {
  const a = Array.from(String(v ?? '')), n = a.length; if (n < 3) return a.join('');
  const k = n >= 6 ? 3 : 1, st = Math.ceil((n - k) / 2);
  return a.map((c, i) => (i >= st && i < st + k) ? '○' : c).join('');
}
let SORT = {k: null, d: 1}, PG = 1;
const PAGE = 20;
const SORTVAL = {rank: r => r.rank, hospital: r => r.hospital, student_id: r => r.student_id, name: r => r.name, gpa: r => r.gpa,
  pref: r => r.pref_score, time: r => r.travel_time_minutes, score: r => r.composite_score, sat: r => r.satisfaction_index};
function sortBy(k) {
  if (SORT.k !== k) SORT = {k, d: 1};
  else if (SORT.d === 1) SORT.d = -1;
  else SORT = {k: null, d: 1};
  PG = 1; renderAssigned();
}
function goPage(p) { PG = p; renderAssigned(); const t = $('pA'); if (t) t.scrollIntoView({block: 'nearest'}); }
function renderAssigned() {
  const q = $('q').value.trim().toLowerCase();
  let rows = DATA.results.map((r, i) => [r, i]).filter(([r]) => !q || [r.name, r.student_id, r.hospital].some(v => String(v).toLowerCase().includes(q)));
  if (SORT.k) {
    const f = SORTVAL[SORT.k];
    rows.sort(([a], [b]) => {
      const x = f(a), y = f(b);
      if (x === null || x === undefined) return 1;
      if (y === null || y === undefined) return -1;
      const c = (typeof x === 'number' && typeof y === 'number') ? x - y : String(x).localeCompare(String(y), 'ko', {numeric: true});
      return c * SORT.d;
    });
  }
  Object.keys(SORTVAL).forEach(k => { const e = $('ar_' + k); if (e) e.textContent = SORT.k === k ? (SORT.d === 1 ? '▲' : '▼') : ''; });
  const pages = Math.max(1, Math.ceil(rows.length / PAGE));
  if (PG > pages) PG = pages;
  const from = (PG - 1) * PAGE, view = rows.slice(from, from + PAGE);
  $('bodyA').innerHTML = view.map(([r, i]) => `<tr class="click" onclick="openDetail(${i})">
    <td><span class="rk">${esc(r.rank)}위</span></td><td class="l">${esc(r.hospital)}</td><td>${esc(mId(r.student_id))}</td><td><b>${esc(mName(r.name))}</b></td><td>${esc(r.gpa)}</td>
    <td><span class="pill ${r.preference_status.includes('일치') ? 'hit' : 'miss'}">${esc(r.preference_status)}</span></td>
    <td><b>${r.transit_source === 'TMAP' ? '' : '약 '}${esc(r.travel_time_minutes)}분</b>${r.transit_source === '추정' ? '<span class="sub"> 추정</span>' : ''}</td><td><span class="score">${esc(r.composite_score)}</span></td><td>${esc(r.satisfaction_index)}</td></tr>`).join('')
    || `<tr><td colspan="9" class="empty">${q ? '검색 결과가 없습니다.' : '배정된 학생이 없습니다.'}</td></tr>`;
  const pg = [];
  for (let p = 1; p <= pages; p++) if (p === 1 || p === pages || Math.abs(p - PG) <= 2) pg.push(p);
  let btns = '', last = 0;
  pg.forEach(p => { if (p - last > 1) btns += '<span style="padding:0 4px;align-self:center">…</span>'; btns += `<button class="${p === PG ? 'on' : ''}" onclick="goPage(${p})">${p}</button>`; last = p; });
  $('pager').innerHTML = rows.length
    ? `<span>총 ${rows.length}명 중 ${from + 1}–${from + view.length}번째</span><div class="pg"><button onclick="goPage(${PG - 1})" ${PG <= 1 ? 'disabled' : ''}>이전</button>${btns}<button onclick="goPage(${PG + 1})" ${PG >= pages ? 'disabled' : ''}>다음</button></div>`
    : '';
}
function render(d) {
  $('resMeta').innerHTML = `<b>[${esc(d._ctx)}]</b> ${esc(d.period.start)} ~ ${esc(d.period.end)} · 대상 <b>${esc(d.target_hospital)}</b><br><span class="sub">${esc(d.optimization_method)}</span>`;
  $('warn').innerHTML = d.warnings.length ? `<div class="alert"><b>데이터 확인 필요 (${d.warnings.length}건)</b><ul>${d.warnings.map(w => `<li>${esc(w.replace(/학번 (\S+?)(의| 중복)/, (m, id, t) => '학번 ' + mId(id) + t))}</li>`).join('')}</ul></div>` : '';
  renderHero(d); renderKpis(d); renderCharts(d); renderAssigned();
  $('bodyH').innerHTML = d.hospital_summary.map(h => `<tr><td class="l">${esc(h.hospital)}</td><td><b>${h.assigned}</b> / ${h.capacity}</td><td>${h.avg_travel === null ? '-' : esc(h.avg_travel) + '분'}</td><td>${h.pref_matched}명</td></tr>`).join('');
  $('bodyU').innerHTML = d.unassigned.map(u => `<tr><td><span class="pill ${u.status === '대기' ? 'wait' : 'bad'}">${esc(u.status)}</span></td><td>${esc(mId(u.student_id))}</td><td><b>${esc(mName(u.name))}</b></td><td>${esc(u.gpa)}</td><td class="l">${esc(u.reason)}</td></tr>`).join('')
    || '<tr><td colspan="5" class="empty">미배정 학생이 없습니다.</td></tr>';
  $('q').value = ''; SORT = {k: null, d: 1}; PG = 1; renderAssigned(); showPane('pA');
}

/* ===== 학생 상세 ===== */
let MAPS = [];
function killMaps() { MAPS.forEach(m => { try { m.remove(); } catch (e) {} }); MAPS = []; }
function kmBetween(a, b) {
  const R = 6371, rad = x => x * Math.PI / 180, dLa = rad(b[1] - a[1]), dLo = rad(b[0] - a[0]);
  const h = Math.sin(dLa / 2) ** 2 + Math.cos(rad(a[1])) * Math.cos(rad(b[1])) * Math.sin(dLo / 2) ** 2;
  return 2 * R * Math.asin(Math.sqrt(h));
}
let NAVER_FAIL = false;
window.navermap_authFailure = function () { NAVER_FAIL = true; };
function loadNaverMap() {
  if (!CONFIG.map_key_id || window.naver) return;
  const sc = document.createElement('script');
  sc.src = 'https://oapi.map.naver.com/openapi/v3/maps.js?ncpKeyId=' + encodeURIComponent(CONFIG.map_key_id);
  document.head.appendChild(sc);
}
function pin(color, label) {
  return `<div style="transform:translate(-50%,-100%);text-align:center;white-space:nowrap"><div style="display:inline-block;background:#fff;border:2px solid ${color};border-radius:8px;padding:3px 8px;font:700 12px sans-serif;color:#0f172a;box-shadow:0 2px 6px rgba(0,0,0,.2)">${esc(label)}</div><div style="width:0;height:0;margin:0 auto;border-left:6px solid transparent;border-right:6px solid transparent;border-top:8px solid ${color}"></div></div>`;
}
function drawNaver(el, s, h, hname) {
  const N = naver.maps, a = new N.LatLng(s[1], s[0]), b = new N.LatLng(h[1], h[0]);
  const m = new N.Map(el, {center: a, zoom: 13, scaleControl: false, mapDataControl: false});
  new N.Polyline({map: m, path: [a, b], strokeColor: '#2F5BEA', strokeWeight: 4, strokeOpacity: 0.9, strokeStyle: 'shortdash'});
  new N.Marker({position: a, map: m, icon: {content: pin('#FF9A52', '거주지'), anchor: new N.Point(0, 0)}});
  new N.Marker({position: b, map: m, icon: {content: pin('#2F5BEA', hname), anchor: new N.Point(0, 0)}});
  m.fitBounds(new N.LatLngBounds(a, b), {top: 50, right: 40, bottom: 30, left: 40});
  const o = {remove() { try { m.destroy(); } catch (e) {} }};
  MAPS.push(o); return o;
}
function drawMap(el, s, h, hname) {
  if (window.naver && window.naver.maps && !NAVER_FAIL) {
    try { return drawNaver(el, s, h, hname); } catch (e) {}
  }
  if (!window.L) { el.innerHTML = '<div class="mapnone">지도를 불러오지 못했습니다. 인터넷 연결을 확인해 주세요.</div>'; return null; }
  const m = L.map(el, {scrollWheelZoom: false});
  L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {maxZoom: 18, attribution: '© OpenStreetMap'}).addTo(m);
  const a = [s[1], s[0]], b = [h[1], h[0]];
  L.polyline([a, b], {color: '#2F5BEA', weight: 3, dashArray: '6 6'}).addTo(m);
  L.circleMarker(a, {radius: 8, color: '#fff', weight: 2, fillColor: '#FF9A52', fillOpacity: 1}).addTo(m).bindTooltip('거주지', {permanent: true, direction: 'top'});
  L.circleMarker(b, {radius: 8, color: '#fff', weight: 2, fillColor: '#2F5BEA', fillOpacity: 1}).addTo(m).bindTooltip(esc(hname), {permanent: true, direction: 'top'});
  m.fitBounds([a, b], {padding: [40, 40], maxZoom: 15});
  MAPS.push(m); setTimeout(() => m.invalidateSize(), 60);
  return m;
}
function mapNote(s, h) { return `거주지에서 병원까지 약 ${num(kmBetween(s, h))}km (동 기준 대략적인 위치)`; }
function toggleAlt(el, i, k) {
  const r = DATA.results[i], a = r.alternatives[k], box = el.nextElementSibling;
  if (box.classList.contains('open')) {
    box.classList.remove('open');
    if (box._m) { try { box._m.remove(); } catch (e) {} MAPS = MAPS.filter(x => x !== box._m); box._m = null; }
    box.innerHTML = ''; return;
  }
  box.classList.add('open');
  if (!r.s_ll || !a.h_ll) { box.innerHTML = '<div class="mapnone">위치 정보가 없어 지도를 표시할 수 없습니다.</div>'; return; }
  box.innerHTML = '<div class="map"></div><div class="sub" style="margin-top:6px">' + mapNote(r.s_ll, a.h_ll) + '</div>';
  box._m = drawMap(box.firstChild, r.s_ll, a.h_ll, a.hospital);
}
function barcode(seed) {
  let h = 2166136261, x = '', p = 0;
  for (const c of String(seed)) { h = Math.imul(h ^ c.charCodeAt(0), 16777619) >>> 0; }
  for (let i = 0; i < 34; i++) { h = Math.imul(h ^ (h >>> 13), 1274126177) >>> 0; const w = 1 + (h % 3); if (i % 2 === 0) x += `<rect x="${p}" y="0" width="${w}" height="26"/>`; p += w; }
  return `<svg class="bcs" viewBox="0 0 ${p} 26" preserveAspectRatio="none" fill="currentColor">${x}</svg>`;
}
function openDetail(i) {
  killMaps();
  const r = DATA.results[i], w = DATA.weights;
  const dc = r.dist_score * w.distance / 100, pc = r.pref_score * w.preference / 100;
  const alts = (r.alternatives || []).map((a, k) => `<div class="alt click" onclick="toggleAlt(this, ${i}, ${k})"><span>${esc(a.hospital)} <span class="sub">· ${esc(a.preference_status)} · 통학 ${esc(a.travel_time_minutes)}분</span><span class="lnk">지도</span></span><b>${esc(a.composite_score)}점</b></div><div class="altmap"></div>`).join('');
  const note = r.is_top_choice === false
    ? `이 학생의 최고점 병원은 <b>${esc(r.best_hospital)}</b>이지만, 병원 정원과 전체 최적화(모든 학생의 점수 합 최대화) 결과 <b>${esc(r.hospital)}</b>에 배정되었습니다.`
    : (r.is_top_choice === true ? '이 학생에게 가장 점수가 높은 병원에 배정되었습니다.' : '');
  const hasMap = r.s_ll && r.h_ll;
  $('sheet').innerHTML = `<button class="x" onclick="closeModal()" aria-label="닫기">×</button>
    <h3>${esc(mName(r.name))} <span class="sub" style="font-weight:500">${esc(mId(r.student_id))}</span></h3>
    <div class="sub">${esc(r.gender)} · GPA ${esc(r.gpa)} · ${esc(r.birth_year)}년생 · ${esc(r.address || '주소 없음')}</div>
    <div class="box" style="background:rgba(255,255,255,.6)"><b style="color:var(--blue-d)">${esc(r.hospital)}</b> · 병원 내 ${esc(r.rank)}순위 · 종합 <b>${esc(r.composite_score)}점</b></div>
    <div class="box"><h4>거주지와 병원 위치</h4>${hasMap ? '<div class="map" id="mapMain"></div><div class="sub" style="margin-top:6px">' + mapNote(r.s_ll, r.h_ll) + '</div>' : '<div class="mapnone">위치 정보가 없어 지도를 표시할 수 없습니다.</div>'}</div>
    <div class="box"><h4>종합 매칭점수 산출</h4>
      <div class="stack"><i style="width:${dc}%;background:var(--teal)"></i><i style="width:${pc}%;background:var(--amber)"></i></div>
      <div class="eq"><span><i class="dot" style="background:var(--teal)"></i>통학 점수 ${num(r.dist_score)} × ${w.distance}%</span><b>${num(dc)}</b>
        <span><i class="dot" style="background:var(--amber)"></i>지망 점수 ${num(r.pref_score)} (${esc(r.preference_status)}) × ${w.preference}%</span><b>${num(pc)}</b>
        <span><b style="text-align:left">종합 매칭점수</b></span><b>${esc(r.composite_score)}</b></div></div>
    <div class="box"><h4>통학 (${r.transit_source === 'TMAP' ? 'TMAP 대중교통 · ' + (CONFIG.commute_label || '평일 오전 6시 출발') + ' 기준' : (r.transit_source === '직선거리' ? '지도 기반 예상 소요시간' : '규칙 기반 추정')})</h4><div class="eq"><span>소요시간</span><b>${esc(r.travel_time_minutes)}분</b><span>환승</span><b>${esc(r.transfers)}회</b><span>피로지수(MFI)</span><b>${esc(r.fatigue_index)}</b><span>예상 만족도 지수</span><b>${esc(r.satisfaction_index)}</b></div></div>
    ${alts ? `<div class="box"><h4>다른 병원 배정 시 (병원을 누르면 지도 표시)</h4>${alts}</div>` : ''}
    ${note ? `<div class="hint" style="margin-top:12px">${note}</div>` : ''}`;
  $('modal').classList.add('open');
  if (hasMap) drawMap($('mapMain'), r.s_ll, r.h_ll, r.hospital);
}
function closeModal() { $('modal').classList.remove('open'); killMaps(); }
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


LANDING_PAGE = r"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="theme-color" content="#3E5BB8">
<title>AI기반 간호실습 배치 의사결정 프로그램 | 로켓단</title>
<meta name="description" content="통학 거리와 학생 희망을 함께 계산해, 학생에게 더 잘 맞는 간호학과 실습 병원을 추천합니다. 교수자는 근거를 보고 확정합니다.">
<link rel="canonical" href="__BASE__/">
<meta property="og:type" content="website">
<meta property="og:title" content="AI기반 간호실습 배치 의사결정 프로그램 | 로켓단">
<meta property="og:description" content="통학 거리와 학생 희망을 함께 계산해 실습 병원 배정안을 추천합니다.">
<meta property="og:url" content="__BASE__/">
<script type="application/ld+json">{"@context":"https://schema.org","@type":"WebApplication","name":"AI기반 간호실습 배치 의사결정 프로그램 (로켓단)","url":"__BASE__/","applicationCategory":"EducationalApplication","inLanguage":"ko","description":"통학 거리와 학생 희망을 함께 계산해 간호학과 실습 병원 배정안을 추천하는 웹 서비스"}</script>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/variable/pretendardvariable-dynamic-subset.min.css">
<script>document.documentElement.className += ' js';</script>
<style>
  @font-face { font-family:'BMJUA'; src:url(/assets/jua.woff2) format('woff2'); font-weight:400; font-style:normal; font-display:swap; }
  :root {
    --ink:#0C1A3A; --ink-2:#2B3A5E; --muted:#52617F;
    --blue:#3E5BB8; --blue-d:#2D4590; --blue-l:#EAEFFB; --blue-m:#8EA2E0;
    --liner:#F1F4F9; --hair:rgba(12,26,58,.12);
    --gold:#FFD34D; --gold-d:#6B5000; --gold-l:#FFF6D6;
    --green:#0B7A5A; --green-l:#DCF4EA; --red:#C62A44;
    --r:14px; --r-s:8px;
    --lift:0 1px 0 rgba(12,26,58,.06), 0 14px 30px rgba(20,32,80,.22);
    --lift-s:0 1px 0 rgba(12,26,58,.06), 0 8px 20px rgba(20,32,80,.10);
    --ease:cubic-bezier(.16,1,.3,1);
    --font:'BMJUA','Pretendard Variable','Pretendard',-apple-system,'Apple SD Gothic Neo','Noto Sans KR','Malgun Gothic',sans-serif;
  }
  * { box-sizing:border-box; }
  html { font-synthesis:none; scroll-behavior:smooth; scroll-snap-type:y proximity; -webkit-text-size-adjust:100%; scrollbar-color:var(--blue) var(--liner); }
  body { margin:0; font-family:var(--font); color:var(--ink); background:var(--liner); line-height:1.65; word-break:keep-all; overflow-wrap:break-word; font-feature-settings:'tnum' 0; }
  ::selection { background:var(--gold); color:var(--ink); }
  h1, h2, h3 { text-wrap:balance; margin:0; }
  p { text-wrap:pretty; margin:0; }
  a { color:inherit; }
  :focus-visible { outline:3px solid var(--gold); outline-offset:3px; box-shadow:0 0 0 6px var(--ink); border-radius:6px; }
  .wrap { max-width:1240px; margin:0 auto; padding:0 28px; width:100%; }
  .num { font-variant-numeric:tabular-nums; }
  .skip { position:absolute; left:-999px; top:8px; background:var(--gold); color:var(--ink); padding:10px 16px; border-radius:var(--r-s); font-weight:800; z-index:60; }
  .skip:focus { left:12px; }

  /* 공지 줄 */
  .notice { background:var(--ink); color:#E4EAF7; font-size:13px; text-align:center; padding:8px 14px; line-height:1.4; }

  /* 버튼: 노란 라벨(주요) / 흰 윤곽(보조) */
  .btn { display:inline-flex; align-items:center; justify-content:center; min-height:48px; padding:0 26px; border-radius:12px; background:var(--gold); color:var(--ink) !important; font:800 16px/1 var(--font); letter-spacing:-.01em; text-decoration:none; border:2px solid var(--ink); cursor:pointer; box-shadow:0 3px 0 var(--ink); transition:transform .14s var(--ease), box-shadow .14s var(--ease), background .14s; }
  .btn:hover { transform:translateY(-2px); box-shadow:0 5px 0 var(--ink); }
  .btn:active { transform:translateY(2px); box-shadow:0 1px 0 var(--ink); }
  .btn.sm { min-height:40px; padding:0 18px; font-size:14px; }
  .btn.ghost { background:#fff; }
  .btn.ghost:hover { background:var(--blue-l); }

  /* 위에 떠 있는 라벨 머리글 */
  header.top { position:sticky; top:0; z-index:30; }
  header.top .wrap { position:relative; height:0; }
  .bar { position:absolute; left:28px; right:28px; top:12px; height:58px; display:flex; align-items:center; justify-content:space-between; gap:20px; padding:0 10px 0 22px; background:#fff; border-radius:var(--r); box-shadow:var(--lift); }
  .logo { display:flex; align-items:center; gap:10px; text-decoration:none; }
  .logo b { font-size:21px; font-weight:900; letter-spacing:-.01em; color:var(--ink); }
  .ver { font-size:12px; font-weight:800; color:var(--blue-d); background:var(--blue-l); padding:3px 9px; border-radius:99px; letter-spacing:0; }
  .logo small { font-size:13px; font-weight:600; color:var(--muted); letter-spacing:-.01em; }
  .nav { display:flex; align-items:center; gap:26px; }
  .nav a:not(.btn) { position:relative; font-size:15px; font-weight:650; color:var(--ink-2); text-decoration:none; padding:6px 0; background:linear-gradient(var(--blue),var(--blue)) left bottom/0 2px no-repeat; transition:background-size .3s var(--ease), color .2s; }
  .nav a:not(.btn):hover { color:var(--ink); background-size:100% 2px; }

  /* 구역(라벨 시트 한 장 = 한 화면) */
  main > section.sh { min-height:100vh; min-height:100svh; display:flex; flex-direction:column; justify-content:center; padding:112px 0 72px; scroll-snap-align:start; }
  main > section.sh.hero { min-height:calc(100svh - 34px); }
  main > section.sh#steps, main > section.sh.fin, main > section.sh#team { min-height:0; padding:120px 0 104px; }
  .on-blue { background:var(--blue); color:#fff; }
  .on-white { background:#fff; color:var(--ink); }
  .on-liner { background:var(--liner); color:var(--ink); }
  .sec-h { margin-bottom:34px; max-width:46em; }
  .sec-h h2 { font-size:clamp(30px,3.7vw,50px); line-height:1.16; font-weight:850; letter-spacing:-.01em; }
  .sec-h p { margin-top:14px; font-size:18px; line-height:1.6; color:var(--muted); max-width:36em; }
  .on-blue .sec-h p { color:#E3E9FA; }

  /* 라벨: 이 페이지의 유일한 카드 형태 */
  .lb { background:#fff; color:var(--ink); border-radius:var(--r); padding:28px; box-shadow:var(--lift-s); position:relative; }
  .on-blue .lb { box-shadow:var(--lift); }
  .lb h3 { font-size:20px; font-weight:800; letter-spacing:-.01em; line-height:1.3; }
  .lb p { color:var(--muted); font-size:16px; margin-top:8px; }
  .tag { display:inline-flex; align-items:center; height:26px; padding:0 11px; border-radius:99px; background:var(--liner); color:var(--ink-2); font-size:13px; font-weight:700; white-space:nowrap; letter-spacing:-.01em; }
  .tag.hit { background:var(--green-l); color:var(--green); }
  .tag.gold { background:var(--gold-l); color:var(--gold-d); }
  .bc { display:block; color:var(--ink); }

  /* 1. 첫 화면 */
  .hero-grid { display:grid; grid-template-columns:minmax(0,8fr) minmax(0,4fr); gap:12px; align-items:stretch; }
  .lb-main { padding:48px 52px 44px; display:flex; flex-direction:column; }
  .lb-main h1 { font-size:clamp(40px,5vw,68px); line-height:1.12; font-weight:900; letter-spacing:-.01em; }
  .lb-main .sub { margin-top:22px; font-size:19px; line-height:1.6; color:var(--ink-2); max-width:31em; }
  .lb-main .fact { color:var(--ink); margin-top:14px; font-size:19px; font-weight:800; line-height:1.5; letter-spacing:-.01em; max-width:30em; }
  .lb-main .cta { display:flex; flex-wrap:wrap; gap:12px; margin-top:30px; }
  .lb-main .bc { position:absolute; right:30px; bottom:28px; width:118px; height:26px; opacity:.9; }
  .roll { display:grid; grid-template-rows:repeat(4,1fr); gap:12px; }
  .st { padding:18px 22px 16px; display:flex; flex-direction:column; justify-content:center; gap:6px; }
  .st-top { display:flex; justify-content:space-between; align-items:baseline; gap:10px; }
  .st-h { font-size:17px; font-weight:800; letter-spacing:-.01em; line-height:1.3; }
  .st-sc { font-size:22px; font-weight:900; color:var(--blue); letter-spacing:-.01em; font-variant-numeric:tabular-nums; }
  .st-id { font-size:14px; color:var(--muted); font-weight:600; }
  .st-tags { display:flex; gap:6px; flex-wrap:wrap; }
  @keyframes print { from { clip-path:inset(0 0 100% 0); transform:translateY(-10px); } to { clip-path:inset(0 0 0 0); transform:none; } }
  .roll .st { animation:print .8s var(--ease) backwards; animation-delay:calc(.35s + var(--i) * .16s); }
  a.st { text-decoration:none; transition:transform .2s var(--ease), box-shadow .2s var(--ease); }

  /* 2. 왜 필요한가: 크기가 다른 라벨 */
  .mosaic { display:grid; grid-template-columns:repeat(12,minmax(0,1fr)); gap:12px; }
  .m1 { grid-column:span 5; } .m2 { grid-column:span 7; } .m3 { grid-column:span 7; } .m4 { grid-column:span 5; }
  .mosaic .lb { display:flex; flex-direction:column; padding:30px 32px 28px; min-height:218px; }
  .mosaic .lb .viz { margin-top:auto; padding-top:22px; }
  .v-route { display:flex; align-items:center; gap:10px; font-size:13px; font-weight:700; color:var(--ink-2); }
  .v-route .pt { display:inline-flex; align-items:center; gap:7px; white-space:nowrap; }
  .v-route .pt::before { content:''; width:14px; height:14px; border-radius:50%; border:3px solid var(--ink); background:#fff; }
  .v-route .pt.b::before { background:var(--blue); border-color:var(--blue); }
  .v-route .ln { flex:1; height:0; border-top:3px dotted var(--blue-m); position:relative; min-width:40px; }
  .v-route .tm { position:absolute; left:50%; top:-26px; transform:translateX(-50%); background:var(--blue-l); color:var(--blue-d); border-radius:99px; padding:2px 10px; font-size:13px; font-weight:800; white-space:nowrap; }
  .v-pref { display:flex; align-items:center; gap:8px; flex-wrap:wrap; }
  .v-pref .split { flex:1; min-width:160px; height:12px; border-radius:99px; overflow:hidden; display:flex; background:var(--liner); }
  .v-pref .split i { display:block; height:100%; }
  .v-score { display:grid; gap:8px; }
  .v-score .stack { display:flex; height:14px; border-radius:99px; overflow:hidden; background:var(--liner); }
  .v-score .stack i { display:block; height:100%; }
  .v-score .row { display:flex; justify-content:space-between; font-size:14px; font-weight:700; color:var(--ink-2); font-variant-numeric:tabular-nums; }
  .v-sign { display:flex; align-items:center; gap:10px; font-size:14px; font-weight:800; }
  .v-sign span { padding:9px 14px; border-radius:var(--r-s); background:var(--blue-l); color:var(--blue-d); }
  .v-sign span.last { background:#fff; border:2px dashed var(--ink-2); color:var(--ink); }
  .v-sign .arrow { width:34px; height:2px; background:var(--ink-2); position:relative; padding:0; border-radius:0; }
  .v-sign .arrow::after { content:''; position:absolute; right:-1px; top:-4px; border:5px solid transparent; border-left:8px solid var(--ink-2); border-right:0; }

  /* 3. 비교 */
  .cmp-wrap { padding:8px 12px 12px; }
  table.cmp { width:100%; border-collapse:separate; border-spacing:0; font-size:17px; }
  .cmp th, .cmp td { padding:20px 24px; text-align:left; vertical-align:top; border-bottom:1px solid var(--hair); line-height:1.5; }
  .cmp tbody tr:last-child td { border-bottom:none; }
  .cmp thead th { font-size:15px; font-weight:800; color:var(--muted); padding-top:18px; padding-bottom:14px; }
  .cmp thead th.us { color:#fff; background:var(--blue); border-radius:10px 10px 0 0; border-bottom:none; }
  .cmp td:first-child { font-weight:800; width:21%; letter-spacing:-.01em; }
  .cmp td.old { color:var(--muted); width:33%; }
  .cmp td.new { background:var(--blue-l); color:var(--ink); font-weight:650; border-bottom-color:rgba(62,91,184,.14); }
  .cmp tbody tr:last-child td.new { border-radius:0 0 10px 10px; }

  /* 4. 숫자: 무작위 배정과 나란히 */
  .res { padding:12px 34px 14px; }
  .rrow { display:grid; grid-template-columns:minmax(0,1.05fr) minmax(0,2.4fr) minmax(0,1.15fr); gap:34px; align-items:center; padding:24px 0; border-bottom:1px solid var(--hair); }
  .rrow:last-child { border-bottom:none; }
  .rrow h3 { font-size:19px; font-weight:800; letter-spacing:-.01em; }
  .rrow .bars { display:grid; gap:10px; }
  .bar-l { display:grid; grid-template-columns:78px 1fr 84px; align-items:center; gap:12px; font-size:14px; font-weight:700; color:var(--muted); }
  .bar-l .trk { height:16px; border-radius:99px; background:var(--liner); overflow:hidden; }
  .bar-l .trk i { display:block; height:100%; border-radius:99px; background:#B9C4D8; width:var(--w); transition:width 1.1s var(--ease); }
  .bar-l.us { color:var(--ink); }
  .bar-l.us .trk i { background:var(--blue); }
  .bar-l .v { text-align:right; font-weight:800; color:var(--ink); font-size:17px; font-variant-numeric:tabular-nums; letter-spacing:-.01em; }
  .bar-l .v small { font-size:12px; font-weight:700; color:var(--muted); margin-left:1px; }
  .js .res:not(.go) .trk i { width:0; }
  .gain { justify-self:start; display:inline-flex; align-items:center; min-height:34px; padding:0 14px; border-radius:99px; background:var(--green-l); color:var(--green); font-size:15px; font-weight:800; letter-spacing:-.01em; }
  .zero { display:flex; align-items:baseline; gap:6px; }
  .zero b { font-size:46px; font-weight:900; letter-spacing:-.01em; line-height:1; color:var(--blue); font-variant-numeric:tabular-nums; }
  .zero > span { font-size:17px; font-weight:700; color:var(--ink-2); }
  .on-blue .note { margin-top:18px; font-size:14px; color:#E3E9FA; max-width:60em; }

  /* 5. 주요 기능 */
  .svc { display:grid; grid-template-columns:minmax(0,7fr) minmax(0,5fr); gap:12px; align-items:stretch; }
  .svc .lb.l { padding:38px 42px; }
  .svc h3 { font-size:clamp(23px,2.3vw,30px); letter-spacing:-.01em; line-height:1.28; }
  .svc ul { margin:22px 0 28px; padding:0; list-style:none; display:grid; gap:10px; }
  .svc li { position:relative; padding-left:26px; font-size:16.5px; color:var(--ink-2); font-weight:550; }
  .svc li::before { content:''; position:absolute; left:2px; top:.55em; width:12px; height:7px; border-left:3px solid var(--blue); border-bottom:3px solid var(--blue); transform:rotate(-45deg); }
  .receipt { padding:32px 32px 26px; display:flex; flex-direction:column; }
  .receipt h4 { margin:0; font-size:17px; font-weight:800; letter-spacing:-.01em; }
  .receipt .total { font-size:60px; font-weight:900; letter-spacing:-.01em; line-height:1; margin:16px 0 4px; font-variant-numeric:tabular-nums; color:var(--blue); }
  .receipt .total small { font-size:20px; letter-spacing:-.01em; color:var(--ink-2); font-weight:800; margin-left:4px; }
  .receipt .stack { display:flex; height:16px; border-radius:99px; overflow:hidden; background:var(--liner); margin:18px 0 14px; }
  .receipt .stack i { display:block; height:100%; }
  .receipt .row { display:flex; justify-content:space-between; gap:12px; font-size:15px; padding:8px 0; border-bottom:1px dashed var(--hair); color:var(--ink-2); font-weight:600; }
  .receipt .row b { font-variant-numeric:tabular-nums; color:var(--ink); }
  .receipt .row .d { display:inline-block; width:10px; height:10px; border-radius:3px; margin-right:8px; }
  .receipt .small { margin-top:12px; font-size:13.5px; color:var(--muted); }
  .receipt .bc { margin-top:auto; padding-top:18px; width:100%; height:44px; }

  /* 6. 직접 해보기 */
  .demo { display:grid; grid-template-columns:minmax(0,5fr) minmax(0,7fr); gap:12px; }
  .demo .ctl { padding:34px 36px; display:flex; flex-direction:column; }
  .demo .ctl h3 { font-size:24px; letter-spacing:-.01em; }
  .demo .wl { display:flex; justify-content:space-between; gap:12px; margin-top:26px; font-size:16px; font-weight:800; letter-spacing:-.01em; }
  .demo .wl span:first-child { color:var(--blue-d); }
  .demo .wl span:last-child { color:var(--gold-d); }
  .sl { --v:.7; -webkit-appearance:none; appearance:none; width:100%; height:38px; background:transparent; cursor:pointer; margin:6px 0 0; padding:0; }
  .sl::-webkit-slider-runnable-track { height:12px; border-radius:99px; background:linear-gradient(90deg,var(--blue) 0 calc(14px + (100% - 28px) * var(--v)),var(--gold) calc(14px + (100% - 28px) * var(--v)) 100%); }
  .sl::-moz-range-track { height:12px; border-radius:99px; background:linear-gradient(90deg,var(--blue) 0 calc(14px + (100% - 28px) * var(--v)),var(--gold) calc(14px + (100% - 28px) * var(--v)) 100%); }
  .sl::-webkit-slider-thumb { -webkit-appearance:none; width:30px; height:30px; margin-top:-9px; border-radius:50%; background:#fff; border:4px solid var(--ink); box-shadow:0 3px 8px rgba(12,26,58,.3); transition:transform .15s var(--ease); }
  .sl::-moz-range-thumb { width:22px; height:22px; border-radius:50%; background:#fff; border:4px solid var(--ink); box-shadow:0 3px 8px rgba(12,26,58,.3); }
  .sl:hover::-webkit-slider-thumb { transform:scale(1.1); }
  .sl:active::-webkit-slider-thumb { transform:scale(.94); }
  .sl:focus-visible { outline:none; box-shadow:none; }
  .sl:focus-visible::-webkit-slider-thumb { box-shadow:0 0 0 4px var(--gold), 0 0 0 7px var(--ink); }
  .demo .hint { margin-top:14px; font-size:14px; color:var(--muted); }
  .demo .go { margin-top:auto; padding-top:26px; }
  #dRows { display:grid; gap:10px; align-content:start; }
  .dr { position:relative; display:grid; grid-template-columns:1fr auto; gap:6px 14px; padding:16px 22px 16px; background:#fff; border-radius:var(--r); box-shadow:var(--lift-s); transition:transform .45s var(--ease), box-shadow .25s; }
  .dr.top { box-shadow:0 0 0 3px var(--gold), var(--lift-s); }
  .dr .nm { font-size:18px; font-weight:800; letter-spacing:-.01em; display:flex; align-items:center; gap:8px; flex-wrap:wrap; }
  .dr .flag { display:none; align-items:center; height:22px; padding:0 14px 0 9px; background:var(--gold); color:var(--ink); font-size:12px; font-weight:900; font-style:normal; clip-path:polygon(0 0,100% 0,calc(100% - 7px) 50%,100% 100%,0 100%); }
  .dr.top .flag { display:inline-flex; }
  .dr .sub { font-size:13.5px; color:var(--muted); font-weight:600; }
  .dr .sc { font-size:26px; font-weight:900; letter-spacing:-.01em; color:var(--blue); text-align:right; font-variant-numeric:tabular-nums; align-self:center; grid-row:1 / 3; grid-column:2; }
  .dr .tr { grid-column:1 / 3; height:8px; border-radius:99px; background:var(--liner); overflow:hidden; }
  .dr .tr i { display:block; height:100%; background:var(--blue); border-radius:99px; transition:width .4s var(--ease); }
  .dr .st2 { display:flex; flex-direction:column; gap:2px; }

  /* 7. 순서: 절취선으로 이어진 한 줄 라벨 */
  .strip { display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); padding:0; }
  .stp { position:relative; padding:36px 38px 40px; min-height:236px; display:flex; flex-direction:column; }
  .stp + .stp { border-left:3px dashed rgba(12,26,58,.28); }
  .stp + .stp::before, .stp + .stp::after { content:''; position:absolute; left:-12px; width:22px; height:22px; border-radius:50%; background:var(--liner); }
  .stp + .stp::before { top:-11px; } .stp + .stp::after { bottom:-11px; }
  .stp .n { width:46px; height:46px; border-radius:50%; background:var(--gold); color:var(--ink); border:2px solid var(--ink); font-size:21px; font-weight:900; display:grid; place-items:center; margin-bottom:22px; font-variant-numeric:tabular-nums; }
  .stp h3 { font-size:22px; font-weight:800; letter-spacing:-.01em; }
  .stp p { color:var(--muted); font-size:16.5px; margin-top:10px; max-width:22em; }
  .steps-cta { margin-top:30px; }

  /* 8. 자주 묻는 질문 */
  .guide-grid { display:grid; grid-template-columns:minmax(0,4fr) minmax(0,8fr); gap:56px; align-items:start; }
  .guide-l { position:sticky; top:120px; }
  .guide-l .sec-h { margin-bottom:26px; }
  .faq { display:grid; gap:10px; }
  .faq details { background:#fff; border-radius:var(--r); box-shadow:var(--lift-s); interpolate-size:allow-keywords; }
  .faq summary { list-style:none; cursor:pointer; padding:21px 26px; font-size:18px; font-weight:750; letter-spacing:-.01em; display:flex; justify-content:space-between; align-items:center; gap:16px; border-radius:var(--r); }
  .faq summary::-webkit-details-marker { display:none; }
  .faq summary::after { content:''; flex:none; width:28px; height:28px; border-radius:50%; background:linear-gradient(var(--ink),var(--ink)) center/12px 2px no-repeat, linear-gradient(var(--ink),var(--ink)) center/2px 12px no-repeat, var(--blue-l); transition:transform .3s var(--ease); }
  .faq details[open] summary::after { transform:rotate(45deg); }
  .faq summary:hover { background:var(--blue-l); }
  .faq details::details-content { block-size:0; overflow:clip; transition:block-size .32s var(--ease), content-visibility .32s allow-discrete; }
  .faq details[open]::details-content { block-size:auto; }
  .faq p { padding:0 26px 24px; color:var(--muted); font-size:16.5px; max-width:44em; margin:0; }
  .guide-act { margin-bottom:0; }

  /* 9. 만든 사람들: 명찰 */
  .team { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:12px; }
  .mem { padding:44px 18px 26px; text-align:center; display:flex; flex-direction:column; align-items:center; }
  .mem::before { content:''; position:absolute; top:16px; left:50%; width:52px; height:10px; margin-left:-26px; border-radius:99px; background:var(--liner); box-shadow:inset 0 2px 3px rgba(12,26,58,.28); }
  .mem img { height:128px; width:auto; image-rendering:pixelated; image-rendering:crisp-edges; }
  .mem b { display:block; margin-top:18px; font-size:21px; font-weight:850; letter-spacing:-.01em; }
  .mem span { margin-top:6px; font-size:14.5px; color:var(--muted); font-weight:600; }
  .mem .tag { margin-top:10px; }

  /* 마무리 */
  .final { padding:56px 56px 52px; display:grid; grid-template-columns:minmax(0,8fr) auto; gap:30px; align-items:end; }
  .final h2 { font-size:clamp(32px,4.2vw,58px); font-weight:900; letter-spacing:-.01em; line-height:1.14; }
  .final p { margin-top:18px; color:var(--ink-2); font-size:18px; max-width:30em; text-wrap:balance; }

  footer.foot { background:var(--ink); color:#C7D1E8; padding:44px 0 54px; font-size:14.5px; line-height:1.8; }
  footer.foot .links { display:flex; gap:22px; flex-wrap:wrap; margin-bottom:14px; font-weight:700; }
  footer.foot a { color:#fff; text-underline-offset:4px; }
  footer.foot b { color:#fff; }
  .keyhint { margin-top:10px; font-size:13px; color:#97A5C6; }

  /* 오른쪽 점 메뉴 */
  .dots { position:fixed; right:14px; top:50%; transform:translateY(-50%); z-index:25; display:none; flex-direction:column; gap:2px; }
  .dots a { position:relative; display:block; width:28px; height:24px; border-radius:0; }
  .dots a::before { content:''; position:absolute; right:7px; top:6px; width:12px; height:12px; border-radius:50%; background:#fff; border:2px solid var(--ink); transition:transform .2s var(--ease), background .2s; }
  .dots a:hover::before { transform:scale(1.25); }
  .dots a.on::before { background:var(--gold); transform:scale(1.25); }
  .dots a span { position:absolute; right:36px; top:0; white-space:nowrap; background:var(--ink); color:#fff; font-size:13px; font-weight:700; padding:3px 11px; border-radius:6px; opacity:0; pointer-events:none; transition:opacity .15s; line-height:1.7; }
  .dots a:hover span, .dots a:focus-visible span { opacity:1; }
  .dots kbd { font:800 11px/1 var(--font); background:rgba(255,255,255,.22); border-radius:4px; padding:2px 6px; margin-left:8px; }
  @media (min-width:1320px) { .dots { display:flex; } }

  /* 호버: 누를 수 있는 것만 반응 */
  @media (hover:hover) { a.st:hover { transform:translateY(-2px); } .mem:hover { transform:translateY(-2px); } .mem { transition:transform .2s var(--ease); } }

  @media (max-width:1040px) {
    .hero-grid { grid-template-columns:1fr; }
    .roll { grid-template-columns:repeat(2,minmax(0,1fr)); grid-template-rows:none; }
    .mosaic .lb, .m1, .m2, .m3, .m4 { grid-column:span 12; min-height:0; }
    .svc, .demo { grid-template-columns:1fr; }
    .rrow { grid-template-columns:1fr; gap:14px; }
    .final { grid-template-columns:1fr; align-items:start; }
    .logo small { display:none; }
    .guide-grid { grid-template-columns:1fr; gap:6px; }
    .guide-l { position:static; }
  }
  @media (max-width:860px) {
    .nav a:not(.btn) { display:none; }
    .team { grid-template-columns:repeat(2,minmax(0,1fr)); }
    .strip { grid-template-columns:1fr; }
    .stp { min-height:0; }
    .stp + .stp { border-left:0; border-top:3px dashed rgba(12,26,58,.28); }
    .stp + .stp::before { left:-11px; top:-12px; }
    .stp + .stp::after { left:auto; right:-11px; top:-12px; bottom:auto; }
  }
  @media (max-width:640px) {
    .wrap { padding:0 16px; }
    .bar { left:16px; right:16px; padding:0 8px 0 16px; }
    main > section.sh { padding:96px 0 48px; }
    .lb-main { padding:30px 24px 28px; }
    .lb-main .bc { display:none; }
    .lb-main .sub, .lb-main .fact { font-size:17px; }
    .roll { grid-template-columns:1fr; }
    .lb, .mosaic .lb, .svc .lb.l, .demo .ctl, .receipt, .final { padding:24px 22px; }
    .strip { padding:0; } .stp { padding:26px 22px 28px; }
    .res { padding:4px 20px 6px; }
    .bar-l { grid-template-columns:62px 1fr 70px; gap:8px; }
    .btn { width:100%; }
    .btn.sm { width:auto; }
    .cmp-wrap { padding:4px 4px 8px; }
    table.cmp, .cmp tbody, .cmp tr, .cmp td { display:block; }
    .cmp thead { display:none; }
    .cmp tr { padding:14px 14px 6px; border-bottom:1px solid var(--hair); }
    .cmp tbody tr:last-child { border-bottom:none; }
    .cmp td { padding:4px 0 10px !important; border:0 !important; width:auto !important; border-radius:0 !important; }
    .cmp td:first-child { font-size:18px; padding-bottom:6px !important; }
    .cmp td.old::before, .cmp td.new::before { display:block; font-size:12.5px; font-weight:800; color:var(--muted); margin-bottom:2px; }
    .cmp td.old::before { content:'엑셀 수작업'; }
    .cmp td.new::before { content:'로켓단'; color:var(--blue-d); }
    .cmp td.new { background:var(--blue-l); padding:10px 12px !important; border-radius:8px !important; margin-bottom:8px; }
    .team { gap:10px; }
    .mem img { height:104px; }
    .faq summary { font-size:16.5px; padding:18px 18px; }
    .faq p { padding:0 18px 20px; font-size:15.5px; }
  }
  @media (prefers-reduced-motion:reduce) {
    html { scroll-behavior:auto; scroll-snap-type:none; }
    *, *::before, *::after { animation:none !important; transition:none !important; }
  }



  /* ===== 유리(글래스모피즘) 층: 맑은 청록·하늘 오로라 위에 반투명 유리 카드 ===== */
  :root {
    --ink:#0B1B3A; --ink-2:#2A3F66; --muted:#4F6285;
    --blue:#2F5BEA; --blue-d:#1E3FA8; --blue-l:rgba(47,91,234,.13); --blue-m:#86A5F5;
    --liner:#E6EEFB; --hair:rgba(11,27,58,.12);
    --gold:#FF9A52; --gold-d:#8A3D00; --gold-l:rgba(255,154,82,.26);
    --green:#0A7A55; --green-l:rgba(16,185,129,.17); --red:#C2334D;
    --r:22px; --r-s:12px;
    --glass:rgba(255,255,255,.50); --glass-hi:rgba(255,255,255,.72); --glass-edge:rgba(255,255,255,.85);
    --lift:0 18px 50px rgba(30,63,168,.20), inset 0 1px 0 rgba(255,255,255,.9);
    --lift-s:0 10px 34px rgba(30,63,168,.13), inset 0 1px 0 rgba(255,255,255,.9), inset 0 -1px 0 rgba(255,255,255,.3);
  }
  html { scrollbar-color:#8FA9E8 #E6EEFB; }
  body { background:#E6EEFB; position:relative; isolation:isolate; }
  body::before { content:''; position:fixed; inset:-18%; z-index:-2; pointer-events:none;
    background:
      radial-gradient(38% 34% at 14% 16%, #38D6F0 0%, transparent 70%),
      radial-gradient(34% 32% at 86% 14%, #6A93FF 0%, transparent 70%),
      radial-gradient(40% 36% at 80% 84%, #FFB48F 0%, transparent 72%),
      radial-gradient(34% 34% at 10% 86%, #7DF0D0 0%, transparent 70%),
      radial-gradient(30% 28% at 50% 48%, #D3E2FF 0%, transparent 75%),
      linear-gradient(180deg,#F0F5FF,#E1ECFB);
    animation:aur 38s ease-in-out infinite alternate; }
  body::after { content:''; position:fixed; inset:-10%; z-index:-1; pointer-events:none; opacity:.9;
    background:
      radial-gradient(14% 16% at 30% 30%, rgba(255,255,255,.75), transparent 70%),
      radial-gradient(12% 14% at 72% 58%, rgba(255,255,255,.55), transparent 70%),
      radial-gradient(16% 14% at 40% 82%, rgba(255,180,143,.55), transparent 70%),
      radial-gradient(12% 12% at 90% 40%, rgba(106,147,255,.55), transparent 70%);
    animation:aur2 46s ease-in-out infinite alternate; }
  @keyframes aur { from { transform:translate3d(-2%,-1%,0) scale(1); } to { transform:translate3d(3%,2%,0) scale(1.08); } }
  @keyframes aur2 { from { transform:translate3d(2%,2%,0); } to { transform:translate3d(-3%,-2%,0) scale(1.06); } }
  ::selection { background:#BCD0FF; color:var(--ink); }
  :focus-visible { outline:3px solid #2F5BEA; outline-offset:3px; box-shadow:0 0 0 6px rgba(255,255,255,.9); }

  .notice { background:rgba(11,27,58,.88); backdrop-filter:blur(10px); -webkit-backdrop-filter:blur(10px); }
  .on-blue, .on-liner { background:transparent; color:var(--ink); }
  .on-white { background:rgba(255,255,255,.26); color:var(--ink); }
  .on-blue .sec-h p, .on-blue .note { color:var(--muted); }

  /* 유리 카드 */
  .lb, .bar, .faq details, .dr {
    background:var(--glass); border:1px solid var(--glass-edge);
    -webkit-backdrop-filter:blur(22px) saturate(170%); backdrop-filter:blur(22px) saturate(170%);
    box-shadow:var(--lift-s); }
  .lb::after { content:''; position:absolute; inset:0; border-radius:inherit; pointer-events:none;
    background:linear-gradient(135deg, rgba(255,255,255,.55) 0%, rgba(255,255,255,0) 38%); }
  .lb > * { position:relative; z-index:1; }
  .on-blue .lb { box-shadow:var(--lift); }
  .bar { background:rgba(255,255,255,.62); border-radius:20px; box-shadow:var(--lift); }
  .bar .btn { position:relative; z-index:1; }
  .ver { background:rgba(47,91,234,.14); color:var(--blue-d); }
  .tag { background:rgba(11,27,58,.07); }
  .tag.hit { background:var(--green-l); color:var(--green); }
  .tag.gold { background:var(--gold-l); color:var(--gold-d); }

  /* 버튼: 청록 유리 광택(주요) / 흰 유리(보조) */
  .btn { background:linear-gradient(135deg,#3A6AF3 0%,#2347C9 100%) !important; color:#fff !important;
    border:1px solid rgba(255,255,255,.55) !important; border-radius:14px;
    box-shadow:0 10px 24px rgba(35,71,201,.32), inset 0 1px 0 rgba(255,255,255,.45) !important; }
  .btn:hover { transform:translateY(-2px); box-shadow:0 14px 30px rgba(35,71,201,.4), inset 0 1px 0 rgba(255,255,255,.5) !important; }
  .btn:active { transform:translateY(1px); box-shadow:0 4px 10px rgba(35,71,201,.3), inset 0 1px 0 rgba(255,255,255,.4) !important; }
  .btn.ghost { background:rgba(255,255,255,.62) !important; color:var(--ink) !important;
    -webkit-backdrop-filter:blur(14px); backdrop-filter:blur(14px);
    box-shadow:0 8px 20px rgba(30,63,168,.12), inset 0 1px 0 rgba(255,255,255,.9) !important; }
  .btn.ghost:hover { background:rgba(255,255,255,.85) !important; }
  .nav a:not(.btn) { background:linear-gradient(var(--blue),var(--blue)) left bottom/0 2px no-repeat; }

  /* 첫 화면 */
  .lb-main h1 { color:var(--ink); }
  .st-sc, .zero b, .receipt .total, .dr .sc { color:var(--blue); }
  .v-route .pt::before { background:rgba(255,255,255,.9); border-color:var(--ink); }
  .v-route .pt.b::before { background:var(--blue); border-color:var(--blue); }
  .v-route .ln { border-top-color:var(--blue-m); }
  .v-route .tm { background:rgba(47,91,234,.14); color:var(--blue-d); }
  .v-sign span { background:rgba(47,91,234,.14); color:var(--blue-d); }
  .v-sign span.last { background:rgba(255,255,255,.7); border-color:var(--ink-2); }
  .v-pref .split, .v-score .stack, .receipt .stack, .dr .tr, .bar-l .trk { background:rgba(11,27,58,.09); }

  /* 비교표 */
  .cmp thead th.us { background:linear-gradient(135deg,#3A6AF3,#2347C9); color:#fff; }
  .cmp td.new { background:rgba(47,91,234,.11); border-bottom-color:rgba(47,91,234,.16); }
  .cmp td.new::before { color:var(--blue-d); }

  /* 숫자 막대 */
  .bar-l .trk i { background:#9DB8C4; }
  .bar-l.us .trk i { background:linear-gradient(90deg,#2F5BEA,#22C7F0); }
  .gain { background:var(--green-l); color:var(--green); }

  /* 슬라이더 */
  .sl::-webkit-slider-runnable-track { background:linear-gradient(90deg,#2F5BEA 0 calc(14px + (100% - 28px) * var(--v)),var(--gold) calc(14px + (100% - 28px) * var(--v)) 100%); }
  .sl::-moz-range-track { background:linear-gradient(90deg,#2F5BEA 0 calc(14px + (100% - 28px) * var(--v)),var(--gold) calc(14px + (100% - 28px) * var(--v)) 100%); }
  .sl::-webkit-slider-thumb { border-color:#fff; background:var(--blue); box-shadow:0 4px 12px rgba(30,63,168,.4); }
  .sl::-moz-range-thumb { border-color:#fff; background:var(--blue); box-shadow:0 4px 12px rgba(30,63,168,.4); }
  .sl:focus-visible::-webkit-slider-thumb { box-shadow:0 0 0 4px rgba(255,255,255,.95), 0 0 0 7px var(--blue); }
  .demo .wl span:last-child { color:var(--gold-d); }
  .dr.top { box-shadow:0 0 0 2px var(--blue), var(--lift-s); background:var(--glass-hi); }
  .dr .flag { background:var(--blue); color:#fff; }
  .dr .tr i { background:linear-gradient(90deg,#2F5BEA,#22C7F0); }

  /* 순서 */
  .stp .n { background:linear-gradient(135deg,#3A6AF3,#2347C9); color:#fff; border:1px solid rgba(255,255,255,.6); box-shadow:0 8px 18px rgba(35,71,201,.3); }
  .stp + .stp { border-left:2px dashed rgba(11,27,58,.22); }
  .stp + .stp::before, .stp + .stp::after { display:none; }

  /* 질문·팀·마무리 */
  .faq summary:hover { background:rgba(255,255,255,.4); }
  .mem::before { background:rgba(11,27,58,.12); box-shadow:inset 0 2px 3px rgba(11,27,58,.2); }
  .final { background:rgba(255,255,255,.58); }
  footer.foot { background:rgba(11,27,58,.9); -webkit-backdrop-filter:blur(14px); backdrop-filter:blur(14px); }
  .dots a::before { background:rgba(255,255,255,.85); border-color:var(--ink); }
  .dots a.on::before { background:var(--blue); border-color:var(--blue); }
  .dots a span { background:rgba(11,27,58,.92); }

  @media (max-width:640px) {
    .lb, .bar, .faq details, .dr { -webkit-backdrop-filter:blur(14px) saturate(150%); backdrop-filter:blur(14px) saturate(150%); }
    .stp + .stp { border-left:0; border-top:2px dashed rgba(11,27,58,.22); }
  }
  @media (prefers-reduced-motion:reduce) { body::before, body::after { animation:none !important; } }

  /* ===== 이동·전환 ===== */
  /* 1) 페이지 사이: 소개 ↔ 배정 도구가 부드럽게 넘어가고, 로고는 제자리에서 이어진다 */
  @view-transition { navigation:auto; }
  .logo b { view-transition-name:brandmark; }
  ::view-transition-old(root) { animation:vtOut .22s ease both; }
  ::view-transition-new(root) { animation:vtIn .42s var(--ease) both; }
  ::view-transition-group(brandmark) { animation-duration:.45s; animation-timing-function:var(--ease); }
  @keyframes vtOut { to { opacity:0; transform:translateY(-8px); } }
  @keyframes vtIn { from { opacity:0; transform:translateY(14px); } }
  /* 2) 읽은 위치: 맨 위 진행선 + 메뉴 밑줄이 현재 구간으로 미끄러진다 */
  .prog { position:fixed; top:0; left:0; right:0; height:3px; z-index:80; pointer-events:none; }
  .prog i { display:block; height:100%; background:var(--blue); transform-origin:left; transform:scaleX(0); }
  .nav { position:relative; }
  .nav a:not(.btn) { background:none !important; transition:color .2s; }
  .nav a:not(.btn):hover, .nav a.cur { color:var(--ink); }
  .nav-ind { position:absolute; left:0; bottom:-2px; height:3px; border-radius:3px; background:var(--blue); width:var(--w,0); transform:translateX(var(--x,0)); opacity:0; transition:transform .5s var(--ease), width .5s var(--ease), opacity .25s; pointer-events:none; }
  .nav-ind.on { opacity:1; }
  /* 3) 점 메뉴: 현재 위치 점이 부드럽게 커진다 */
  .dots a::before { transition:transform .35s var(--ease), background .25s, border-color .25s; }
  /* 4) 눌렀을 때 구간 제목이 한 번 짚어 준다 */
  .sec-h h2.arrive, .final h2.arrive { animation:arrive .7s var(--ease); }
  @keyframes arrive { 0% { transform:translateX(-14px); opacity:.35; } 100% { transform:none; opacity:1; } }
  @media (prefers-reduced-motion:reduce) { .prog i, .nav-ind { transition:none !important; } ::view-transition-old(root), ::view-transition-new(root) { animation:none !important; } }
</style>
</head>
<body>
<div class="prog" aria-hidden="true"><i></i></div>
<a class="skip" href="#intro">본문으로 건너뛰기</a>
<div class="notice">캡스톤디자인 작품 · 학생 명단 파일은 서버에 저장되지 않습니다</div>
<header class="top"><div class="wrap">
  <div class="bar">
    <a class="logo" href="/" aria-label="로켓단 처음으로"><b>로켓단</b><span class="ver">__APP_VERSION__</span><small>AI기반 간호실습 배치 의사결정 프로그램</small></a>
    <nav class="nav" aria-label="주요 메뉴"><a href="#why">소개</a><a href="#steps">순서</a><a href="#guide">자주 묻는 질문</a><a href="#team">만든 사람들</a><a class="btn sm" href="/app">배정 시스템 열기</a></nav>
  </div>
</div></header>

<main>
<nav class="dots" aria-label="구역 이동"><a href="#intro" data-s="intro" aria-label="처음"><span>처음<kbd>1</kbd></span></a><a href="#why" data-s="why" aria-label="왜 필요한가"><span>왜 필요한가<kbd>2</kbd></span></a><a href="#compare" data-s="compare" aria-label="비교"><span>비교<kbd>3</kbd></span></a><a href="#numbers" data-s="numbers" aria-label="시험 결과"><span>시험 결과<kbd>4</kbd></span></a><a href="#service" data-s="service" aria-label="주요 기능"><span>주요 기능<kbd>5</kbd></span></a><a href="#start" data-s="start" aria-label="직접 해보기"><span>직접 해보기<kbd>6</kbd></span></a><a href="#steps" data-s="steps" aria-label="순서"><span>순서<kbd>7</kbd></span></a><a href="#guide" data-s="guide" aria-label="자주 묻는 질문"><span>자주 묻는 질문<kbd>8</kbd></span></a><a href="#team" data-s="team" aria-label="만든 사람들"><span>만든 사람들<kbd>9</kbd></span></a></nav>

<section class="sh hero on-blue" id="intro"><div class="wrap">
  <div class="hero-grid">
    <div class="lb lb-main">
      <h1>실습 병원 배정,<br>거리와 희망을<br>함께 계산합니다</h1>
      <p class="sub">통학 거리와 희망 병원을 함께 계산해 학생에게 더 잘 맞는 실습지를 추천하고, 그 근거까지 보여 줍니다. 확정은 교수자가 합니다.</p>
      <p class="fact">가상 학생 221명 시험에서 평균 통학시간이 56.7분에서 35분으로 줄었습니다.</p>
      <div class="cta"><a class="btn" href="/app">배정 시스템 열기</a><a class="btn ghost" href="#start">비중 바꿔 보기</a></div>
    </div>
    <div class="roll" aria-label="배정 결과 예시(가상 학생 데이터, 이름·학번 일부 가림)">
      <article class="lb st" style="--i:0"><div class="st-top"><b class="st-h">고려대학교구로병원</b><span class="st-sc">91.6</span></div><div class="st-id">김○수 · 202○○○07</div><div class="st-tags"><span class="tag hit">1지망</span><span class="tag">통학 약 24분</span></div></article>
      <article class="lb st" style="--i:1"><div class="st-top"><b class="st-h">순천향대학교부천병원</b><span class="st-sc">97.2</span></div><div class="st-id">이○은 · 202○○○15</div><div class="st-tags"><span class="tag hit">1지망</span><span class="tag">통학 약 18분</span></div></article>
      <article class="lb st" style="--i:2"><div class="st-top"><b class="st-h">단원병원</b><span class="st-sc">77.9</span></div><div class="st-id">박○준 · 202○○○31</div><div class="st-tags"><span class="tag hit">2지망</span><span class="tag">통학 약 29분</span></div></article>
      <article class="lb st" style="--i:3"><div class="st-top"><b class="st-h">아주대학교병원</b><span class="st-sc">67.4</span></div><div class="st-id">최○린 · 202○○○42</div><div class="st-tags"><span class="tag">지망 외</span><span class="tag">통학 약 21분</span></div></article>
    </div>
  </div>
</div></section>

<section class="sh on-liner" id="why"><div class="wrap">
  <div class="sec-h"><h2>배정 업무를 더 합리적으로</h2><p>엑셀로 손수 맞추던 실습 배정을, 설명할 수 있는 추천안으로 바꿉니다.</p></div>
  <div class="mosaic">
    <article class="lb m1"><h3>통학 부담 반영</h3><p>지도 기반 거리와 소요시간으로, 멀리 가는 학생을 줄이는 방향으로 계산합니다.</p>
      <div class="viz v-route" aria-hidden="true"><span class="pt a">거주지</span><span class="ln"><span class="tm">약 24분</span></span><span class="pt b">병원</span></div></article>
    <article class="lb m2"><h3>학생 희망 반영</h3><p>1~3지망을 점수에 넣고, 거리와 희망의 비중은 슬라이더로 조절합니다.</p>
      <div class="viz v-pref" aria-hidden="true"><span class="tag hit">1지망</span><span class="tag hit">2지망</span><span class="tag hit">3지망</span><span class="split"><i style="width:70%;background:var(--blue)"></i><i style="width:30%;background:var(--gold)"></i></span></div></article>
    <article class="lb m3"><h3>근거를 보여 주는 결과</h3><p>학생마다 점수 계산 내역과 다른 병원이었을 때의 점수, 위치 지도를 확인할 수 있습니다.</p>
      <div class="viz v-score" aria-hidden="true"><div class="stack"><i style="width:40.1%;background:var(--blue)"></i><i style="width:30%;background:var(--gold)"></i></div><div class="row"><span>통학 40.1</span><span>지망 30.0</span><span>종합 70.1</span></div></div></article>
    <article class="lb m4"><h3>학생에게 더 좋은 배정</h3><p>학생마다 더 잘 맞는 병원을 추천하고, 교수자는 근거를 보고 확정합니다.</p>
      <div class="viz v-sign" aria-hidden="true"><span>추천안</span><span class="arrow"></span><span class="last">교수자 확정</span></div></article>
  </div>
</div></section>

<section class="sh on-liner" id="compare"><div class="wrap">
  <div class="sec-h"><h2>엑셀 수작업과 무엇이 다른가</h2><p>학생에게 더 맞는 배정을 찾고, 교수자는 근거를 보며 확정합니다. 계산과 정리는 시스템이 맡습니다.</p></div>
  <div class="lb cmp-wrap">
    <table class="cmp"><thead><tr><th scope="col"><span class="sr" style="position:absolute;left:-999px">항목</span></th><th scope="col">엑셀 수작업</th><th scope="col" class="us">로켓단</th></tr></thead><tbody>
      <tr><td>통학 거리</td><td class="old">학생 주소를 보고 사람이 판단</td><td class="new">지도 기반 거리와 소요시간을 자동 계산</td></tr>
      <tr><td>학생 희망</td><td class="old">1~3지망을 보며 하나씩 조정</td><td class="new">지망을 점수에 반영해 한 번에 계산</td></tr>
      <tr><td>조건이 바뀔 때</td><td class="old">정원이나 기준이 바뀌면 다시 맞춤</td><td class="new">조건을 바꾸고 다시 계산하면 바로 새 추천안</td></tr>
      <tr><td>결과 설명</td><td class="old">필요할 때 따로 정리</td><td class="new">학생별 점수 근거와 다른 병원 점수를 자동 제공</td></tr>
      <tr><td>최종 결정</td><td class="old">교수자</td><td class="new">교수자 (시스템은 추천안만 제공)</td></tr>
    </tbody></table>
  </div>
</div></section>

<section class="sh on-white" id="numbers"><div class="wrap">
  <div class="sec-h"><h2>숫자로 보는 로켓단</h2><p>가상 학생 221명 · 4학년 1학기 관리 실습 9개 병원으로, 정원만 지킨 무작위 배정 100회 평균과 비교했습니다.</p></div>
  <div class="lb res" id="resBox">
    <div class="rrow"><h3>평균 통학시간</h3>
      <div class="bars"><div class="bar-l"><span>무작위 배정</span><span class="trk"><i style="--w:100%"></i></span><span class="v">56.7<small>분</small></span></div>
        <div class="bar-l us"><span>로켓단</span><span class="trk"><i style="--w:61.7%"></i></span><span class="v"><span class="cnt" data-n="35" data-d="0">35</span><small>분</small></span></div></div>
      <span class="gain">21.7분 단축 (−38%)</span></div>
    <div class="rrow"><h3>지망 충족률</h3>
      <div class="bars"><div class="bar-l"><span>무작위 배정</span><span class="trk"><i style="--w:19.4%"></i></span><span class="v">19.4<small>%</small></span></div>
        <div class="bar-l us"><span>로켓단</span><span class="trk"><i style="--w:76%"></i></span><span class="v"><span class="cnt" data-n="76" data-d="0">76</span><small>%</small></span></div></div>
      <span class="gain">+56.6%p</span></div>
    <div class="rrow"><h3>평균 매칭점수</h3>
      <div class="bars"><div class="bar-l"><span>무작위 배정</span><span class="trk"><i style="--w:41.8%"></i></span><span class="v">41.8<small>점</small></span></div>
        <div class="bar-l us"><span>로켓단</span><span class="trk"><i style="--w:75.6%"></i></span><span class="v"><span class="cnt" data-n="75.6" data-d="1">75.6</span><small>점</small></span></div></div>
      <span class="gain">+33.8점</span></div>
    <div class="rrow"><h3>미배정 학생</h3>
      <div class="zero"><b><span class="cnt" data-n="0" data-d="0">0</span></b><span>명 · 221명 전원 배정</span></div>
      <span class="gain">병원 정원 100% 준수</span></div>
  </div>
  <p class="note">가상 명단 기준 시뮬레이션 결과이며, 통학시간은 지도 기반 예상값입니다. 실제 학과의 배정 결과와 비교한 값이 아닙니다.</p>
</div></section>

<section class="sh on-liner" id="service"><div class="wrap">
  <div class="sec-h"><h2>학생에게 더 맞게 계산하고, 교수자가 확정합니다</h2><p>최적화 알고리즘(헝가리안 알고리즘)이 병원 정원 안에서 학생 전체의 점수 합이 가장 높은 조합을 찾습니다.</p></div>
  <div class="svc">
    <div class="lb l">
      <h3>조건을 정하고 명단을 올리면,<br>추천안이 나옵니다</h3>
      <p>병원별 정원을 지키면서 전체 학생의 점수 합이 가장 높은 배정안을 계산합니다.</p>
      <ul>
        <li>학년·학기·교과목과 실습 기간(8~10일) 설정</li>
        <li>병원별 정원 직접 입력</li>
        <li>성별, GPA, 출생연도, 과거 실습 기관 조건</li>
        <li>결과 표 정렬·검색, 엑셀 다운로드</li>
        <li>화면에서 이름·학번 일부 가림</li>
      </ul>
      <a class="btn" href="/app">배정 시스템 열기</a>
    </div>
    <div class="lb receipt">
      <h4>점수는 이렇게 계산합니다</h4>
      <div class="total num">70.1<small>점</small></div>
      <div class="stack"><i style="width:57.2%;background:var(--blue)"></i><i style="width:42.8%;background:var(--gold)"></i></div>
      <div class="row"><span><i class="d" style="background:var(--blue)"></i>통학 점수 57.3 × 70%</span><b>40.1</b></div>
      <div class="row"><span><i class="d" style="background:var(--gold)"></i>지망 점수 100 (1지망 일치) × 30%</span><b>30.0</b></div>
      <div class="row" style="border-bottom:0"><b>종합 매칭점수</b><b>70.1</b></div>
      <div class="small">예시 화면입니다. 비중은 슬라이더로 바꿀 수 있습니다.</div>
    </div>
  </div>
</div></section>

<section class="sh on-liner" id="start"><div class="wrap">
  <div class="sec-h"><h2>먼저 비중을 바꿔 보세요</h2><p>슬라이더를 움직여 보세요. 통학 거리를 중요하게 보면 가까운 병원이, 희망을 중요하게 보면 희망 병원이 올라옵니다.</p></div>
  <div class="demo">
    <div class="lb ctl">
      <h3>예시 학생 한 명</h3>
      <p>1지망 고려대학교구로병원, 2지망 서울특별시서남병원을 적은 학생입니다.</p>
      <div class="wl"><span id="dLblD">통학 거리 70%</span><span id="dLblP">희망 병원 30%</span></div>
      <input class="sl" type="range" id="dSlider" min="0" max="100" step="5" value="70" aria-label="통학 거리와 희망 병원의 비중">
      <div class="hint">오른쪽으로 밀수록 통학 거리, 왼쪽으로 밀수록 희망 병원의 비중이 커집니다. 예시 값으로 계산한 화면이며 실제 데이터가 아닙니다.</div>
      <div class="go"><a class="btn" href="/app">배정 시스템 열기</a></div>
    </div>
    <div id="dRows" aria-live="polite"></div>
  </div>
</div></section>

<section class="sh on-liner" id="steps"><div class="wrap">
  <div class="sec-h"><h2>세 단계로 끝납니다</h2><p>조건을 정하고, 명단을 올리고, 추천안을 확인합니다.</p></div>
  <div class="lb strip">
    <article class="stp"><div class="n">1</div><h3>조건 설정</h3><p>배정 방식, 교과목, 실습 기간, 거리와 희망의 비중, 병원별 정원을 정합니다.</p></article>
    <article class="stp"><div class="n">2</div><h3>명단 업로드</h3><p>엑셀이나 CSV 명단을 올립니다. 파일은 서버에 저장하지 않고 처리 후 사라집니다.</p></article>
    <article class="stp"><div class="n">3</div><h3>추천안 확인</h3><p>병원별 현황, 통학시간 분포, 지망 충족률을 보고 학생별 근거를 확인합니다.</p></article>
  </div>
  <div class="steps-cta"><a class="btn" href="/app">배정 시스템 열기</a></div>
</div></section>

<section class="sh on-liner" id="guide"><div class="wrap"><div class="guide-grid">
  <div class="guide-l">
    <div class="sec-h"><h2>처음 쓰는 분을 위한 안내</h2><p>자주 묻는 내용을 모았습니다.</p></div>
    __GUIDE_BTN__
  </div>
  <div class="faq">
    <details><summary>학생 명단은 서버에 저장되나요?</summary><p>저장하지 않습니다. 올린 파일은 메모리에서 처리한 뒤 사라지고, 결과는 엑셀로 내려받아 보관합니다.</p></details>
    <details><summary>명단 파일은 어떤 형식이어야 하나요?</summary><p>엑셀(.xlsx) 또는 CSV입니다. 필수 열은 학번, 이름, 성별, GPA, 출생연도이며, 양식 샘플을 내려받아 확인할 수 있습니다.</p></details>
    <details><summary>통학시간 앞의 '약'은 무슨 뜻인가요?</summary><p>지도 위치를 바탕으로 계산한 예상 소요시간이라는 뜻입니다. '추정'은 주소 정보가 부족해 규칙으로 계산한 값입니다.</p></details>
    <details><summary>지도에 표시되는 거주지는 실제 집인가요?</summary><p>아닙니다. 주소가 동 단위까지만 있어서, 동네 기준의 대략적인 위치로 표시합니다.</p></details>
    <details><summary>배정 결과를 그대로 확정하나요?</summary><p>아닙니다. 시스템은 추천안과 근거를 제공하고, 검토와 최종 결정은 교수자가 합니다.</p></details>
    <details><summary>미배정 학생이 나오면 어떻게 하나요?</summary><p>미배정·사유 탭에서 이유(정원 부족, 조건 불일치 등)를 확인하고, 정원이나 조건을 조정해 다시 배정할 수 있습니다.</p></details>
  </div>
</div></div></section>

<section class="sh on-liner" id="team"><div class="wrap">
  <div class="sec-h"><h2>만든 사람들</h2><p>안산대학교 간호학과 팀 로켓단입니다.</p></div>
  <div class="team">
    <article class="lb mem"><img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACEAAAAzCAYAAAAZ+mH/AAADdklEQVR42tVYTUhUURT+3jRIhGiIiyGcGMbJxkU2WJkKA2KzbCGWBi2GiAhxISaznIWLljLJW8RjaNPbqSktWhREGANqWoPlIsmfEUdjFiIpLkSE12J8z3vf3Pc7byY7m3d/zr33u+d+55z7LucPtsKCSHKhw7NPdUznqsgqZ2VSt1UA5OKemkrkdg8oUCdgJCtAXFYQswB4aio1dUoCghTZAvK3GHHhDMiZAGGFmJjOVSHg3gIA1NVS3oCtnTwXVo/rSgJCIisfIhFd5fr3y+pxhl7CmYgTkh3GE+7KFWsJiXRJPS8xmIMrxhJMK2zt7FOcYAGUwZmxhtuOmdWkzO0eFAArmXfIHsACYxeAZRDqhZICr6n7tG/AeRBJgacmPqyoUcrRgWGlLPLDpbWE1u7tLGw7bCcFHg1dcZy7eqegT6vdERCsndc3trGjJdGeFHhdzli2hNnJ7IphsAKAX2+f25q8oStu6rpnZAlq8OGnRapzLi5gLi5AT8eJ3EHJ+c4QVfcEAvA9iujqOOaifzY3mO0XOwOafU5ywjYvzPLBCISkN3Cef0LVWwZewQq/TB+HMPUOANDkrT6NjqksomEvMPua0p1ZSJ32nciP7B4AoK/7riVOSEbbiYa9EFNZDPl9VHtCBcDAspyuJWYWUgU7YQFJpIAh/3oewLpfFwA5Z/utsDOXGgeEsorbzBHoWUMum5HtTAYT46JS7+mNAoDE+YOt0gVfh+bARCJGEVMX2ItvEJ/dKDjOJm81tjMZAEDztSsAgPTSigLEsT8wMZVVvnKZtAAJYC0cRnX/YwDAxLiY58QlX7BoECFfFUI+7XumDKC+sQ1rP2eV9vTSSh7E741l5sD++yHHGUkCMJVFX75ZxOrMR0cWlznAajN00UB7BFNjkwCA7gf3NPX0yHu9YhPfjy4jvbTCJCbnD7ZKRsFKXmBqbFIXiFrUYXtiXMSX+TRutzSTLgq3KoRKtdufSbpRk1oBwJKe3ihisUFlcTlYlfORhGPUOSYniHuAkkW1ckipXmqoo5HPUgZDyuLGPoShh6b6jTK0y8B8nJM7jsUGMTIy+m+yaIVbnx+ucgBovtkCAOBHR5lH4yqHBdJf5/Ws8n88pkrFTH50bK7N8F+UlfXUf+FGOvIREAC4snuHavecZRdlpWAbOlzRTwMlyBuU/AV+2j+v68a6jwAAAABJRU5ErkJggg==" alt="강지우 캐릭터"><b>강지우</b><span>팀장</span></article>
    <article class="lb mem"><img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABoAAAAvCAYAAAD9/drQAAACqklEQVR42mNU0rBgwAH+M5AHGLEJsuCzhIeFNBu+/IHrxbCMCZcmUi0hpIeFUJAZmdkQtODcqSPYzEDxFSOWOPpPqm+MzGxQLIMGISNRQYcNvHnzAisfi48YiI4jagMmUnwjIiKBkz/0fITuelJ8Q9Cic2taGBgYGBh0pFjgNLIYTBwmhg8QlZAXTWogSoyioDu3poVBLaCGgVndGcOnyOKEfEV0HClrWpIkTlTQERPm+PSpBdRgFunYiiAGBgaGWxvIswzJEoJFEIqCH/suoEieqJnBcKJmBgM+NdiqCYKpjsPJAIUvoaLCoJDgglcNScn7w6MHWMUFnFRwyuGtdnFU5WTFE674wWYR3nbCqUkpKHyzvDlEtx8wgm7Gui0MDAwMDHqy/IhS4PBjhjhbWQaG4wtR1B47fRghBwWXHn9kYGBgYMgI8sGII4KtnThbWYZFhx8zFCkpoIj3oVmCr1nAqKRh8f/Y6cMYLkL2EbLPipTuQSy5pwS3ZN3KtSjqVKxcUPRbmdoysJCTKfdtWcBgwMDAcOkaLCmiBhPDtgKGS5RUfHG2sgx995TIKjFI9lGcrSzDpWsMDK7Fi+FirhiqwuGs3b2x9K3KWSjR3L39Pk65TV9mMRwObaeORaWeirjlGNop85HI04MMTj4JDH9v7sWZMEo9FRmeBwYySK5fT3sfiXXkMPy9uZfBySeBgWFODcQiK1NbjCKIkjiKmVXEINaRgxF0jKR0vt5I2zNcmpnI4Fq8mKFUHYciz/UMf2/upV6qwwdgrSNYPsJpEazMoxbAWR9hi6sLDz4xzCiKYpiRokKSJRlz7mD4iJGYuMqYcwevQ6heMpDSMx983ZYhbdF/Kpj7n6i2993rxwn2KohQwzggQYe923L5NuEuChFqiOq2UDsvAQDmxOC3ExOgygAAAABJRU5ErkJggg==" alt="최영락 캐릭터"><b>최영락</b><span>팀원</span></article>
    <article class="lb mem"><img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAB0AAAAwCAYAAADtoXHnAAACzklEQVR42u1YP2gTURj/XbihFLElZJBohmoqOrRDiLEWskkGJ0nRVodwk2MQ6eTQwcGpIGQqnY4MkjTkJnHQbgdRmnBIExFMA0KkUwmpBHHyHOw73728u3t3SbD++eDg7vu+937fv/e99066eGUJDmQiGEleCrIbYOhb3xfa96lZMtYVOOQo4AD2+j1XUFEjQ0HCqih5KEreTcV1DomTU5O1WFHyUNUCev0ewrNhGx8AVLXAC7MUGJQAkvDSoDydsYDS+eQBuhSU5LuQWBIFDFpIE6fTDdoobQjxhFqWSCGlLkes972PR1xeo7SB5NoTtpBMXouU/Xrng0wAqOwUAQB37uasFhkSAUwoBcxcvzekw/JZIys7RSQW5pFYmLfAAZghUQ8vXb3BNY7lk7EE0DKQAv4t1cstJAAw1HygCRPKz3ZoqHl0ppOWt0azTfLK9dTWvozSS5uwtamhtanBTccK/dcGjGbbBghAkj0tX7tl+z4XjyNye9FVhwUm3nsumaPDz3xBKuwso+h5K4Ts6srJV8EWQUfQzGPNd14pj3ydkWwdpDOdpHOBV0+zQ4bRa5LQ+vpDykuBg9mW9gIAsBibsSYYRFPIpWMYvH5m063VdRT1Ls4c7uH83BxqdR373WMc1HY9G75JPVzKpWMo6l1EohdsT1HvIpeO2XQParuIL9/0Dm+trluM/e6xI/C2DjwIvwUAbPeWhgDHtrVlV1eglauOcq1ctfKnlaueXnruMk7ekvegJPsd4ARGPHZKj29QEmJ2GZCwa+UqBtGUTbZ8LT26pywwAfTIoTRyeOmiEg3lWHLq1mkmBkoT6Vp/xLlXBoCp95VfnLOZv/OEL7P74JaWce3BkziYmewWR9O7T1+w9ei+kNzPOpVG/LMi9Ifl/1Xx3wA1xzCv//9InQ9vnE/sJ7c0AR3p9HUkloxm23OgiI7vq+I4up2T4AfMHhxmQ6yU0AAAAABJRU5ErkJggg==" alt="정현민 캐릭터"><b>정현민</b><span>팀원</span></article>
    <article class="lb mem"><img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABwAAAAxCAYAAADJP8l8AAADCElEQVR42r1YTWgTQRT+VkPpoYdYAoYQ2xKlRrEFMbRpQUELsSBIKihBEIr0HBQ8KJQQtaAHTznvwVMpCpqjFFFRTE31EKyaWJKSlhCWEiSHHIJo1kO66+44M5ndbHyX3Z03M9987837SaRAMAyGqLAnEk/p4oH5oVhCqsCrrWWC7iNAVLtgxBrTXkyGfijaKXWJhI8zAVY/fOOCk3sBgGTwoWqVlXYYFvAeoCTiQyHhMRTxYUcTWmUuBNhrsWzSV2mZqTsXXbBuUrtgjjI8ExpHcjFO1SWXUtbSEC8srDIiTSocFkq9Aa97AAvJFdO4nIzp76Tube4L/Ia1tkwqJ2MYjS4CADbTSyYQmm4rp3R/aUr5NUy7m1SWpM5W4JNybeokpk8ELOtsATZqTbhGQlhvDKJRa+rs5GSMqbMVFuvpCC7OZzDg6WduxNMJM9TKyUR0FQDQytWAaqvzLtVWey5lL14cdozFzYcvTLrR27OWYpAKqJs1xciLn3+1n+P0iJqIy9z+RmI0UaoWX0ap75RN3+6hETPjvbjk9TRMhtTLRLAm2Ah1cC4SpJRfo67MbmwDykvT2PKTFUyODVPnHz42RRKQ9LDgLTRK2ReB3+dpX4pqDWi1uBlKO+jVKzF79dDv86D/xyP9/b9U/ObgLdP3nacF0/eDy0FnKv7k2DDKld2/5q3sCrmha4akdGJERkDXgEY5fXaOOu7zttNcVVGcBQSAeOoxdTwVn7dm0uzGNnWM5cdLfZ/092c/Q/bCwri5nQtjy4eVarsMHTxwiDvPyEooLM7fz1LHE3IGCTnjbCOs+Ut7Gs2nlAvdd96hviJ+fy8CAPYfnQEA3Iv+W7FPzd1FvdAuwhduBpG4Hmbexp4FvlHevX5OVgs6oKFwopSfcfIM1HookenH7pV3pBHu+Q9SWkaxmoEsAWrVOTh7A0cCQyhu7TAXL7+pmvRKuaDfYlFAqcu/vTp2bVwf8tiJ6EUBVQfuhiralwKAymoVjQEtMEcSzjTvP37tSEFkjjDDXmUZAPgDFuQhjuV0nZQAAAAASUVORK5CYII=" alt="강민정 교수님 캐릭터"><b>강민정 교수님</b><span>지도교수 · 간호관리학</span></article>
  </div>
</div></section>

<section class="sh on-blue fin"><div class="wrap">
  <div class="lb final">
    <div><h2>더 좋은 실습지를,<br>더 공정하게</h2><p>통학 거리와 학생의 희망을 함께 계산해 학생마다 더 잘 맞는 실습지를 추천합니다. 교수자는 이유를 확인하고 확정합니다.</p></div>
    <a class="btn" href="/app">배정 시스템 열기</a>
  </div>
</div></section>
</main>

<footer class="foot"><div class="wrap">
  <div class="links"><a href="/app">배정 시스템 열기</a>__GUIDE_FOOT__</div>
  <b>로켓단</b> · AI기반 간호실습 배치 의사결정 프로그램<br>안산대학교 간호학과 캡스톤디자인 작품 · 팀원 강지우(팀장), 최영락, 정현민 · 지도교수 강민정
  <div class="keyhint">키보드 숫자 1~9를 누르면 해당 구역으로 이동합니다.</div>
</div></footer>

<script>
(function () {
  var still = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;

  /* 바코드: 글자에서 만든 고정 무늬 (장식, 읽히는 코드가 아님) */
  function hash(s) { var h = 2166136261; for (var i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619); } return h >>> 0; }
  function rng(a) { return function () { a |= 0; a = a + 0x6D2B79F5 | 0; var t = Math.imul(a ^ a >>> 15, 1 | a); t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t; return ((t ^ t >>> 14) >>> 0) / 4294967296; }; }
  document.querySelectorAll('svg[data-bc]').forEach(function (svg) {
    var r = rng(hash(svg.getAttribute('data-bc'))), x = 0, out = '';
    while (x < 118) { var w = 1 + Math.floor(r() * 3), g = 1 + Math.floor(r() * 2); out += '<rect x="' + x + '" y="0" width="' + w + '" height="26" fill="currentColor"/>'; x += w + g; }
    svg.innerHTML = out;
  });

  /* 비중 체험: 앱과 같은 환산(15분=100점, 90분 이상=0점), 순서가 바뀔 때 자리를 부드럽게 옮김 */
  var H = [
    {n: '고려대학교구로병원', t: 47, p: 100, g: '1지망'},
    {n: '서울특별시서남병원', t: 50, p: 70, g: '2지망'},
    {n: '단원병원', t: 21, p: 10, g: '지망 외'},
    {n: '시화병원', t: 28, p: 10, g: '지망 외'}
  ];
  var sl = document.getElementById('dSlider'), box = document.getElementById('dRows'), els = {};
  H.forEach(function (h) {
    var d = document.createElement('div'); d.className = 'dr';
    d.innerHTML = '<div class="st2"><span class="nm"><span class="t"></span><em class="flag">추천</em></span><span class="sub"></span></div><div class="sc"></div><div class="tr"><i></i></div>';
    els[h.n] = d;
  });
  function draw() {
    var v = +sl.value, w = v / 100;
    sl.style.setProperty('--v', w);
    document.getElementById('dLblD').textContent = '통학 거리 ' + v + '%';
    document.getElementById('dLblP').textContent = '희망 병원 ' + (100 - v) + '%';
    var before = {};
    H.forEach(function (h) { var e = els[h.n]; if (e.parentNode) before[h.n] = e.getBoundingClientRect().top; });
    var r = H.map(function (h) {
      var d = Math.max(0, Math.min(100, 100 - (h.t - 15) / 75 * 100));
      return {n: h.n, t: h.t, g: h.g, s: d * w + h.p * (1 - w)};
    }).sort(function (a, b) { return b.s - a.s; });
    r.forEach(function (x, i) {
      var e = els[x.n];
      e.classList.toggle('top', i === 0);
      e.querySelector('.t').textContent = x.n;
      e.querySelector('.sub').textContent = x.g + ' · 통학 약 ' + x.t + '분';
      e.querySelector('.sc').textContent = x.s.toFixed(1) + '점';
      e.querySelector('.tr i').style.width = Math.max(2, x.s) + '%';
      box.appendChild(e);
    });
    if (!still) r.forEach(function (x) {
      var e = els[x.n];
      if (before[x.n] === undefined) return;
      var dy = before[x.n] - e.getBoundingClientRect().top;
      if (dy) { e.style.transition = 'none'; e.style.transform = 'translateY(' + dy + 'px)'; e.getBoundingClientRect(); e.style.transition = ''; e.style.transform = ''; }
    });
  }
  sl.addEventListener('input', draw); draw();

  /* 숫자 키 1~9: 해당 구역으로 이동 */
  var ORDER = ['intro', 'why', 'compare', 'numbers', 'service', 'start', 'steps', 'guide', 'team'];
  document.addEventListener('keydown', function (e) {
    if (e.ctrlKey || e.metaKey || e.altKey || e.shiftKey) return;
    var t = e.target, tag = t && t.tagName;
    if (tag === 'TEXTAREA' || tag === 'SELECT' || (tag === 'INPUT' && t.type !== 'range')) return;
    var n = parseInt(e.key, 10);
    if (!(n >= 1 && n <= ORDER.length)) return;
    var el = document.getElementById(ORDER[n - 1]);
    if (el) { e.preventDefault(); el.scrollIntoView({behavior: still ? 'auto' : 'smooth', block: 'start'}); }
  });

  /* 오른쪽 점 메뉴: 지금 보는 구역 표시 */
  var dots = document.querySelectorAll('.dots a');
  if (dots.length && 'IntersectionObserver' in window) {
    function mark(id) { dots.forEach(function (a) { a.classList.toggle('on', a.getAttribute('data-s') === id); }); }
    var so = new IntersectionObserver(function (es) {
      es.forEach(function (e) { if (e.isIntersecting) mark(e.target.id); });
    }, {rootMargin: '-45% 0px -50% 0px', threshold: 0});
    dots.forEach(function (a) { var t = document.getElementById(a.getAttribute('data-s')); if (t) so.observe(t); });
    mark('intro');
  }

  /* 시험 결과: 화면에 들어오면 막대가 차오르고 숫자가 올라감(한 번) */
  var rb = document.getElementById('resBox');
  if (!rb) return;
  if (still || !('IntersectionObserver' in window)) { rb.classList.add('go'); return; }
  function run(el) {
    var to = +el.getAttribute('data-n'), dec = +el.getAttribute('data-d') || 0, t0 = null;
    function step(ts) {
      if (t0 === null) t0 = ts;
      var k = Math.min(1, (ts - t0) / 1100), e = 1 - Math.pow(1 - k, 3);
      el.textContent = (to * e).toFixed(dec);
      if (k < 1) requestAnimationFrame(step); else el.textContent = to.toFixed(dec);
    }
    el.textContent = (0).toFixed(dec); requestAnimationFrame(step);
  }
  var io = new IntersectionObserver(function (es) {
    es.forEach(function (e) {
      if (!e.isIntersecting) return;
      rb.classList.add('go');
      rb.querySelectorAll('.cnt').forEach(run);
      io.unobserve(rb);
    });
  }, {threshold: 0.45});
  io.observe(rb);
})();

/* ===== 이동·전환: 진행선, 메뉴 밑줄, 도착 표시 ===== */
(function () {
  var prog = document.querySelector('.prog i'), nav = document.querySelector('.nav');
  var links = nav ? [].slice.call(nav.querySelectorAll('a:not(.btn)')) : [];
  var secs = links.map(function (a) { return document.querySelector(a.getAttribute('href')); });
  var ind = null;
  if (nav) { ind = document.createElement('span'); ind.className = 'nav-ind'; ind.setAttribute('aria-hidden', 'true'); nav.appendChild(ind); }
  var tick = false;
  function update() {
    tick = false;
    var max = document.documentElement.scrollHeight - window.innerHeight;
    if (prog) prog.style.transform = 'scaleX(' + (max > 0 ? Math.min(1, window.scrollY / max) : 0) + ')';
    if (!ind) return;
    var y = window.innerHeight * 0.45, cur = -1;
    secs.forEach(function (s, i) { if (s && s.getBoundingClientRect().top <= y) cur = i; });
    links.forEach(function (a, i) { a.classList.toggle('cur', i === cur); });
    if (cur < 0 || !links[cur].offsetWidth) { ind.classList.remove('on'); return; }
    ind.style.setProperty('--x', links[cur].offsetLeft + 'px');
    ind.style.setProperty('--w', links[cur].offsetWidth + 'px');
    ind.classList.add('on');
  }
  function req() { if (!tick) { tick = true; requestAnimationFrame(update); } }
  window.addEventListener('scroll', req, { passive: true });
  window.addEventListener('resize', req);
  update();
  document.addEventListener('click', function (e) {
    var a = e.target.closest && e.target.closest('a[href^="#"]');
    if (!a) return;
    var t = document.querySelector(a.getAttribute('href'));
    var h = t && t.querySelector('h2');
    if (!h) return;
    setTimeout(function () { h.classList.remove('arrive'); void h.offsetWidth; h.classList.add('arrive'); }, 450);
  });
})();
</script>
</body>
</html>
"""


def _base_url(request: Request) -> str:
    host = request.headers.get("x-forwarded-host") or request.headers.get("host") or ""
    local = host.startswith(("localhost", "127.0.0.1"))
    return f"{'http' if local else 'https'}://{host}"


# ── 글꼴: BM JUA(배달의민족 주아체, SIL Open Font License 1.1). 한글 KS X 1001 + 화면에 쓰인 글자만 담은 woff2 ──
JUA_B64 = "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"

@app.get("/assets/jua.woff2")
async def jua_font():
    return Response(content=base64.b64decode(JUA_B64), media_type="font/woff2", headers={"Cache-Control": "public, max-age=31536000, immutable"})


@app.get("/", response_class=HTMLResponse)
def render_landing(request: Request):
    import html as _html
    guide = os.environ.get("GUIDE_URL", "").strip()
    link = f'<a href="{_html.escape(guide, quote=True)}" target="_blank" rel="noopener">사용설명서</a>' if guide.startswith("http") else ""
    btn = ('<div class="guide-act">' + link.replace("<a ", '<a class="btn ghost" ', 1) + '</div>') if link else ''
    page = LANDING_PAGE.replace("__GUIDE_BTN__", btn).replace("__GUIDE_LINK__", link).replace("__GUIDE_FOOT__", link).replace("__BASE__", _base_url(request)).replace("__APP_VERSION__", APP_VERSION)
    return HTMLResponse(content=page)


@app.get("/app", response_class=HTMLResponse)
def render_ui():
    return HTMLResponse(content=HTML_PAGE.replace("__APP_VERSION__", APP_VERSION))


@app.get("/robots.txt", response_class=PlainTextResponse)
def robots(request: Request):
    return f"User-agent: *\nAllow: /\nDisallow: /api/\nSitemap: {_base_url(request)}/sitemap.xml\n"


@app.get("/sitemap.xml")
def sitemap(request: Request):
    xml = ('<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n'
           f'  <url><loc>{_base_url(request)}/</loc></url>\n</urlset>\n')
    return Response(content=xml, media_type="application/xml")


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
