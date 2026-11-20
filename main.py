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
  f.append('max_period_capacity', $('capPeriod').value || 50); f.append('max_hospital_capacity', $('capHosp').value || 10);
  f.append('gender_criteria', $('gender').value); f.append('exclude_past_hospital', $('exPast').checked);
  f.append('use_transit_api', true);
  f.append('hospital_caps', capsJson());
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
  finally { b.disabled = false; b.textContent = '배정 실행'; }
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
  const pc = [['1지망', '#0f766e'], ['2지망', '#2dd4bf'], ['3지망', '#99f6e4'], ['비지망', '#cbd5e1']];
  const pv = pc.map(([l]) => rs.filter(r => l === '비지망' ? r.preference_status.startsWith('비') : r.preference_status.startsWith(l)).length);
  const tot = pv.reduce((a, b) => a + b, 0); let acc = 0;
  const stops = tot ? pc.map(([, col], i) => { const a = acc / tot * 100; acc += pv[i]; return `${col} ${a}% ${acc / tot * 100}%`; }).join(',') : '#e2e8f0 0 100%';
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
  new N.Polyline({map: m, path: [a, b], strokeColor: '#0f766e', strokeWeight: 4, strokeOpacity: 0.9, strokeStyle: 'shortdash'});
  new N.Marker({position: a, map: m, icon: {content: pin('#d97706', '거주지'), anchor: new N.Point(0, 0)}});
  new N.Marker({position: b, map: m, icon: {content: pin('#1d4ed8', hname), anchor: new N.Point(0, 0)}});
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
  L.polyline([a, b], {color: '#0f766e', weight: 3, dashArray: '6 6'}).addTo(m);
  L.circleMarker(a, {radius: 8, color: '#fff', weight: 2, fillColor: '#d97706', fillOpacity: 1}).addTo(m).bindTooltip('거주지', {permanent: true, direction: 'top'});
  L.circleMarker(b, {radius: 8, color: '#fff', weight: 2, fillColor: '#1d4ed8', fillOpacity: 1}).addTo(m).bindTooltip(esc(hname), {permanent: true, direction: 'top'});
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
    <div class="box" style="background:#f0fdfa"><b style="color:var(--teal-d)">${esc(r.hospital)}</b> · 병원 내 ${esc(r.rank)}순위 · 종합 <b>${esc(r.composite_score)}점</b></div>
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


LANDING_PAGE = """<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>AI기반 간호실습 배치 의사결정 프로그램 | 로켓단</title>
<meta name="description" content="통학 거리와 학생 희망을 함께 계산해 간호학과 실습 병원 배정안을 추천합니다. 최종 결정은 교수자가 하고, 시스템은 근거를 설명합니다.">
<link rel="canonical" href="__BASE__/">
<meta property="og:type" content="website">
<meta property="og:title" content="AI기반 간호실습 배치 의사결정 프로그램 | 로켓단">
<meta property="og:description" content="통학 거리와 학생 희망을 함께 계산해 실습 병원 배정안을 추천합니다.">
<meta property="og:url" content="__BASE__/">
<script type="application/ld+json">{"@context":"https://schema.org","@type":"WebApplication","name":"AI기반 간호실습 배치 의사결정 프로그램 (로켓단)","url":"__BASE__/","applicationCategory":"EducationalApplication","inLanguage":"ko","description":"통학 거리와 학생 희망을 함께 계산해 간호학과 실습 병원 배정안을 추천하는 웹 서비스"}</script>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css" media="print" onload="this.media='all'">
<style>
  :root { --navy:#0f172a; --blue:#1e3a8a; --teal:#0d9488; --teal-d:#0f766e; --teal-l:#ccfbf1; --muted:#475569; --line:#e2e8f0; --bg:#f5f8fc; }
  * { box-sizing:border-box; }
  html { scroll-behavior:smooth; scroll-padding-top:70px; }
  body { margin:0; font-family:'Pretendard',-apple-system,'Noto Sans KR','Malgun Gothic',sans-serif; color:var(--navy); background:#fff; line-height:1.65; }
  body { word-break:keep-all; overflow-wrap:break-word; }
  h1, h2, h3, p, li { text-wrap:balance; }
  .wrap { max-width:1100px; margin:0 auto; padding:0 22px; }
  .notice { background:#f0fdfa; border-bottom:1px solid #d5f3ee; font-size:13px; color:var(--teal-d); text-align:center; padding:8px 12px; }
  header.top { position:sticky; top:0; background:rgba(255,255,255,.95); backdrop-filter:blur(8px); border-bottom:1px solid var(--line); z-index:5; }
  .top-in { display:flex; align-items:center; justify-content:space-between; height:64px; gap:20px; }
  .logo { font-weight:900; font-size:21px; letter-spacing:-.5px; color:var(--blue); text-decoration:none; }
  .nav { display:flex; gap:28px; align-items:center; font-size:15px; font-weight:600; }
  .nav a { color:var(--muted); text-decoration:none; }
  .nav a:hover { color:var(--navy); }
  .btn { display:inline-block; background:var(--teal-d); color:#fff !important; padding:13px 24px; border-radius:10px; font-weight:700; text-decoration:none; font-size:15px; }
  .btn:hover { background:#0b5f58; }
  .btn.sm { padding:9px 18px; font-size:14px; }
  .btn.ghost { background:#fff; color:var(--navy) !important; border:1.5px solid var(--line); }
  .btn.ghost:hover { background:#f8fafc; }
  .hero { background:linear-gradient(135deg,#e6faf6 0,#eef4ff 100%); padding:96px 0 84px; text-align:center; }
  .hero .eyebrow { color:var(--teal-d); font-weight:700; font-size:14px; margin-bottom:14px; }
  .hero h1 { font-size:46px; line-height:1.25; letter-spacing:-1.3px; margin:0 0 18px; }
  .hero p { font-size:18px; color:var(--muted); max-width:820px; margin:0 auto 32px; }
  .hero .cta { display:flex; gap:12px; flex-wrap:wrap; justify-content:center; }
  section { padding:72px 0; }
  section.alt { background:var(--bg); }
  .sec-h { text-align:center; margin-bottom:38px; }
  .sec-h .tag { color:var(--teal-d); font-weight:700; font-size:14px; margin-bottom:6px; }
  h2 { font-size:30px; letter-spacing:-.7px; margin:0 0 10px; }
  .sec-h p { color:var(--muted); margin:0 auto; max-width:760px; }
  .grid { display:grid; gap:16px; }
  .g4 { grid-template-columns:repeat(4,minmax(0,1fr)); }
  .g3 { grid-template-columns:repeat(3,minmax(0,1fr)); }
  .g2 { grid-template-columns:repeat(2,minmax(0,1fr)); }
  .card { background:#fff; border:1px solid var(--line); border-radius:14px; padding:24px; }
  .card h3 { margin:0 0 8px; font-size:17px; }
  .card p { margin:0; color:var(--muted); font-size:15px; }
  .ico { width:40px; height:40px; border-radius:10px; background:var(--teal-l); color:var(--teal-d); font-weight:800; display:flex; align-items:center; justify-content:center; margin-bottom:14px; }
  .stats { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:16px; }
  .stat { text-align:center; background:#fff; border:1px solid var(--line); border-radius:14px; padding:28px 14px; }
  .stat .big { font-size:36px; font-weight:900; color:var(--teal-d); letter-spacing:-1px; }
  .stat .big small { font-size:16px; color:var(--muted); font-weight:700; }
  .stat .from { color:var(--muted); font-size:13px; min-height:20px; }
  .stat .t { font-weight:700; margin-top:6px; font-size:15px; }
  .stat .gain { display:inline-block; margin-top:10px; font-size:13px; font-weight:800; color:var(--teal-d); background:#ccfbf1; padding:3px 10px; border-radius:99px; }
  .note { font-size:13px; color:var(--muted); margin-top:16px; text-align:center; }
  .svc { display:grid; grid-template-columns:1.1fr 1fr; gap:28px; align-items:center; background:#fff; border:1px solid var(--line); border-radius:18px; padding:32px; }
  .svc h3 { font-size:24px; margin:0 0 10px; }
  .svc p { color:var(--muted); margin:0 0 16px; }
  .svc ul { margin:0 0 22px; padding-left:18px; color:var(--muted); font-size:15px; }
  .svc li { margin:4px 0; }
  .score { background:var(--bg); border-radius:14px; padding:22px; }
  .score .bar { display:flex; height:16px; border-radius:8px; overflow:hidden; background:#e2e8f0; margin:12px 0; }
  .score .row { display:flex; justify-content:space-between; font-size:14px; padding:4px 0; }
  .news { background:#fff; border:1px solid var(--line); border-radius:14px; }
  .news li { list-style:none; display:flex; gap:16px; padding:15px 22px; border-bottom:1px solid var(--line); font-size:15px; }
  .news li:last-child { border-bottom:none; }
  .news ul { margin:0; padding:0; }
  .news b { color:var(--teal-d); min-width:50px; }
  .final { background:var(--navy); color:#fff; border-radius:18px; padding:48px 34px; text-align:center; }
  .final h2 { color:#fff; }
  .final p { color:#cbd5e1; margin:0 auto 24px; max-width:620px; }
  footer { padding:34px 0 50px; color:var(--muted); font-size:13px; border-top:1px solid var(--line); }
  footer .links { display:flex; gap:18px; flex-wrap:wrap; margin-bottom:8px; }
  footer a { color:var(--muted); }
  .mock { max-width:880px; margin:48px auto 0; background:#fff; border:1px solid var(--line); border-radius:14px; box-shadow:0 24px 60px rgba(15,23,42,.14); overflow:hidden; text-align:left; }
  .mock-bar { display:flex; align-items:center; gap:7px; padding:11px 14px; background:#f1f5f9; border-bottom:1px solid var(--line); font-size:12px; color:var(--muted); }
  .mock-bar i { width:10px; height:10px; border-radius:50%; background:#cbd5e1; display:block; }
  .mock-bar span { margin-left:10px; }
  .mock-body { padding:20px; }
  .mk-kpis { display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:12px; margin-bottom:16px; }
  .mk-kpi { border:1px solid var(--line); border-left:4px solid var(--teal); border-radius:10px; padding:10px 14px; }
  .mk-kpi small { display:block; color:var(--muted); font-size:12px; }
  .mk-kpi b { font-size:22px; letter-spacing:-.5px; }
  .mk-tbl { width:100%; border-collapse:collapse; font-size:13px; }
  .mk-tbl th { background:#f8fafc; color:#475569; font-size:12px; padding:9px 8px; text-align:center; border-bottom:1px solid var(--line); white-space:nowrap; }
  .mk-tbl td { padding:10px 8px; text-align:center; border-bottom:1px solid #f1f5f9; white-space:nowrap; }
  .mk-tbl td.l, .mk-tbl th.l { text-align:left; }
  .mk-tbl .rk { background:#fef3c7; color:#92400e; font-weight:700; font-size:12px; padding:2px 8px; border-radius:99px; }
  .mk-tbl .hit { background:#ccfbf1; color:#115e59; font-weight:700; font-size:12px; padding:2px 8px; border-radius:99px; }
  .mk-tbl .miss { background:#f1f5f9; color:#475569; font-weight:700; font-size:12px; padding:2px 8px; border-radius:99px; }
  .mk-tbl .sc { color:var(--blue); font-weight:800; }
  .mock-cap { font-size:12px; color:var(--muted); padding:0 20px 14px; }
  .cmp { width:100%; border-collapse:separate; border-spacing:0; background:#fff; border:1px solid var(--line); border-radius:14px; overflow:hidden; }
  .cmp th, .cmp td { padding:16px 20px; text-align:left; font-size:15px; border-bottom:1px solid var(--line); vertical-align:top; }
  .cmp tr:last-child td { border-bottom:none; }
  .cmp thead th { background:#f8fafc; font-size:14px; }
  .cmp thead th.us { background:var(--teal-d); color:#fff; }
  .cmp td:first-child { font-weight:700; width:22%; }
  .cmp td.old { color:var(--muted); }
  .cmp td.new { background:#f0fdfa; color:var(--navy); }
  .demo { display:grid; grid-template-columns:1fr 1.2fr; gap:28px; background:#fff; border:1px solid var(--line); border-radius:18px; padding:30px; align-items:start; }
  .demo h3 { font-size:20px; margin:0 0 8px; }
  .demo p { color:var(--muted); margin:0 0 18px; font-size:15px; }
  .demo .wl { display:flex; justify-content:space-between; font-size:14px; font-weight:700; margin-bottom:8px; }
  .demo input[type=range] { width:100%; accent-color:var(--teal-d); }
  .demo .hint { font-size:12px; color:var(--muted); margin-top:10px; }
  .dr { display:grid; grid-template-columns:1fr auto; gap:4px 10px; padding:11px 12px; border-radius:10px; border:1px solid var(--line); margin-bottom:8px; background:#fff; transition:background .2s, border-color .2s; }
  .dr.top { background:#f0fdfa; border-color:var(--teal); }
  .dr .nm { font-weight:700; font-size:15px; }
  .dr .nm em { font-style:normal; font-size:11px; background:var(--teal-d); color:#fff; padding:1px 7px; border-radius:99px; margin-left:6px; vertical-align:middle; }
  .dr .sub { font-size:12px; color:var(--muted); }
  .dr .sc { font-weight:800; font-size:17px; color:var(--blue); text-align:right; }
  .dr .tr { grid-column:1 / 3; height:7px; border-radius:4px; background:#e2e8f0; overflow:hidden; }
  .dr .tr i { display:block; height:100%; background:var(--teal); transition:width .25s; }
  .faq { max-width:780px; margin:0 auto; }
  .faq details { background:#fff; border:1px solid var(--line); border-radius:12px; margin-bottom:10px; }
  .faq summary { cursor:pointer; padding:16px 20px; font-weight:700; font-size:16px; list-style:none; }
  .faq summary::-webkit-details-marker { display:none; }
  .faq summary::after { content:'＋'; float:right; color:var(--muted); }
  .faq details[open] summary::after { content:'－'; }
  .faq p { margin:0; padding:0 20px 18px; color:var(--muted); font-size:15px; }
  .dots { position:fixed; right:18px; top:50%; transform:translateY(-50%); z-index:4; display:none; flex-direction:column; gap:4px; }
  .dots a { position:relative; display:block; width:26px; height:22px; }
  .dots a::before { content:''; position:absolute; right:8px; top:7px; width:8px; height:8px; border-radius:50%; background:#cbd5e1; transition:all .2s; }
  .dots a:hover::before { background:var(--teal); }
  .dots a.on::before { background:var(--teal-d); width:10px; height:10px; right:7px; top:6px; box-shadow:0 0 0 4px rgba(13,148,136,.18); }
  .dots a span { position:absolute; right:30px; top:1px; white-space:nowrap; background:var(--navy); color:#fff; font-size:12px; font-weight:600; padding:3px 10px; border-radius:6px; opacity:0; pointer-events:none; transition:opacity .15s; }
  .dots a:hover span, .dots a:focus-visible span { opacity:1; }
  .dots kbd { font:700 11px/1 inherit; background:rgba(255,255,255,.22); border-radius:4px; padding:2px 6px; margin-left:8px; }
  .team { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:16px; }
  .mem { text-align:center; padding:26px 10px 20px; border:1px solid var(--line); border-radius:14px; background:#fff; }
  .mem img { height:120px; image-rendering:pixelated; image-rendering:crisp-edges; }
  .mem b { display:block; margin-top:12px; font-size:16px; }
  .mem span { font-size:13.5px; color:var(--muted); }
  @media (max-width:860px) { .team { grid-template-columns:repeat(2,minmax(0,1fr)); } }
  .logo-wrap { display:flex; align-items:center; gap:10px; }
  .ver { font-size:12px; font-weight:700; color:var(--teal-d); background:var(--teal-l); padding:2px 8px; border-radius:99px; }
  .keyhint { font-size:12px; color:var(--muted); margin-top:6px; }
  @media (min-width:1100px) { .dots { display:flex; } }
  /* 구역마다 화면 한 장(헤더 아래 전체 높이)에 맞춤 */
  html { scroll-snap-type:y proximity; }
  main > section, main > .hero, main > section.fin { min-height:calc(100vh - 70px); min-height:calc(100svh - 70px); display:flex; flex-direction:column; justify-content:center; scroll-snap-align:start; }
  main > section > .wrap, main > .hero > .wrap { width:100%; }
  main > .hero { padding-top:40px; padding-bottom:40px; }
  main > section { padding-top:40px; padding-bottom:40px; }
  @media (max-width:860px) {
    .demo { grid-template-columns:1fr; }
    .g4, .stats { grid-template-columns:repeat(2,minmax(0,1fr)); }
    .svc { grid-template-columns:1fr; }
  }
  @media (max-width:640px) {
    .mk-kpis { grid-template-columns:1fr; } .mock-body { padding:12px; overflow-x:auto; } .mk-tbl { font-size:12px; } .mk-tbl .mh, .mk-tbl th:nth-child(3), .mk-tbl td:nth-child(3) { display:none; } .mk-tbl td, .mk-tbl th { padding:8px 4px; } .cmp th, .cmp td { padding:12px; font-size:14px; }
    .hero { padding:56px 0 48px; } .hero h1 { font-size:31px; } .hero p { font-size:16px; }
    .g3, .g2, .g4 { grid-template-columns:1fr; } .nav a.hide { display:none; } section { padding:48px 0; }
    .nav { gap:14px; } h2 { font-size:25px; } .stat .big { font-size:30px; }
  }
</style>
</head>
<body>
<div class="notice">캡스톤디자인 작품 · 학생 명단 파일은 서버에 저장되지 않습니다</div>
<header class="top"><div class="wrap top-in">
  <div class="logo-wrap"><a class="logo" href="/">로켓단</a><span class="ver">__APP_VERSION__</span></div>
  <nav class="nav"><a class="hide" href="#intro">소개</a><a class="hide" href="#steps">순서</a><a class="hide" href="#guide">자주 묻는 질문</a><a class="hide" href="#team">만든 사람들</a><a class="btn sm" href="/app">배정 시스템 열기</a></nav>
</div></header>

<main>
<nav class="dots" aria-label="구역 이동"><a href="#intro" data-s="intro" title="소개" aria-label="소개"><span>소개<kbd>1</kbd></span></a><a href="#why" data-s="why" title="왜 필요한가" aria-label="왜 필요한가"><span>왜 필요한가<kbd>2</kbd></span></a><a href="#compare" data-s="compare" title="비교" aria-label="비교"><span>비교<kbd>3</kbd></span></a><a href="#numbers" data-s="numbers" title="시험 결과" aria-label="시험 결과"><span>시험 결과<kbd>4</kbd></span></a><a href="#service" data-s="service" title="주요 기능" aria-label="주요 기능"><span>주요 기능<kbd>5</kbd></span></a><a href="#start" data-s="start" title="직접 해보기" aria-label="직접 해보기"><span>직접 해보기<kbd>6</kbd></span></a><a href="#steps" data-s="steps" title="순서" aria-label="순서"><span>순서<kbd>7</kbd></span></a><a href="#guide" data-s="guide" title="자주 묻는 질문" aria-label="자주 묻는 질문"><span>자주 묻는 질문<kbd>8</kbd></span></a><a href="#team" data-s="team" title="만든 사람들" aria-label="만든 사람들"><span>만든 사람들<kbd>9</kbd></span></a></nav>
<div class="hero" id="intro"><div class="wrap">
  <div class="eyebrow">AI기반 간호실습 배치 의사결정 프로그램</div>
  <h1>실습 배정의 모든 과정,<br>거리와 희망으로 스마트하게</h1>
  <p>통학 거리와 희망 병원을 함께 계산해 배정안을 추천하고, 근거까지 보여 줍니다.<br>최종 결정은 교수자가 합니다.</p>
  <div class="cta"><a class="btn" href="/app">배정 시스템 열기</a><a class="btn ghost" href="#service">서비스 알아보기</a></div>
  <div class="mock">
    <div class="mock-bar"><i></i><i></i><i></i><span>로켓단 · 결과 확인</span></div>
    <div class="mock-body">
      <div class="mk-kpis">
        <div class="mk-kpi"><small>시험한 학생 수</small><b>221명</b></div>
        <div class="mk-kpi"><small>평균 통학시간</small><b>35분</b></div>
        <div class="mk-kpi"><small>지망 충족률</small><b>76%</b></div>
      </div>
      <table class="mk-tbl"><thead><tr><th class="l">배정 병원</th><th class="mh">학번</th><th>이름</th><th>지망</th><th>통학시간</th><th>종합점수</th></tr></thead><tbody>
        <tr><td class="l">고려대학교구로병원</td><td class="mh">202○○○07</td><td><b>김○수</b></td><td><span class="hit">1지망</span></td><td><b>약 24분</b></td><td class="sc">91.6</td></tr>
        <tr><td class="l">순천향대학교부천병원</td><td class="mh">202○○○15</td><td><b>이○은</b></td><td><span class="hit">1지망</span></td><td><b>약 18분</b></td><td class="sc">97.2</td></tr>
        <tr><td class="l">단원병원</td><td class="mh">202○○○31</td><td><b>박○준</b></td><td><span class="hit">2지망</span></td><td><b>약 29분</b></td><td class="sc">77.9</td></tr>
        <tr><td class="l">아주대학교병원</td><td class="mh">202○○○42</td><td><b>최○린</b></td><td><span class="miss">지망 외</span></td><td><b>약 21분</b></td><td class="sc">67.4</td></tr>
      </tbody></table>
    </div>
    <div class="mock-cap">화면 예시입니다. 가상 학생 데이터이고, 이름과 학번은 일부를 가렸습니다.</div>
  </div>
</div></div>

<section id="why"><div class="wrap">
  <div class="sec-h"><div class="tag">Why 로켓단?</div><h2>배정 업무를 더 합리적으로</h2><p>엑셀로 손수 맞추던 실습 배정을, 설명할 수 있는 추천안으로 바꿉니다.</p></div>
  <div class="grid g4">
    <div class="card"><div class="ico">거리</div><h3>통학 부담 반영</h3><p>지도 기반 거리와 소요시간으로, 멀리 가는 학생을 줄이는 방향으로 계산합니다.</p></div>
    <div class="card"><div class="ico">희망</div><h3>학생 희망 반영</h3><p>1~3지망을 점수에 넣고, 거리와 희망의 비중은 슬라이더로 조절합니다.</p></div>
    <div class="card"><div class="ico">근거</div><h3>근거를 보여 주는 결과</h3><p>학생마다 점수 계산 내역과 다른 병원이었을 때의 점수, 위치 지도를 확인할 수 있습니다.</p></div>
    <div class="card"><div class="ico">결정</div><h3>최종 결정은 교수자</h3><p>시스템은 추천안을 제공하고, 검토와 확정은 교수자가 합니다.</p></div>
  </div>
</div></section>

<section id="compare" class="alt"><div class="wrap">
  <div class="sec-h"><div class="tag">비교</div><h2>엑셀 수작업과 무엇이 다른가</h2><p>최종 결정은 그대로 교수자가 합니다. 계산과 근거 정리를 시스템이 돕습니다.</p></div>
  <table class="cmp"><thead><tr><th></th><th>엑셀 수작업</th><th class="us">로켓단</th></tr></thead><tbody>
    <tr><td>통학 거리</td><td class="old">학생 주소를 보고 사람이 판단</td><td class="new">지도 기반 거리와 소요시간을 자동 계산</td></tr>
    <tr><td>학생 희망</td><td class="old">1~3지망을 보며 하나씩 조정</td><td class="new">지망을 점수에 반영해 한 번에 계산</td></tr>
    <tr><td>조건이 바뀔 때</td><td class="old">정원이나 기준이 바뀌면 다시 맞춤</td><td class="new">조건을 바꾸고 다시 계산하면 바로 새 추천안</td></tr>
    <tr><td>결과 설명</td><td class="old">필요할 때 따로 정리</td><td class="new">학생별 점수 근거와 다른 병원 점수를 자동 제공</td></tr>
    <tr><td>최종 결정</td><td class="old">교수자</td><td class="new">교수자 (시스템은 추천안만 제공)</td></tr>
  </tbody></table>
</div></section>

<section id="numbers"><div class="wrap">
  <div class="sec-h"><div class="tag">시험 결과</div><h2>숫자로 보는 로켓단</h2><p>가상 학생 221명 · 4학년 1학기 관리 실습 9개 병원으로, 정원만 지킨 무작위 배정 100회 평균과 비교했습니다.</p></div>
  <div class="stats">
    <div class="stat"><div class="from">무작위 배정 56.7분</div><div class="big"><span class="cnt" data-n="35">35</span><small> 분</small></div><div class="t">평균 통학시간</div><div class="gain">21.7분 단축 (−38%)</div></div>
    <div class="stat"><div class="from">무작위 배정 19.4%</div><div class="big"><span class="cnt" data-n="76">76</span><small> %</small></div><div class="t">지망 충족률</div><div class="gain">+56.6%p</div></div>
    <div class="stat"><div class="from">무작위 배정 41.8점</div><div class="big"><span class="cnt" data-n="75.6">75.6</span><small> 점</small></div><div class="t">평균 매칭점수</div><div class="gain">+33.8점</div></div>
    <div class="stat"><div class="from">221명 전원 배정</div><div class="big"><span class="cnt" data-n="0">0</span><small> 명</small></div><div class="t">미배정 학생</div><div class="gain">병원 정원 100% 준수</div></div>
  </div>
  <div class="note">가상 명단 기준 시뮬레이션 결과이며, 통학시간은 지도 기반 예상값입니다. 실제 학과의 배정 결과와 비교한 값이 아닙니다.</div>
</div></section>

<section id="service" class="alt"><div class="wrap">
  <div class="sec-h"><div class="tag">주요 기능</div><h2>시스템이 계산하고, 교수자가 결정합니다</h2><p>최적화 알고리즘(헝가리안 알고리즘)이 병원 정원 안에서 학생 전체의 점수 합이 가장 높은 조합을 찾습니다.</p></div>
  <div class="svc">
    <div>
      <h3>조건을 정하고 명단을 올리면, 추천안이 나옵니다</h3>
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
    <div class="score">
      <b>점수는 이렇게 계산합니다</b>
      <div class="bar"><i style="width:40%;background:#0d9488"></i><i style="width:30%;background:#f59e0b"></i></div>
      <div class="row"><span>통학 점수 57.3 × 70%</span><b>40.1</b></div>
      <div class="row"><span>지망 점수 100 (1지망 일치) × 30%</span><b>30.0</b></div>
      <div class="row" style="border-top:1px solid var(--line);margin-top:6px;padding-top:8px"><b>종합 매칭점수</b><b>70.1</b></div>
      <div style="font-size:13px;color:var(--muted);margin-top:8px">예시 화면입니다. 비중은 슬라이더로 바꿀 수 있습니다.</div>
    </div>
  </div>
  </div></section>

<section id="start"><div class="wrap">
  <div class="sec-h"><div class="tag">직접 해보기</div><h2>먼저 비중을 바꿔 보세요</h2><p>슬라이더를 움직여 보세요. 통학 거리를 중요하게 보면 가까운 병원이, 희망을 중요하게 보면 희망 병원이 올라옵니다.</p></div>
  <div class="demo">
    <div>
      <h3>예시 학생 한 명</h3>
      <p>1지망 고려대학교구로병원, 2지망 서울특별시서남병원을 적은 학생입니다.</p>
      <div class="wl"><span id="dLblP">희망 병원 30%</span><span id="dLblD">통학 거리 70%</span></div>
      <input type="range" id="dSlider" min="0" max="100" step="5" value="70" aria-label="통학 거리와 희망 병원의 비중">
      <div class="hint">왼쪽은 희망 병원, 오른쪽은 통학 거리의 비중이 커집니다. 예시 값으로 계산한 화면이며 실제 데이터가 아닙니다.</div>
    </div>
    <div id="dRows"></div>
  </div>
  <div style="text-align:center;margin-top:28px"><a class="btn" href="/app" style="padding:16px 34px;font-size:17px">배정 시스템 열기</a></div>
</div></section>

<section id="steps" class="alt"><div class="wrap">
  <div class="sec-h"><div class="tag">순서</div><h2>세 단계로 끝납니다</h2><p>조건을 정하고, 명단을 올리고, 추천안을 확인합니다.</p></div>
  <div class="grid g3" style="margin-top:16px">
    <div class="card"><div class="ico">1</div><h3>조건 설정</h3><p>배정 방식, 교과목, 실습 기간, 거리와 희망의 비중, 병원별 정원을 정합니다.</p></div>
    <div class="card"><div class="ico">2</div><h3>명단 업로드</h3><p>엑셀이나 CSV 명단을 올립니다. 파일은 서버에 저장하지 않고 처리 후 사라집니다.</p></div>
    <div class="card"><div class="ico">3</div><h3>추천안 확인</h3><p>병원별 현황, 통학시간 분포, 지망 충족률을 보고 학생별 근거를 확인합니다.</p></div>
  </div>

  <div style="text-align:center;margin-top:28px"><a class="btn" href="/app">배정 시스템 열기</a></div>
</div></section>

<section id="guide"><div class="wrap">
  <div class="sec-h"><div class="tag">사용법</div><h2>처음 쓰는 분을 위한 안내</h2><p>자주 묻는 내용을 모았습니다.</p></div>
  __GUIDE_BTN__
  <div class="faq">
    <details><summary>학생 명단은 서버에 저장되나요?</summary><p>저장하지 않습니다. 올린 파일은 메모리에서 처리한 뒤 사라지고, 결과는 엑셀로 내려받아 보관합니다.</p></details>
    <details><summary>명단 파일은 어떤 형식이어야 하나요?</summary><p>엑셀(.xlsx) 또는 CSV입니다. 필수 열은 학번, 이름, 성별, GPA, 출생연도이며, 양식 샘플을 내려받아 확인할 수 있습니다.</p></details>
    <details><summary>통학시간 앞의 '약'은 무슨 뜻인가요?</summary><p>지도 위치를 바탕으로 계산한 예상 소요시간이라는 뜻입니다. '추정'은 주소 정보가 부족해 규칙으로 계산한 값입니다.</p></details>
    <details><summary>지도에 표시되는 거주지는 실제 집인가요?</summary><p>아닙니다. 주소가 동 단위까지만 있어서, 동네 기준의 대략적인 위치로 표시합니다.</p></details>
    <details><summary>배정 결과를 그대로 확정하나요?</summary><p>아닙니다. 시스템은 추천안과 근거를 제공하고, 검토와 최종 결정은 교수자가 합니다.</p></details>
    <details><summary>미배정 학생이 나오면 어떻게 하나요?</summary><p>미배정·사유 탭에서 이유(정원 부족, 조건 불일치 등)를 확인하고, 정원이나 조건을 조정해 다시 배정할 수 있습니다.</p></details>
  </div>
</div></section>

<section id="team" class="alt"><div class="wrap">
  <div class="sec-h"><div class="tag">팀 로켓단</div><h2>만든 사람들</h2><p>안산대학교 간호학과 팀 로켓단입니다.</p></div>
  <div class="team">
    <div class="mem"><img class="px" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACEAAAAzCAYAAAAZ+mH/AAADdklEQVR42tVYTUhUURT+3jRIhGiIiyGcGMbJxkU2WJkKA2KzbCGWBi2GiAhxISaznIWLljLJW8RjaNPbqSktWhREGANqWoPlIsmfEUdjFiIpLkSE12J8z3vf3Pc7byY7m3d/zr33u+d+55z7LucPtsKCSHKhw7NPdUznqsgqZ2VSt1UA5OKemkrkdg8oUCdgJCtAXFYQswB4aio1dUoCghTZAvK3GHHhDMiZAGGFmJjOVSHg3gIA1NVS3oCtnTwXVo/rSgJCIisfIhFd5fr3y+pxhl7CmYgTkh3GE+7KFWsJiXRJPS8xmIMrxhJMK2zt7FOcYAGUwZmxhtuOmdWkzO0eFAArmXfIHsACYxeAZRDqhZICr6n7tG/AeRBJgacmPqyoUcrRgWGlLPLDpbWE1u7tLGw7bCcFHg1dcZy7eqegT6vdERCsndc3trGjJdGeFHhdzli2hNnJ7IphsAKAX2+f25q8oStu6rpnZAlq8OGnRapzLi5gLi5AT8eJ3EHJ+c4QVfcEAvA9iujqOOaifzY3mO0XOwOafU5ywjYvzPLBCISkN3Cef0LVWwZewQq/TB+HMPUOANDkrT6NjqksomEvMPua0p1ZSJ32nciP7B4AoK/7riVOSEbbiYa9EFNZDPl9VHtCBcDAspyuJWYWUgU7YQFJpIAh/3oewLpfFwA5Z/utsDOXGgeEsorbzBHoWUMum5HtTAYT46JS7+mNAoDE+YOt0gVfh+bARCJGEVMX2ItvEJ/dKDjOJm81tjMZAEDztSsAgPTSigLEsT8wMZVVvnKZtAAJYC0cRnX/YwDAxLiY58QlX7BoECFfFUI+7XumDKC+sQ1rP2eV9vTSSh7E741l5sD++yHHGUkCMJVFX75ZxOrMR0cWlznAajN00UB7BFNjkwCA7gf3NPX0yHu9YhPfjy4jvbTCJCbnD7ZKRsFKXmBqbFIXiFrUYXtiXMSX+TRutzSTLgq3KoRKtdufSbpRk1oBwJKe3ihisUFlcTlYlfORhGPUOSYniHuAkkW1ckipXmqoo5HPUgZDyuLGPoShh6b6jTK0y8B8nJM7jsUGMTIy+m+yaIVbnx+ucgBovtkCAOBHR5lH4yqHBdJf5/Ws8n88pkrFTH50bK7N8F+UlfXUf+FGOvIREAC4snuHavecZRdlpWAbOlzRTwMlyBuU/AV+2j+v68a6jwAAAABJRU5ErkJggg==" alt="강지우 캐릭터"><b>강지우</b><span>팀장</span></div>
    <div class="mem"><img class="px" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABoAAAAvCAYAAAD9/drQAAACqklEQVR42mNU0rBgwAH+M5AHGLEJsuCzhIeFNBu+/IHrxbCMCZcmUi0hpIeFUJAZmdkQtODcqSPYzEDxFSOWOPpPqm+MzGxQLIMGISNRQYcNvHnzAisfi48YiI4jagMmUnwjIiKBkz/0fITuelJ8Q9Cic2taGBgYGBh0pFjgNLIYTBwmhg8QlZAXTWogSoyioDu3poVBLaCGgVndGcOnyOKEfEV0HClrWpIkTlTQERPm+PSpBdRgFunYiiAGBgaGWxvIswzJEoJFEIqCH/suoEieqJnBcKJmBgM+NdiqCYKpjsPJAIUvoaLCoJDgglcNScn7w6MHWMUFnFRwyuGtdnFU5WTFE674wWYR3nbCqUkpKHyzvDlEtx8wgm7Gui0MDAwMDHqy/IhS4PBjhjhbWQaG4wtR1B47fRghBwWXHn9kYGBgYMgI8sGII4KtnThbWYZFhx8zFCkpoIj3oVmCr1nAqKRh8f/Y6cMYLkL2EbLPipTuQSy5pwS3ZN3KtSjqVKxcUPRbmdoysJCTKfdtWcBgwMDAcOkaLCmiBhPDtgKGS5RUfHG2sgx995TIKjFI9lGcrSzDpWsMDK7Fi+FirhiqwuGs3b2x9K3KWSjR3L39Pk65TV9mMRwObaeORaWeirjlGNop85HI04MMTj4JDH9v7sWZMEo9FRmeBwYySK5fT3sfiXXkMPy9uZfBySeBgWFODcQiK1NbjCKIkjiKmVXEINaRgxF0jKR0vt5I2zNcmpnI4Fq8mKFUHYciz/UMf2/upV6qwwdgrSNYPsJpEazMoxbAWR9hi6sLDz4xzCiKYpiRokKSJRlz7mD4iJGYuMqYcwevQ6heMpDSMx983ZYhbdF/Kpj7n6i2993rxwn2KohQwzggQYe923L5NuEuChFqiOq2UDsvAQDmxOC3ExOgygAAAABJRU5ErkJggg==" alt="최영락 캐릭터"><b>최영락</b><span>팀원</span></div>
    <div class="mem"><img class="px" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAB0AAAAwCAYAAADtoXHnAAACzklEQVR42u1YP2gTURj/XbihFLElZJBohmoqOrRDiLEWskkGJ0nRVodwk2MQ6eTQwcGpIGQqnY4MkjTkJnHQbgdRmnBIExFMA0KkUwmpBHHyHOw73728u3t3SbD++eDg7vu+937fv/e99066eGUJDmQiGEleCrIbYOhb3xfa96lZMtYVOOQo4AD2+j1XUFEjQ0HCqih5KEreTcV1DomTU5O1WFHyUNUCev0ewrNhGx8AVLXAC7MUGJQAkvDSoDydsYDS+eQBuhSU5LuQWBIFDFpIE6fTDdoobQjxhFqWSCGlLkes972PR1xeo7SB5NoTtpBMXouU/Xrng0wAqOwUAQB37uasFhkSAUwoBcxcvzekw/JZIys7RSQW5pFYmLfAAZghUQ8vXb3BNY7lk7EE0DKQAv4t1cstJAAw1HygCRPKz3ZoqHl0ppOWt0azTfLK9dTWvozSS5uwtamhtanBTccK/dcGjGbbBghAkj0tX7tl+z4XjyNye9FVhwUm3nsumaPDz3xBKuwso+h5K4Ts6srJV8EWQUfQzGPNd14pj3ydkWwdpDOdpHOBV0+zQ4bRa5LQ+vpDykuBg9mW9gIAsBibsSYYRFPIpWMYvH5m063VdRT1Ls4c7uH83BxqdR373WMc1HY9G75JPVzKpWMo6l1EohdsT1HvIpeO2XQParuIL9/0Dm+trluM/e6xI/C2DjwIvwUAbPeWhgDHtrVlV1eglauOcq1ctfKnlaueXnruMk7ekvegJPsd4ARGPHZKj29QEmJ2GZCwa+UqBtGUTbZ8LT26pywwAfTIoTRyeOmiEg3lWHLq1mkmBkoT6Vp/xLlXBoCp95VfnLOZv/OEL7P74JaWce3BkziYmewWR9O7T1+w9ei+kNzPOpVG/LMi9Ifl/1Xx3wA1xzCv//9InQ9vnE/sJ7c0AR3p9HUkloxm23OgiI7vq+I4up2T4AfMHhxmQ6yU0AAAAABJRU5ErkJggg==" alt="정현민 캐릭터"><b>정현민</b><span>팀원</span></div>
    <div class="mem"><img class="px" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABwAAAAxCAYAAADJP8l8AAADCElEQVR42r1YTWgTQRT+VkPpoYdYAoYQ2xKlRrEFMbRpQUELsSBIKihBEIr0HBQ8KJQQtaAHTznvwVMpCpqjFFFRTE31EKyaWJKSlhCWEiSHHIJo1kO66+44M5ndbHyX3Z03M9987837SaRAMAyGqLAnEk/p4oH5oVhCqsCrrWWC7iNAVLtgxBrTXkyGfijaKXWJhI8zAVY/fOOCk3sBgGTwoWqVlXYYFvAeoCTiQyHhMRTxYUcTWmUuBNhrsWzSV2mZqTsXXbBuUrtgjjI8ExpHcjFO1SWXUtbSEC8srDIiTSocFkq9Aa97AAvJFdO4nIzp76Tube4L/Ia1tkwqJ2MYjS4CADbTSyYQmm4rp3R/aUr5NUy7m1SWpM5W4JNybeokpk8ELOtsATZqTbhGQlhvDKJRa+rs5GSMqbMVFuvpCC7OZzDg6WduxNMJM9TKyUR0FQDQytWAaqvzLtVWey5lL14cdozFzYcvTLrR27OWYpAKqJs1xciLn3+1n+P0iJqIy9z+RmI0UaoWX0ap75RN3+6hETPjvbjk9TRMhtTLRLAm2Ah1cC4SpJRfo67MbmwDykvT2PKTFUyODVPnHz42RRKQ9LDgLTRK2ReB3+dpX4pqDWi1uBlKO+jVKzF79dDv86D/xyP9/b9U/ObgLdP3nacF0/eDy0FnKv7k2DDKld2/5q3sCrmha4akdGJERkDXgEY5fXaOOu7zttNcVVGcBQSAeOoxdTwVn7dm0uzGNnWM5cdLfZ/092c/Q/bCwri5nQtjy4eVarsMHTxwiDvPyEooLM7fz1LHE3IGCTnjbCOs+Ut7Gs2nlAvdd96hviJ+fy8CAPYfnQEA3Iv+W7FPzd1FvdAuwhduBpG4Hmbexp4FvlHevX5OVgs6oKFwopSfcfIM1HookenH7pV3pBHu+Q9SWkaxmoEsAWrVOTh7A0cCQyhu7TAXL7+pmvRKuaDfYlFAqcu/vTp2bVwf8tiJ6EUBVQfuhiralwKAymoVjQEtMEcSzjTvP37tSEFkjjDDXmUZAPgDFuQhjuV0nZQAAAAASUVORK5CYII=" alt="강민정 교수님 캐릭터"><b>강민정 교수님</b><span>지도교수 · 간호관리학</span></div>
  </div>
</div></section>

<section class="fin"><div class="wrap">
  <div class="final">
    <h2>결정은 교수자가, 근거는 시스템이</h2>
    <p>로켓단은 배정을 대신 결정하지 않습니다. 더 합리적으로 선택할 수 있도록 추천안과 그 이유를 제공합니다.</p>
    <a class="btn" href="/app">배정 시스템 열기</a>
  </div>
</div></section>
</main>

<footer><div class="wrap">
  <div class="links"><a href="/app">배정 시스템 열기</a>__GUIDE_FOOT__</div>
  <b style="color:var(--navy)">로켓단</b> · AI기반 간호실습 배치 의사결정 프로그램<br>안산대학교 간호학과 캡스톤디자인 작품 · 팀원 강지우(팀장), 최영락, 정현민 · 지도교수 강민정
  <div class="keyhint">키보드 숫자 1~9를 누르면 해당 구역으로 이동합니다.</div>
</div></footer>
<script>
(function () {
  /* 비중 체험: 앱과 같은 환산(15분=100점, 90분 이상=0점) */
  var H = [
    {n: '고려대학교구로병원', t: 47, p: 100, g: '1지망'},
    {n: '서울특별시서남병원', t: 50, p: 70, g: '2지망'},
    {n: '단원병원', t: 21, p: 10, g: '지망 외'},
    {n: '시화병원', t: 28, p: 10, g: '지망 외'}
  ];
  var sl = document.getElementById('dSlider'), box = document.getElementById('dRows');
  function esc(x) { return String(x).replace(/[&<>"']/g, function (c) { return {'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'}[c]; }); }
  function draw() {
    var w = +sl.value / 100;
    document.getElementById('dLblD').textContent = '통학 거리 ' + sl.value + '%';
    document.getElementById('dLblP').textContent = '희망 병원 ' + (100 - sl.value) + '%';
    var r = H.map(function (h) {
      var d = Math.max(0, Math.min(100, 100 - (h.t - 15) / 75 * 100));
      return {n: h.n, t: h.t, g: h.g, s: d * w + h.p * (1 - w)};
    }).sort(function (a, b) { return b.s - a.s; });
    box.innerHTML = r.map(function (x, i) {
      return '<div class="dr' + (i === 0 ? ' top' : '') + '"><div><span class="nm">' + esc(x.n) + (i === 0 ? '<em>추천</em>' : '') + '</span><div class="sub">' + esc(x.g) + ' · 통학 약 ' + x.t + '분</div></div><div class="sc">' + x.s.toFixed(1) + '점</div><div class="tr"><i style="width:' + Math.max(2, x.s) + '%"></i></div></div>';
    }).join('');
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
    if (el) { e.preventDefault(); el.scrollIntoView({behavior: 'smooth', block: 'start'}); }
  });

  /* 오른쪽 점 메뉴: 지금 보는 구역 표시 */
  var dots = document.querySelectorAll('.dots a');
  if (dots.length && 'IntersectionObserver' in window) {
    var cur = null;
    function mark(id) { dots.forEach(function (a) { a.classList.toggle('on', a.getAttribute('data-s') === id); }); }
    var so = new IntersectionObserver(function (es) {
      es.forEach(function (e) { if (e.isIntersecting) { cur = e.target.id; mark(cur); } });
    }, {rootMargin: '-45% 0px -50% 0px', threshold: 0});
    dots.forEach(function (a) { var t = document.getElementById(a.getAttribute('data-s')); if (t) so.observe(t); });
    mark('intro');
  }

  /* 숫자 올라가는 효과 */
  var els = document.querySelectorAll('.cnt');
  var still = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  if (still || !('IntersectionObserver' in window)) return;
  function run(el) {
    var to = +el.getAttribute('data-n'), t0 = null;
    function step(ts) {
      if (t0 === null) t0 = ts;
      var k = Math.min(1, (ts - t0) / 1100), e = 1 - Math.pow(1 - k, 3);
      el.textContent = Math.round(to * e);
      if (k < 1) requestAnimationFrame(step); else el.textContent = to;
    }
    el.textContent = '0'; requestAnimationFrame(step);
  }
  var io = new IntersectionObserver(function (es) {
    es.forEach(function (e) { if (e.isIntersecting) { run(e.target); io.unobserve(e.target); } });
  }, {threshold: 0.6});
  els.forEach(function (el) { io.observe(el); });
})();
</script>
</body>
</html>
"""


def _base_url(request: Request) -> str:
    host = request.headers.get("x-forwarded-host") or request.headers.get("host") or ""
    local = host.startswith(("localhost", "127.0.0.1"))
    return f"{'http' if local else 'https'}://{host}"


@app.get("/", response_class=HTMLResponse)
def render_landing(request: Request):
    import html as _html
    guide = os.environ.get("GUIDE_URL", "").strip()
    link = f'<a href="{_html.escape(guide, quote=True)}" target="_blank" rel="noopener">사용설명서</a>' if guide.startswith("http") else ""
    btn = ('<div style="text-align:center;margin-bottom:24px">' + link.replace("<a ", '<a class="btn" ', 1) + '</div>') if link else ''
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
        uvicorn.run(app, host="0.0.0.0", port=8000)"""
AI기반 간호실습 배치 의사결정 프로그램

버전 기록 (수정할 때마다 APP_VERSION 을 올립니다)
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

APP_VERSION = "v13.2"
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
  f.append('max_period_capacity', $('capPeriod').value || 50); f.append('max_hospital_capacity', $('capHosp').value || 10);
  f.append('gender_criteria', $('gender').value); f.append('exclude_past_hospital', $('exPast').checked);
  f.append('use_transit_api', true);
  f.append('hospital_caps', capsJson());
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
  finally { b.disabled = false; b.textContent = '배정 실행'; }
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
  const pc = [['1지망', '#0f766e'], ['2지망', '#2dd4bf'], ['3지망', '#99f6e4'], ['비지망', '#cbd5e1']];
  const pv = pc.map(([l]) => rs.filter(r => l === '비지망' ? r.preference_status.startsWith('비') : r.preference_status.startsWith(l)).length);
  const tot = pv.reduce((a, b) => a + b, 0); let acc = 0;
  const stops = tot ? pc.map(([, col], i) => { const a = acc / tot * 100; acc += pv[i]; return `${col} ${a}% ${acc / tot * 100}%`; }).join(',') : '#e2e8f0 0 100%';
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
  new N.Polyline({map: m, path: [a, b], strokeColor: '#0f766e', strokeWeight: 4, strokeOpacity: 0.9, strokeStyle: 'shortdash'});
  new N.Marker({position: a, map: m, icon: {content: pin('#d97706', '거주지'), anchor: new N.Point(0, 0)}});
  new N.Marker({position: b, map: m, icon: {content: pin('#1d4ed8', hname), anchor: new N.Point(0, 0)}});
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
  L.polyline([a, b], {color: '#0f766e', weight: 3, dashArray: '6 6'}).addTo(m);
  L.circleMarker(a, {radius: 8, color: '#fff', weight: 2, fillColor: '#d97706', fillOpacity: 1}).addTo(m).bindTooltip('거주지', {permanent: true, direction: 'top'});
  L.circleMarker(b, {radius: 8, color: '#fff', weight: 2, fillColor: '#1d4ed8', fillOpacity: 1}).addTo(m).bindTooltip(esc(hname), {permanent: true, direction: 'top'});
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
    <div class="box" style="background:#f0fdfa"><b style="color:var(--teal-d)">${esc(r.hospital)}</b> · 병원 내 ${esc(r.rank)}순위 · 종합 <b>${esc(r.composite_score)}점</b></div>
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


LANDING_PAGE = """<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>AI기반 간호실습 배치 의사결정 프로그램 | 로켓단</title>
<meta name="description" content="안산대학교 간호학과 팀 로켓단이 만든 실습 병원 배정 도우미. 통학시간과 희망 병원을 함께 계산해 배정안을 추천하고, 최종 결정은 교수자가 합니다.">
<link rel="canonical" href="__BASE__/">
<meta property="og:type" content="website">
<meta property="og:title" content="AI기반 간호실습 배치 의사결정 프로그램 | 로켓단">
<meta property="og:description" content="통학시간과 희망 병원을 함께 계산해 간호학과 실습 병원 배정안을 추천합니다.">
<meta property="og:url" content="__BASE__/">
<script type="application/ld+json">{"@context":"https://schema.org","@type":"WebApplication","name":"AI기반 간호실습 배치 의사결정 프로그램 (로켓단)","url":"__BASE__/","applicationCategory":"EducationalApplication","inLanguage":"ko","description":"통학시간과 희망 병원을 함께 계산해 간호학과 실습 병원 배정안을 추천하는 웹 서비스"}</script>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css" media="print" onload="this.media='all'">
<style>
  :root { --pri:#3830f9; --pri-d:#241e9c; --pri-l:#eeedff; --ink:#1b1f3b; --muted:#565b70; --line:#e3e5ee; --alt:#f6f6fb; }
  * { box-sizing:border-box; }
  html { scroll-behavior:smooth; scroll-padding-top:72px; }
  body { margin:0; font-family:'Pretendard',-apple-system,'Noto Sans KR','Malgun Gothic',sans-serif; color:var(--ink); background:#fff; line-height:1.7; word-break:keep-all; overflow-wrap:break-word; }
  img.px { image-rendering:pixelated; image-rendering:crisp-edges; }
  .wrap { max-width:1080px; margin:0 auto; padding:0 24px; }
  a { color:var(--pri-d); }

  header.top { position:sticky; top:0; z-index:5; background:#fff; border-bottom:1px solid var(--line); }
  .top-in { display:flex; align-items:center; justify-content:space-between; height:64px; gap:16px; }
  .logo { display:flex; align-items:center; gap:10px; text-decoration:none; color:var(--ink); }
  .logo img { height:30px; }
  .logo b { font-size:19px; font-weight:800; letter-spacing:-.4px; }
  .logo span { font-size:13px; color:var(--muted); }
  .logo .ver { font-style:normal; font-size:12px; font-weight:700; color:var(--pri-d); background:var(--pri-l); padding:2px 8px; border-radius:99px; }
  .nav { display:flex; gap:24px; align-items:center; font-size:15px; font-weight:600; }
  .nav a { color:var(--muted); text-decoration:none; }
  .nav a:hover { color:var(--ink); }
  .btn { display:inline-block; background:var(--pri); color:#fff !important; padding:13px 22px; border-radius:8px; font-weight:700; text-decoration:none; font-size:15px; }
  .btn:hover { background:var(--pri-d); }
  .btn.sm { padding:8px 16px; font-size:14px; }
  .btn.line { background:#fff; color:var(--ink) !important; border:1.5px solid var(--line); }
  .btn.line:hover { border-color:var(--ink); }

  .hero { padding:72px 0 64px; border-bottom:1px solid var(--line); }
  .hero-in { display:grid; grid-template-columns:1.05fr 1fr; gap:48px; align-items:center; }
  .hero .eyebrow { font-size:14px; font-weight:700; color:var(--pri-d); margin-bottom:14px; }
  .hero h1 { font-size:44px; line-height:1.28; letter-spacing:-1.2px; margin:0 0 18px; }
  .hero p { font-size:17px; color:var(--muted); margin:0 0 28px; }
  .hero .cta { display:flex; gap:10px; flex-wrap:wrap; }
  .crew { background:var(--pri); border-radius:20px; padding:36px 28px 22px; color:#fff; }
  .crew img { width:100%; display:block; }
  .crew .names { display:grid; grid-template-columns:33fr 26fr 29fr 28fr; gap:6px; margin-top:14px; text-align:center; font-size:13px; font-weight:700; }
  .crew .names small { display:block; font-weight:500; opacity:.75; font-size:12px; }

  section { padding:80px 0; }
  section.alt { background:var(--alt); }
  h2 { font-size:32px; letter-spacing:-.8px; line-height:1.35; margin:0 0 14px; }
  .lead { font-size:17px; color:var(--muted); margin:0; max-width:640px; }

  .story { display:grid; grid-template-columns:1fr 1fr; gap:56px; align-items:start; }
  .story p { color:#33374f; font-size:16.5px; margin:0 0 14px; }
  .pains { border-top:2px solid var(--ink); }
  .pains div { display:grid; grid-template-columns:96px 1fr; gap:14px; padding:18px 0; border-bottom:1px solid var(--line); }
  .pains b { font-size:15px; }
  .pains span { color:var(--muted); font-size:15.5px; }
  .pains div:last-child b { color:var(--pri-d); }
  .pains div:last-child span { color:var(--ink); font-weight:600; }

  .how { margin-top:44px; display:flex; flex-direction:column; gap:56px; }
  .row { display:grid; grid-template-columns:1fr 1.25fr; gap:44px; align-items:center; }
  .row.flip { grid-template-columns:1.25fr 1fr; }
  .row.flip .txt { order:2; }
  .row .no { font-size:14px; font-weight:800; color:var(--pri); margin-bottom:6px; }
  .row h3 { font-size:23px; margin:0 0 10px; letter-spacing:-.4px; }
  .row p { color:var(--muted); margin:0 0 10px; font-size:16px; }
  .shot { border:1px solid var(--line); border-radius:12px; overflow:hidden; background:#fff; box-shadow:0 10px 30px rgba(27,31,59,.08); }
  .shot img { width:100%; display:block; }
  .shot.tall { max-width:400px; justify-self:center; }
  .formula { font-size:14px; background:var(--pri-l); color:var(--pri-d); border-radius:8px; padding:10px 14px; margin-top:6px; }

  .res { display:grid; grid-template-columns:.9fr 1.1fr; gap:48px; align-items:start; }
  .big { font-size:64px; font-weight:900; color:var(--pri); letter-spacing:-2px; line-height:1.1; margin:24px 0 4px; }
  .big + div { font-weight:700; }
  table.cmp { width:100%; border-collapse:collapse; background:#fff; font-size:15.5px; }
  .cmp th, .cmp td { padding:15px 14px; border-bottom:1px solid var(--line); text-align:right; }
  .cmp th:first-child, .cmp td:first-child { text-align:left; }
  .cmp thead th { font-size:14px; color:var(--muted); border-bottom:2px solid var(--ink); }
  .cmp td.us { font-weight:800; color:var(--pri-d); }
  .cmp td small { display:block; font-weight:600; color:var(--muted); font-size:12.5px; }
  .note { font-size:13px; color:var(--muted); margin-top:12px; }

  .demo { display:grid; grid-template-columns:1fr 1.15fr; gap:36px; margin-top:36px; align-items:start; }
  .demo .who { display:flex; gap:14px; align-items:center; margin-bottom:18px; }
  .demo .who img { height:64px; }
  .demo .who p { margin:0; color:var(--muted); font-size:15px; }
  .demo .wl { display:flex; justify-content:space-between; font-size:14px; font-weight:700; margin-bottom:8px; }
  .demo input[type=range] { width:100%; accent-color:var(--pri); }
  .dr { display:grid; grid-template-columns:1fr auto; gap:4px 10px; padding:11px 14px; border-radius:10px; border:1px solid var(--line); margin-bottom:8px; background:#fff; transition:background .2s, border-color .2s; }
  .dr.top { background:var(--pri-l); border-color:var(--pri); }
  .dr .nm { font-weight:700; font-size:15px; }
  .dr .nm em { font-style:normal; font-size:11px; background:var(--pri); color:#fff; padding:1px 7px; border-radius:4px; margin-left:6px; vertical-align:middle; }
  .dr .sub { font-size:12.5px; color:var(--muted); }
  .dr .sc { font-weight:800; font-size:17px; color:var(--pri-d); text-align:right; }
  .dr .tr { grid-column:1 / 3; height:6px; border-radius:3px; background:#e6e7f0; overflow:hidden; }
  .dr .tr i { display:block; height:100%; background:var(--pri); transition:width .25s; }

  .faq { max-width:760px; margin-top:28px; }
  .faq details { border-bottom:1px solid var(--line); }
  .faq summary { cursor:pointer; padding:18px 0; font-weight:700; font-size:16px; list-style:none; display:flex; justify-content:space-between; gap:12px; }
  .faq summary::-webkit-details-marker { display:none; }
  .faq summary::after { content:'+'; color:var(--muted); font-weight:400; font-size:20px; line-height:1; }
  .faq details[open] summary::after { content:'−'; }
  .faq p { margin:0; padding:0 0 18px; color:var(--muted); font-size:15.5px; }

  .team { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:16px; margin-top:36px; }
  .mem { text-align:center; padding:26px 10px 20px; border:1px solid var(--line); border-radius:14px; background:#fff; }
  .mem img { height:120px; }
  .mem b { display:block; margin-top:12px; font-size:16px; }
  .mem span { font-size:13.5px; color:var(--muted); }
  .end { margin-top:44px; display:flex; align-items:center; justify-content:space-between; gap:20px; flex-wrap:wrap; border-top:1px solid var(--line); padding-top:28px; }
  .end p { margin:0; font-size:17px; font-weight:700; }

  footer { padding:28px 0 44px; color:var(--muted); font-size:13px; background:var(--alt); }
  footer a { color:var(--muted); margin-right:14px; }

  @media (max-width:880px) {
    .hero-in, .story, .row, .row.flip, .res, .demo { grid-template-columns:1fr; gap:28px; }
    .row.flip .txt { order:0; }
    .team { grid-template-columns:repeat(2,minmax(0,1fr)); }
    .nav a.hide { display:none; }
  }
  @media (max-width:560px) {
    .hero { padding:44px 0 40px; } .hero h1 { font-size:31px; } .hero p { font-size:16px; }
    section { padding:56px 0; } h2 { font-size:25px; } .big { font-size:48px; }
    .logo span { display:none; } .crew { padding:24px 16px 16px; } .crew .names { font-size:11.5px; } .crew .names small { font-size:10.5px; }
    .cmp th, .cmp td { padding:12px 8px; font-size:14px; }
    .pains div { grid-template-columns:72px 1fr; }
  }
</style>
</head>
<body>
<header class="top"><div class="wrap top-in">
  <a class="logo" href="/"><img class="px" src="data:image/png;base64,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" alt=""><b>로켓단</b><span>안산대학교 간호학과</span><em class="ver">__APP_VERSION__</em></a>
  <nav class="nav"><a class="hide" href="#why">만든 이유</a><a class="hide" href="#how">사용 방법</a><a class="hide" href="#result">시험 결과</a><a class="hide" href="#faq">자주 묻는 질문</a><a class="btn sm" href="/app">배정 시스템 열기</a></nav>
</div></header>

<main>
<div class="hero" id="intro"><div class="wrap hero-in">
  <div>
    <div class="eyebrow">안산대학교 간호학과 캡스톤디자인 · 팀 로켓단</div>
    <h1>실습 병원 배정,<br>통학과 희망을 같이 봅니다</h1>
    <p>간호학과 3학년 학생 셋이 만든 실습지 배정 도우미입니다. 학생의 통학시간과 희망 병원을 점수로 바꿔 배정안을 추천하고, 교수님은 그 근거를 보고 최종 결정합니다.</p>
    <div class="cta"><a class="btn" href="/app">배정 시스템 열기</a><a class="btn line" href="#how">화면 먼저 보기</a></div>
  </div>
  <div class="crew">
    <img class="px" src="data:image/png;base64,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" alt="팀 로켓단 캐릭터: 강지우, 최영락, 정현민, 강민정 교수님">
    <div class="names"><div>강지우<small>팀장</small></div><div>최영락<small>팀원</small></div><div>정현민<small>팀원</small></div><div>강민정<small>지도교수</small></div></div>
  </div>
</div></div>

<section id="why"><div class="wrap story">
  <div>
    <h2>왜 만들었나요</h2>
    <p>간호학과 3~4학년은 매 학기 여러 병원으로 임상실습을 나갑니다. 그런데 이 배정을 교수님 한 분과 조교 선생님이 엑셀로 한 명씩 맞추고 계셨습니다.</p>
    <p>학생들의 거주지, 희망 병원, 병원 정원, 실습 규칙(평일 8~10일, 금요일 종료)을 한꺼번에 따지기는 어렵습니다. 그러다 보니 집에서 먼 병원에 가게 되는 학생도, 왜 그 병원인지 모르는 학생도 생깁니다.</p>
    <p>저희가 직접 겪고 있는 문제라서 직접 풀어 보기로 했습니다.</p>
  </div>
  <div class="pains">
    <div><b>교수님은</b><span>수백 명을 손으로 대조하고, 정원이나 기준이 바뀌면 처음부터 다시 맞춰야 합니다.</span></div>
    <div><b>학생은</b><span>통학 부담이 충분히 반영되지 않고, 배정 이유를 듣기 어렵습니다.</span></div>
    <div><b>그래서</b><span>계산은 프로그램이 하고, 결정은 교수님이 하도록 만들었습니다.</span></div>
  </div>
</div></section>

<section id="how" class="alt"><div class="wrap">
  <h2>이렇게 씁니다</h2>
  <p class="lead">아래는 실제 프로그램 화면입니다. 학생 데이터는 가상이고, 이름과 학번은 일부 가렸습니다.</p>
  <div class="how">
    <div class="row">
      <div class="txt">
        <div class="no">1단계</div>
        <h3>조건 정하기</h3>
        <p>학년·학기·교과목을 고르면 그 교과목의 실습 병원 목록이 불러와집니다. 실습 기간(8~10일), 병원별 정원, 성별·GPA·출생연도·과거 실습기관 조건도 여기서 정합니다.</p>
      </div>
      <div class="shot"><img src="data:image/jpeg;base64,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T1pG7taXHfq5XEcgfJsr7tVdbYtRY/EZHIsqymCaav1LWNkHum/3kjS7bf8AwghHS0JRsKKliMtWzePivVHOdBIXDx2lrmua4tc0g9xBBH5li4tZ4ySa3FwXg+rO+u/hpSyAuadjs5jSCPzpDmBOEmwoter6zx1nD0sq2HImtc4urbFQmneOEkHibE13D3eVT19U4+3HiZKxmkiykz4IXGIsLXMa8u4mv2c3bq3Du33SGJRmkWvv1UJbFmLH4jI5JlWQwzTVuqaxkg903+8kYXEeXhBHk71ao6kxuQr4+eKZwbfkfDC17HBxkYHFzSNuRHA7ff4Eh6xJlkVG/lIcfYx0ErZC+/Oa8ZaBsHCN79zz7tmHu357KjjtXYzLTsip9ul43uYJRj7AhJaSD/elnByII3370SbEpGcRUcPlYc3ja9+u2RsM4LmtkADhzI57E/ArymooREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAFiH+/S/KP71l1iH+/S/KP71qkjPbV6Xlq9LZkHuUT1Ke5RPQFzH+9P+X/AK2qmP96f8v8AgFbXN6zSC1zXONsZbTslarXmsS9orSGKCURSOaydj3cLy5vC7hadjxDn5VsaKdKZTg+oMJao1MkJYshj2zQ5OzFHZstsTPj6muNnvLpNwS0/4t9gOYW6aox0l7UVWLrrljLNhdYxravVwsotaWh8ji9x6xziQ0jY+KSOEbknfbVOvdifFZgimie0sc2RocC094O/kKr5PCYzNRsiyWOqXY4zuxtmFsgafhHEDst3tW4zd17zWNC06lPK59kFFmOsg122KccMbGNcGHx2mMkEO338hG3MdywWpdOm7qPJ4+lJZnmvCO9ZhbDXDY28mDaR/j8+q7m77eXbcb9Ix2Lo4iv2fH0q1OvuXdVXibG3c952aAN1DlMBis31fhPGUrvVb8HaYGycG/ftxA7Je/deF39t01DE46LJ+yvIUo5aWbsW+qneyGE2og2KLeFsnEWncAEEnYF3MAgrC3aUmEwlupFgs3jsNM+DeKSWrLFXd1rN3tAm4/GIG43I3JO25O/UqNCpjKzKtGrBVrR78MMEYYxvl5AcgvdmtDchdDYiZLE4glkjQ4HY7jkfjAKU1XWo3cg6ZWJpmrrOTyFqXFVqGTdjrVK1VnkhhYQJHhgjeC5zdwB1nlHeveiK0oy2pLBmtwxdrZH2CQRFsZ7NB4xLQXcXk5PLfi8q3VRQ1oa75nxRMY+d/WSFrQC92wG5+E7AD8wWaXHD1kNTBz3Vuv5sbkblbHPnd1OPsbAUpHjtbXMEbd+Hny4/iWQ07l4GQZ+xV7Rcms3DYDIK0gcwOY1reT2t39757b7brd0TCIzrkrmZzqg5Bbr5W/Vx9vIw5SjkrT8VWmsvjgHBLHP4zmjiduS6Qkbs28X49ld1dgdQWMhL2VtyV7oo4+3064ZI5rSSA8i7EHHm7f8AuwPGO2y6bPWhtNY2eJkrWPbI0PaDwuadw4fGCNwVKtX/ADnyJd8jleOdPX01qmS1YdjxUeTba+ttI53VMJ43PktBzDGWDk0nl3LXKGMuw3qBipZPlPG2MsqH/ZwXAF7BJjGsbwgk+6byHeu2S4yjPFbilp15IrnOwx8YLZvFDfHB914oA5+QBYyPRGmIZGSR6cw7JGEOa5tKIFpHcQeFWmtTL3GXRhC3n7qvtEOnLj4r0lZkELn2J2RB8xia0l/Bza0PIHIncD4Fqr9CVziKT8bXhswRWKdigyeFnX0ohOySRrZSd+Hh4iG945jc7gDok0MdiKSGaNskUjS17Hjdrge8EHvCwkeiNMQyMkj05h2SMIc1zaUQLSO4g8KzTVDnsfDPntNVUyoNQzOjZqmau3KeLN2S9I+Z1qGlWM0XE3hLBJLOw7bfA3y96qYLD2aWH1Jj8LQs4+9GabJK7BDBI8NYzrCx0b3ND3sDtjuCCRz8q6woo60MU000cTGyzEGR7WgF+w2G58uw5KqtpR2cg6Zd7OJy+9Qv9kc1vhbH42bJ45tVt6wJ7EcnXDrHtL3SbN9xsHEjcE7bHnkLunZZstldPMzMg8MxC/NJJUa54DBFEdnh7QHHhafcbcz+RdAnrQ2msbPEyVrHtkaHtB4XNO4cPjBG4KpZPT2IzT435LFUbr4xsx1muyQtHwDiB2UVUZ3L1QdM53v0ZjdO4W5i8xnprM81iK0+F0U8xZxPDYwDyaABseXctXzWNzgt5Gbqsjx9fNIbZuf7IanVO2j6nj913D3G/F42633GYPF4VsjcZjadJspBeK0DYw/bu34QN+8q85rXtLXAOa4bEEbghSpzwNU4azjmDowQaYzFnI5G+5mNfSyAnjZEZQY6sTmtA4A0gd3Md3ee8rJX69+4cZV1Ky5Pj7t+DqhbNSdnECXBjmRsb7ru33kA237tyulRUKsAmEVaFgm2EgawDj2aGjf4eQA/IFQoaVwOKsi1QwmNqWQCBLXqxxvAPfzA3W3XNV7szyOaoaoVO6M8SvqzNx6bwFqSHiZZ6h7ajI4HSDrAw8I2aDsN9u/ktJh1YM5l8d2uxJHBStwWG8VKZpfvXe14GzPOPHftyBXVUWaao1mqlOC38zQNTT5PM3rmMbjcq7FS12M3ihj2MrZeJ2xc9vItAHftzVCpjMjd09n3V5JbDJ8nY4sdNBHIGgTu34OGSIh2+xJMhHi8l05RQVoarXtgiZE173SODGgbucdyT8ZPNFVCjOuQ1Lnf6Qcdq0M5QydSTIR3GssXo+oferuligkIAaA0ZB3LcE7kHmStgzOLmkt5SnVhz3g7KyOddrwVoeGQ8Ia7gke4Foe1oB7/AIuFdCnrQ2RGJomSCN4kZxtB4XDucPjHwqVW+893sLueJp+hXmS1qp7oHwOdlNzFIWlzP9nh8U8JI3HxEj41rGpohT1FYbWhnvZZnV2e3cVo2KwPGGMb1NV7Gx7cQ4T7rc8QJ5rqMFSCs+d8MMcbp39ZKWNAL3bAcR+E7ADf4gsdktKYHNWe05LCYy7YDQzrbNWOR/CO4buBO3NRVYrsXJC7g1vfmc4yjrE+jbty5TtVBZzDTLV6kydeJHxxghksTXFvPfhLee23lXiUSvzem43yZNsxstrQTWMPFWEDQxz+FruDu8TbhBA/QulR6VwMOOlxseExrMdM7jkqtqxiJ7uXMs22J5Dnt5AoqWjNN421HapafxNa1Ed2TQU42Pae7k4N3C0qks7kjLpb5+bZQ6QKNOfCx3rzyIcXN2oN6lkoeeBzA0tf4p988pAB2O4WmUdNOwmTwptRZCo21le0MaJozXEhY48HUsds0kDkWhwHlPlXW5YmTxvilY18bwWuY4bhwPeCPKFjMfpbBYmwLOPwuNqWAC0S16rI3bHvG4AKzRVdcmq6byznpNDl0/ay+q8kynMSKLzJYc5jIo3WuGN8Xi+O0+K/m8xkgsHxLb9G0chRo5DwnAyC1Penm4Y5esaWudyIds3ff8g/IFbvaTwGTsvtXsHjLVl+3FNPUje92w2G5I3PIbK7jsXRxFc18dSrU65cXdVXibG3c952aAN+QRVRTG6A6cZznE5XlsFqV2Vllhq3W7WXTSPowGOKyeEt3LfCLeR5Huad2j8imnr5K/oXHijQbahbbe8QwNkrTwvZI7n77JueMPJPH5QusqKvWhqRCKvEyKIFzuBjQBuSSTsPhJJ/Ol5wlsF3FvOJxvSp1NR1DjYMhDeY+zNII3XbUvVluzncPDzBIYDtv3kAqvmKktvM5LqZLt10U7oZp46xJkc3YbSGPGOa4gbDbidy8q7ZJWhmlhlkijfLCS6N7mglhIIJB8nIkcvIVibOjdOXbEtizp/EzWJXcUkstONznn4SSNyVq+plol2FgVdCQ2I9MVevbagc7iDa88UcZgaCWgNa2KLZpA4huwHxlyvLY+xYzNs0y+5HFbkZZmOP610rhuD40eMc0O4tt+bu4j413Gjj6mMrMq0asFWszfhhgjDGN3O52A5Dmd1i7GjNN255bFjT+JmnlcXvkkpxuc9x7ySW7kqXledQu/tg0mrQy4wWnOpq54U953WKlR8NWaHuDBtwwDh3BdzaD43NWK2I8H5bSLmQ5alEy9OxtO/NFJuTBM8yeI53Mlx73fDyXQqVGrjazKtKtDWrR78EUDAxjdzudgOQ5le5a0M8kMksTHyQOL43OaCWOILSR8B2JH5CUdcthU4Qcdttky+WvvdFaZNUuvhsPx2Jm6u4Wt4f7wsmAeOYOx5+KOe3JZvEzX7+C0rlaWJ7TBTuzO7HRhjrOiiEc0IAbJIG+6O/eO/u5Lb7eidM37Mtm3p3EWLMruKSWWlE97z8JJbuSsrRo1cZVjqUq0NWrENmQwRhjGc9+TRyHMpeV1LsF39zfaalk79u/ldLutYqxjw3KEMbPJG5zv8AZbG/vbnAfpXnSU2XwWKr4uzprJOeyebeeOaqYg18z3B3v3Fts4f4d/iW5S1oZ3wvliY98D+sjc5oJY7Yt3HwHZxH5CVKpJYxnPSaPoS5WxGl8K25kX8WRkMVeObh2D93Hgbs0eQE8ye7vW8KKtWhpwMgrxMhhYNmxxtDWt/IApVG5clSgIiKFCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiALEP9+l+Uf3rLrEP9+l+Uf3rVJGe2r0vLV6WzIPconqU9yiegLmP96f8AL/gFbVTH+9P+X/AK2ub1mkERFChERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAFiH+/S/KP71l1iH+/S/KP71qkjPbV6Xlq9LZkHuUT1Ke5RPQFzH+9P8Al/wCtqpj/en/AC/4BW1zes0gvxzg1pc4gNA3JPkX6uedM+RtV9IRYylKYrWevQYpkje9olJ4v0ta4fnVpV5wcv1FsrCyqtWphGOk19qHWtuxW0JUqsxld5ilzmQ3ML3DvELBzft8J5f8iffsU6Rn+M7pMawnvazBQED8hJ3W84nFVMHjKmNoxNiqVYxFGwDuAH7/ACk+Uq4t30v4o+ev01VavW9dTe6p0pdiTXOWc79iXSL+E/8AYNf609iXSL+E/wDYNf610REvvdwRdCs+tV46/uOd+xLpF/Cf+wa/1p7EukX8J/7Br/WuiIl97uCGhWfWq8df3HO/Yl0i/hP/AGDX+tPYl0i/hP8A2DX+tdERL73cENCs+tV46/uOd+xLpF/Cf+wa/wBaexLpF/Cf+wa/1roiJfe7ghoVn1qvHX9xzv2JdIv4T/2DX+tPYl0i/hP/AGDX+tdERL73cENCs+tV46/uOd+xLpF/Cf8AsGv9aexLpF/Cf+wa/wBa6IiX3u4IaFZ9arx1/cc79iXSL+E/9g1/rT2JdIv4T/2DX+tdERL73cENCs+tV46/uOd+xLpF/Cf+wa/1p7EukX8J/wCwa/1roiJfe7ghoVn1qvHX9xzv2JdIv4T/ANg1/rT2JdIv4T/2DX+tdERL73cENCs+tV46/uOd+xLpF/Cf+wa/1p7EukX8J/7Br/WuiIl97uCGhWfWq8df3HO/Yl0i/hP/AGDX+tPYl0i/hP8A2DX+tdERL73cENCs+tV46/uOd+xLpF/Cf+wa/wBaexLpF/Cf+wa/1roiJfe7ghoVn1qvHX9xzv2JdIv4T/2DX+tPYl0i/hP/AGDX+tdERL73cENCs+tV46/uOd+xLpF/Cf8AsGv9aexL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" alt="배정 조건 설정 화면" loading="lazy"></div>
    </div>
    <div class="row flip">
      <div class="txt">
        <div class="no">2단계</div>
        <h3>명단 올리고 결과 보기</h3>
        <p>엑셀이나 CSV 명단을 올리면, 병원 정원을 지키면서 학생 전체의 점수 합이 가장 높은 조합을 계산합니다(헝가리안 알고리즘). 무작위로 배정했을 때와 비교한 결과가 함께 나옵니다.</p>
        <p>올린 파일은 서버에 저장하지 않고, 결과는 엑셀로 내려받습니다.</p>
      </div>
      <div class="shot"><img src="data:image/jpeg;base64,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89PIJfj98D/z1ekdG7W5+w1e02GnaqFuH5BP8fvgf+eq+QT/AB++B/56vSWjdrc/Ymr2mw08Rbh+QU/H74H/AJ6eQU/H74H/AJ6dJaL29z9hq9psNPUW4XkFPx++B/56r5BX8fvgf+er0lova3P2Gr2mw09VVuD5BX8fvgj+eq+QV/H74I/nq9J6L29z9hq9psNPkW4PkFvx9+CP56eQW/H34I/np0novb3P2Gr2mw0+VVuB5Bb8ffgj+eq+QX/H34I/nq9J6L29z9iavabDT9FuB5Bf8ffgj+enkGPx9+CP56dJ6L29z9iata7DUBFuB5Bj8ffgj+enkGPx9+CP56vSmi9vc/Yata7DT9VW3/kGPx9+CP56eQZ/H34I/nq9KaL29z9hq1rs9DUFVW33kGfx9+CP56eQa/H34I/nq9KaJ29z9iata7DUFVW3vkGvx9+CP56r5Br8fPgj+enSuidvc/Yata7DUFVW3vkG/wAfPgj+enkG/wAfPgj+er0ronb3P2GrWuw1DRbe+Qb/AB8+CP56eQc/Hz4I/np0ronb3P2GrWuz0NvERF8cesEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBEJAGTwC4HVsDTgyj9XFCNpZnOi6/PqfzweIpz6n88HiKEvLadhF1+fU/ng8RTn1P54PEUF5bTsIuvz6n88HiKc+p/PB4igvLadhF1+fU/ng8RTn1P54PEUF5bTsIuvz6n88HiKc+p/PB4igvLadhF1+fU/ng8RTn1P54PEUF5bTsIuvz6n88HiKc+p/PB4igvLadhF1+fU/ng8RTn1P54PEUF5bTsIuvz6n88HiKc+p/PB4igvLadhF1+fU/ng8RTn1P54PEUF5bTsIuvz6n88HiKc+p/PB4igvLadhF1+fU/ng8RTn1P54PEUF5bTsIuvz6n88HiKc+p/PB4igvLadhF1+fU/ng8RTn1P54PEUF5bTsIuvz6n88HiKc+p/PB4igvLadhF1+fU/ng8RTn1P54PEUF5bTsIuvz6n88HiKc+p/PB4igvLadhFxR1EUpwx4J7y5ULWoREQoREQFAHSP5NnB2MknuBdgUU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" alt="배정 결과 대시보드 화면" loading="lazy"></div>
    </div>
    <div class="row">
      <div class="txt">
        <div class="no">3단계</div>
        <h3>학생별 근거 확인</h3>
        <p>학생을 누르면 거주지와 병원 위치, 점수 계산 과정, 다른 병원이었다면 몇 점인지가 나옵니다. 학생에게 배정 이유를 설명할 때 이 화면을 그대로 보여 주면 됩니다.</p>
        <div class="formula">종합 매칭점수 = 통학 점수 × 70% + 지망 점수 × 30%<br>통학 점수는 소요시간·환승·도보로 만든 피로지수(MFI)로 계산</div>
      </div>
      <div class="shot tall"><img src="data:image/jpeg;base64,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iKpFS73fGPim93xj4pA0iKpFS73fGPim93xj4pA0iKpFS73fGPim93xj4pA0iKpFS73fGPim93xj4pA0iKpFS73fGPim93xj4pA0iKpFS73fGPim93xj4pA0iKpFS73fGPim93xj4pA0iKpFS73fGPim93xj4pA0iKpFS73fGPim93xj4pA0iKpFS73fGPim93xj4pA0iKpFS73fGPim93xj4pA0iKpFS73fGPim93xj4pA0iKpFS73fGPim93xj4pA0iKpFS73fGPim93xj4pA0iKpFS73fGPim93xj4pA0iKpFS73fGPim93xj4pA0iKpFS73fGPim93xj4pA0iKpFS73fGPim93xj4pA0iKpFS73fGPim93xj4pA0iKpFS73fGPim93xj4pA0iKpFS73fGPim93xj4pA0iKpFS73fGPim93xj4pA0iKpFS73fGPim93xj4pA0iKpFS73fGPim93xj4pA0iKpFSiRw7yqhj949veoaprTPpERDR4/1SuNNafDHUPyhUfaOXZb/AFSuNNafDHUPyhUfaOQHZipX+u73qqVK/wBd3vVRztNR8oiwzUjtdRXF7rLWaWgtbtrYxcopzLuI5glrwO3s5KrWfPXVdUxJmaLVNm1xqcXDWlqv1bpqjqrDBTSR1rIpW0zTKHE8TdJnAAHYR2qxUXldv9TcdS26J9ov7KC2edRVenmhoY85GXGWYtIbjJA5+wqul8J9zh8VRhM4uPv1uN5otJWbyi6vrbV5P6G3R2utvF/t01VPUXJz42Zjxn+aHLIPRUd28pOoKDXk9hqLvSwzi6WeBtHEIzmOWMmoDNzd7m7sczzGR2ZW9FVeu9a7vqc/j7O5fhxC403o8YN8ouc7j5UNfU+pqh8dXaTZqGqNsnqmW2Z1DHO5wA3u4nE3AkNy30efzrMfJ9rXUd/q9cWiUumu9vuFTHR1c9I4UEYbtDYi9uCSCSdud2O9ZVDavLdPpzTNLLaHXo4czHrya8UbbRan1NePKlpjT1zvM8+jZYbfA+ofHHTVW5waM4GZO1ZO7+NOobBYa+13a32uoqKSOaqZLQmoa572NOG/lG7QCT3nuUu4T11gdFby7t1zE9YmYotY2q96ttvlLoNNXi62+5UlTbpa1z6ehNO6Pa4NA9d2eZVDc/KDdDerrS22+wPioql1O+OLStdVmFw/MdJHJtcQCOwDtVuzEbecEeU0pN1KIcbN0741G3EWkqvynaut+gnXmqt8TbkL6ygijko30gqackAODJXZYXZPMnA+YrDdV+X3VNK3VlEKe22irpuCKVj62OSohyAXbQ1r2TZ94wD1VVm3q61czlXl9lQk6px7vH9rOn0Wo9c6t1XQV9tgt8k9HbJLeKh1dS2l1xfLU/0JY0jhgjnkjmT2r705rbUF11XoqkuLoaU3C01FRXUUO1zRM1zdvPmWnBztzyzg5wis2+PCeXpvN/F0TEPZxaX/ANLjEwbZRFgl41tdbZrqWx0un6u7UjbXHW/7C6Bskb3SvYdxmmjBbhowBk5yuc9eCn2Pr2N9a49zO0WhfJ15aLvrq/aJpJ57BELhSVE1bTW24ComLmxgtEkRjzDg55B7ufIkq/1OttXRXaoh80qBWsufm8VsFjqHQSUnFDeKa3PCBMZ4m7IDfULSe3o7Nqq6+sYJODe422iw3Ql9r71XayjrajjR269SUlMNjW8OIQwuDeQGeb3HJyefarjpvUcl/k1Cx0DYfwVcZKFpDs8QNjY7ceh9Ps9i5vBT3T9sOaDw84++PJmQotQ+R/yg3HWlZTMqrrBWtGnaCsnbEI/Qq5JJxJnaMgkMZ6PYMdgyr7qXWWobPrrTNqhtFIzT9zrfMpK6eYumleaeSX8lG3k1reHguecknAbj0lt0NVXetweE93KTYKLX+rdYXqed1r0XbJ6+60k7H1ks8ToaZkTCHPibK8bXyvALGhudpdlxaBzrK/Xhm0bfLzZrZXS3O2073m3VlJLDKJQ3O0sLcuHtZuDgPRJyFmML3Xj4FjGDNEWmaPyk3KpN0gt+prFqSGKzVVa+ts9PtFuqI2jY2T8pI3Dsuwx2HDhnO4ZxtHTFZPcdNWasqX8SpqaOGWR+ANznMBJwOQ5nuVuvF7o4zyZmdXfPCOZdUVJdJzS22snEjojHE5wkZC6YswDzDG83Y6DmVpB2tNffh1oZWynTvALzMbA4XAkOwZW0JfxjACQ0u279zhhhb6SiUuDT1Sb6Ras1tqjVFPabLR6amZVXa8UrxD/7P2VJc1oL5gyaWNkYbuadrznJAwVbaPyk6ikrtLOq4HSOqrlV0NbbqClaJmOigcSx+6Rwy14DssfgjvK0qW5I3CTNyosD8retL7obSlVc9P6f/DFZFFLM8ySiOCmjjjL3SSHtPZgNbzce8LApfLfqKv0v5N6ex2i21OtdbwGeOOoe+OjpWMbukkdjLy0DsAOe3t7DKaXVq8PXkyvDX3vyN8otBR/wh6zT2l9dO1dZKeLVejpYYJqSgmcYK0zHELo3OG5rXE885IHPtOBcdN+U/W9v1vZdJ6/s9loajUtJLUWurtMsj2RSRt3Ohla/tcBzy04zgDOci3G9X+cJw+2JG419bDdiLl+3eWfVEflJumnanVlsuVBbRuhkoLFh13mj2uno4HOqdglAcB6x5ns5c9leQLWl411pi+3G9STunivtZTQxVETI5IIWFuyJwZyy3OD2nPeUVDavLdPGOv8AAbhw98cJ6/ybWREWQEREAREQBERAEREAREQBERAEREAREQBSwesfcolLB659yjNU6ydERQ+g8f6pXGmtPhjqH5QqPtHLst/qlcaa0+GOoflCo+0cgOzFSv8AXd71VKlf67veqjnaaj5WmPKOde6obDbqPQgNLQXOGriqjd4Bx2RPyPQOC3cOvZ7VudFqlw09x8ttZaWh0S0nujU/FM0idJ6i1c3WdHXaUk0/JqeCMTXCa6xVbGPjwGNbGwAgYz3q7Xbyf3q5aXqqPUtfbq6go6N5gpbbSTUz3SsjIY4yCUl3f6OMHK2wiVVNqFh/GryOdGS0JzU29fHF6oTk0fpDQV+rtN+Tu8W64wWm4We2S07o62jdKcyYz6O5uCA3v6qw3bSd6i19VV1xo5K2pN8shbcWURjZK1rHcVzO3DQcZwTjAyujkXTSu/e7/wD6vepx+Ao0aonZHlTdmNWqDUeovJJV3HV1RfbfFpuPfM2oa2qpal+ZAB6b2NmbG52cnJZlU2gJ9TWSzazlh01FcLy7UE++GGU0bKkFrd0zTMXcs9gBwtyoud76br1RHpyOryWi/pKcHM8Gvc5Ut3k513DBaqW4aFdVUVJb6q3zRx3mCJ0ommMm4OyduM4xzytuVvk3tGpdPi63vT1TQXyOiMYpY7pM7hcNpawZieGO5AHkO9bPRarrdajV0+bM2ORWdnVLd5bnHct25I0r5JPJbYhpbTmpa+01n8aY4uK6WarqGP3h5IBYX7QOQ5EY9iseoNK6rra1tZZtEXq0z1NxbW3DzfU4YyqZ/wBY3a2UBrncvSA5YXQyKu0brdRlZBZqyVksI1tJS8InU8e/vNXW60V1Vpa90Fx0NM4RMNTS0t6urbk2pnDTtblznFoyB3gcz7VqR1trnaaqKOjpJ5LncuEazTMWkX0cLzuHoOqGBpa1uT6fE7vaurEUVcOfDh1wLaZGqkknqnj4Ryx1HP3lg0rRVWpLBHJpj/Y6mNs1dcqS1S11QwxFu2HLDja4Dacg8ldNF2eyjyiUFwtMd/o4IoaljKGq0/5pTR8TDjtkbHGByaB6W4nlzW7USm0aUePH+MA8jpdo6/DZ/bG5rbj490FtmsFvqL7SXuSFxudJBJTQy8V4DY3lpcNgO05LG8yCRjksE8runKC5w2qsqNLWy6flxT1dfNZW3Opoqba926KLaXOy/aOQdjcXbTgrZqLG7uPt395p7R9PxtU2CmoYrlcrHaqWoaye42R9s/B5cGtayMCOGORrm5bsEbtuAdw7DjVTpXRUWtr1LeNEGjbQcWhohbNGT1MU8T2QPbUcRlPJG6RrhK0HuBIx256GRavfVe7vXqPuSMIMA8lVqpLbQX6ajrLpVRV9zfVZudqlt8seYo27dkjGbgAwek1oHPHaCpbfou/Wm9Xapo9SU8druVebhJSfg3MuS1jXM4pkIwRH3MB5n2LOkWanPlH2w5IRhD3z98ebNL+RjTVdp65UbKu1zUTm6VtkExdDsBnbLU72k4wXgFuR28x1Wba2sVwu990NU0dPxYLXd3VVU7e1vDj81nZuwSCfSe0YGTz6ZWZItVVt1Xutch4z38oNF6h4DrnfLbVR6vt0tHdaZ9LVW+lvNVHVUgEMso3w7mZcTNHyIx0653ZLRYNSWyr8xdqmmhc8RvNXPc6GbIw70eMWPx7W8jzGe0LOUUT+mOtnIPFyaat2hZ49Pa3nq5dTVFbFNXw0EdZd62Zr4eGRHtifIWvBycHacrYVitMddoO2Wq4wSCKe2RU1RC4ujdgxBrmkjDmntHLBCyNEnB0vu4JrjOI2qpbG+MekYFmq9P0X8WvwLFSySUUFO2KKnFQ9jiGAbG8TO780DOc9crnVtjc6lbcZKS4x6nFOW/gv+Ll4cxrzzMHnAmwRuAHED9nLdjHJdRoiqct7+upGpJLrrrYag8pNmlt2hNONobDJNNTV8D5acR1FdJTMeHcXD4mvmA5kFzOeDjsOFaNK3O4V2s9HUFNpsUFooZKuWR8Nlr6ZkRfC/wBJz6mnjblzjzIJcSefat7Iqq9c7eRGsF3czG/KHaqu+6B1Ta6CLjV1dbKmngi3Bu+R8TmtGSQBkkcyQFo2PyY610zYfJDqO02WKv1Jo+jfR19lfVxxumjlZtcGS5LNzeZ7cHPf2HpZFKanTiu7hK4puTTxUPv4xyUHLtZ5D9W6/wBO+Uy83ukprJqfVU9LPQ241LZm0raXHDbJI0bSX9hI7O32DIrXpbXHlL8pOltQa20nT6dtWmKOoj83dXRVZr55mbHYDMhrAOeD7ueTjoBFVXGEYbO7CPTmR469b/zByRdLVeqfU97082qrNO6WsFwjksVJS6CN0h5NDuMyRsRw4PJ9LJJJPNbW/g4Ut0pdG3r8LQVEdTLfKuVs1RbjQvqmOLSJjCQC3fzPZ7O5bhRVVu7D3R6Y+LglSlz3z64cQiIsFCIiAIiIAiIgCIiAIiIAiIgCIiAIiIApYPXPuUSlg9c+5RmqdZOiIofQeP8AVK401p8MdQ/KFR9o5dlv9UrjTWnwx1D8oVH2jkB2YqeVhDs9xVQhGUM1UyijRVPCZ0ThM6fSrJz0bKZFU8JnT6U4TOn0pI0bKZFU8JnT6U4TOn0pI0bKZFU8JnT6U4TOn0pI0bKZFU8JnT6U4TOn0pI0bKZFU8JnT6U4TOn0pI0bKZFU8JnT6U4TOn0pI0bKZFU8JnT6U4TOn0pI0bKZFU8JnT6U4TOn0pI0bKZFU8JnT6U4TOn0pI0bKZFU8JnT6U4TOn0pI0bKZFU8JnT6U4TOn0pI0bKZFU8JnT6U4TOn0pI0bKZFU8JnT6U4TOn0pI0bKZFU8JnT6U4TOn0pI0bKZFU8JnT6U4TOn0pI0bKZFU8JnT6U4TOn0pI0bKZFU8JnT6U4TOn0pI0bKZFU8JnT6U4TOn0pI0bKZFU8JnT6U4TOn0pI0bKZFU8JnT6U4TOn0pI0bKZFU8JnT6U4TOn0pI0bKZFU8JnT6U4TOn0pI0bKZFU8JnT6U4TOn0pI0bKZFU8JnT6U4TOn0pI0bKZFU8JnT6U4TOn0pI0bKZFU8JnT6U4TOn0pI0bKZFU8JnT6U4TOn0pI0bKZTwsIBJ719iNo7AvpJNU0Q5YREUOh4/1SuNNafDHUPyhUfaOXZb/VK401p8MdQ/KFR9o5AdmIihkec4HYgJdw6hNzeo8VTIrBJKnc3qPFNzeo8VTIkCSp3N6jxTc3qPFUyJAkqdzeo8U3N6jxVMiQJKnc3qPFNzeo8VTIkCSp3N6jxTc3qPFUyJAkqdzeo8U3N6jxVMiQJKnc3qPFNzeo8VTIkCSp3N6jxTc3qPFUyJAkqdzeo8U3N6jxVMiQJKnc3qPFNzeo8VTIkCSp3N6jxTc3qPFUyJAkqdzeo8U3N6jxVMiQJKnc3qPFNzeo8VTIkCSp3N6jxTc3qPFUyJAkqdzeo8U3N6jxVMiQJKnc3qPFNzeo8VTIkCSp3N6jxTc3qPFUyJAkqdzeo8U3N6jxVMiQJKnc3qPFNzeo8VTIkCSp3N6jxTc3qPFUyJAkqdzeo8U3N6jxVMiQJKnc3qPFNzeo8VTIkCSp3N6jxTc3qPFUyJAkqdzeo8U3N6jxVMiQJKnc3qPFNzeo8VTIkCSp3N6jxTc3qPFUyJAkqdzeo8U3N6jxVMiQJKkEHsIXqpVPG/cMHtCQJPtERQp4/wBUrjTWnwx1D8oVH2jl2W/1SuNNafDHUPyhUfaOQHZipnes73qpVM71ne9VEZ4iLT9/dUHyiUUGlq+8XKvo5H111oxdZPNYo9pDISCdjHOcQQ3HY0nGFUpcEeCk3Ai5pjvD/wCTi1UjjU0VbPVXF7quOpr2toh51IC4ilaRKRzw1zmjl0Kvvk4utuipbnQ199uNXQuo3SjUL624xGSEEHcWzDhxOwfWjfz7gtXcG9xm9jHf7m+UXLtvu7a7SF1FbdL066uulHNTUt0rZJC+gkq4nQysY8nPI7S4cweR7VklxuF5i8o0tC66VTLk65NrYqv8Ktbb2WwFrHQug34MmSW42bi4h27kjphdbk/fHwZL+De7+eWHijfyLTXk9N9u/kEkpaKjnprmaGeGgkZUtD5yQ7ZI1wI2ekccyCMZWrbvcNU11XX26r1DNUPc63wG21F2oXVD5WTDe1tPAJC4g4Jw6N3LLt2MDSs5ruTtRaq7tN463Rc8arutyofKc4vF9j/Ltr2ASxAsp4NrZGM/2wRiJ+cnfHuIeeRPZlfkzrHSeSW6VEFVPC+SpuJiqYIXzvjLp5Nr2sYC52Mg4AKxH0X+7r38mWfqu98evX3NuIuUTcKuXTt4AmuPm0YpKWoa59wqI5KtlXCXyvkqYmCJ5Bxsbnt59gWda2u3m3lcsY/jdcqeip2yR1ksTKcwW10wAijeTEQN5Y7+c3EcsEZW9Hilvngp49Y4Gb+De7rh1hib0RYF5JXF1kvRJz/7cuGDn/8AcOWerm1EeC4o3OvxfBhERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBSQ+sfco1JD6x9yjKTIiKFPH+qVxprT4Y6h+UKj7Ry7Lf6pXGmtPhjqH5QqPtHIDsxUzvWd71Uqmd6zveqiM8VNR26jt/H80pIKfjyOml4MYZxJHdr3Y7XHvJ5qpRUhYLLpKjsWnZ7HBNO+lmdUOL5CC8cZ7nu7AByLzjl07VUWrTdDa9NUOnzGKu30lNHShtU1r+IxjQBuGMHs6K7okjrzLHqbStDqmgFJV7o9skUglhDQ8cORsgbkg+iXMGQvj+Iul/PH1v8W7OK2R5ldUChi4jnk5Li7bnOe9X9E2QIRZ9L6cptJ6fobLSSSy01GwsY+YgvIyTzwAO/osUZ5ILYygjtP4Yu50/HUCpbai6DhBwk4gG/hcXG7n6+fathorec3toaTUGE3Hybx1eobje6PUl+tdVcBGJ2UMkAY7ht2t5Picez2rItP2X8AW1tEa+rri1znmer2cR2Tnnsa0fQroik4QI2mIVfk6tdVZbxaTPVtp7tcPwlM4PbubLxGPw30cBuYxyIJxnmskdaqBzKqM0VMWVjt9Q0xNxM7AGXjHpHAA59AqtElxHWzkhBZ9OacptM0lVTUsk0jKmrmrHGUgkPleXuAwByBPJXhEQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBSQ+sfco1JD6x9yjKTIiKFPH+qVxprT4Y6h+UKj7Ry7Lf6pXGmtPhjqH5QqPtHIDsxUzvWd71Uqmd6zveqiM8WOXW/PbM+CmONhw53tWRrXWS6eqJ7eK5ednO3rsbJXHDbPsyKyptK3e2FfBdK2RzAaiU57m8yfcqj+MkbZoI5YLpBx3iNjpoHNbuPYM5Vrp8khrZHRudhoe0Alh6jPJfP4Hrq67zw1V1rXU9G6KaF3DjG5/PP5vPGF52SW1VVLdpW8O9n0ZVZ001JU0ryRdafU1NLV8HfWbHOkZHMWkRyuYMlrXZ7QA7tAyqK4anpamlp6iG43SjZLBHUvdDDxRDG8cjIS12O/wWN3hk9vv9U+ipntYz1S6Nz2Aub6TmjdgE5PMBXhmm5bhZbTwQyF3mcdPVRmV7GzNaOQIAJwDn2r27Kp3FDPgqpUlzp9VUbLdNJHUV9bw6htNxnNAc9xHLa0bfpUzNS05fPFP59SPggNQ51S3HoA4J5E55lYpJapKe210EsEjIxcGh7YwZgQW+sN4yQFU6XsVLK2Zsz3uklikZNEKYRtfGXYA3YB+Kfet3qiXUXoa0oGPjgkkrW1L3sYIpInsd6R5E55Y+dZFuf/AEkn++VrGG3SVNdTPmgqJKo1LNzHvkMkbWnlk42kALY5gdn+dOEVTDpRPuf/AEkn++V46QsaSZJMD+uVAafrI7xVZbqYNzVTP/IR82bvrVTbI0kWyut16iEssd4ZFHPjZFJCXviOOxpDgD15g+C8qae5EU5proWPjbiQTR72ynqcFpB9xx7FXTVYqpeK/k0cmNI7B1969ByMjsRsJFNRsqqamIrK988u4udJjhtaD2ADPYPaSfaviHjCB1JLdHvrNpfxGENeGl3I7TkcuzOFbNUU/GbbZJqZ9VQw1BdUQMbu3DY4NJb3gOIKxmnoquN0sEsWY/wbIDTPj3vLDIdje3kRkKXmW6jLn01RC6Piakrm8R4YxrnQjc49gHodvIq4VMxMU0cdaYZi0hri/Ow9xwTzWrqC3yUj6BsNvfQzy3Gm4cssGG8oSCD2Z9LJVRW0cDrxc53zU0gnn3tIMR5bWjHptJ7QVLzLdRs2EyNbTtfPJI7kHP3EbzjtXw611sj6gm71UbXS7ohEG+gzaBtO4HPME59qioZo5qOkkiLSz0B6OMDnjuV2qZHRQue3GR1W02zDUFBDbZ6V/Fmu9dPG0EmOTYAf91oP0qIVEp58WQZ7tx5L7lqpZmFji0NPbgKFWSGH3LWlwpdUR2uN483c5gLiTuGe3vWSfxptuKWiNxDbg+J5kcZCGxuDT2nOAcgcvaoJYrO2KpknbSi4GbDXPxv/ADcYWHXLyd1daauvinj31NXsij7iC4jJPdzCkssIz4XCR0PFNU5jZIhIwufgZHMj5x9Swm8DWs11p6i3X+qFvkkkYyipaeJ75QWei4yvbtiDSQ7JJzyG13YshqbV+ENMNoZf5wU7W5Hc9o/aFi1ZA6CxGw3qJsrJaZ0sTm1Bj3PDcAZBA5A555GR2LrY13bRSYqpmnAyqxXe6OuNDT3OrtzmvBjfHBXPkmDhz9ICFoz/ALoxzWbVdK5nDdC2V/PDmtkOcdeZWmLWZbRbrbV11JRW6npKjz974GNHEhk9FoiDB6TQ1+CXYIJW7q6N8tPiPO4EHAdtyPeuuU03WoM2blYlEIpT/wBRUD3zD/1Lza5kjWSB7S7OMSk/ejaYu3bqZpwM+k8OXjGCPJjZTMJ7cPAz4L55ZuCnlmcJpGs3OaGlo/KkEO5YPb2dq8Y+Qt9KSTP98qpkY15bxom4dyDwQefTI7FE2j84cBS+i0ci/Ho/5lZxNKCN0rgQ0Plc93qsa4kn6VX0tBL69RK/J7I2vOB7z3lTRw09tjLifTd2uPNzvZ/kqeepdURuY9r4YzghzeZ+cdyqkjgrPNIj3yf/ADHftVueNj6kB0h2v2sbvPQYHb1KqbZAIeMWGIxuII4fIdnTuUDg6Orq5XN9Bjxwx8Z5a0fR9/sVliDyURsAYal7JYxh73Odsccc+eeS+Yctax0jpDG8gc3uDmk8h34IU8bNjAM5PaT1K+ZnMcx0bjkuGMN5lJZIKdgLOLFNUmORjuUkjyAfZ247MKQtAp5HsqHSPD2DLS4AAuAPaefLK9HHDnyTGKIuA9Y9uO/HVfcboC4skfI/jEN3bdrcjswksQeuY1oJL3gDvMh/ao8slGInzPJHIsc4jx7PFA303hkEbXMcRmXL3Dx6r6mqpSYZRIQ3aCQOwkHDvoSWIPmQcB0Rmc+MSRgkGQ4a4dvPPtHgvBNTnsqM+6Un71K6okibUta8bopQcuGfQdg/Rk+C+eLUvaHGQtaexxLW58UliD53MPZ5y4dWiQ/UvQN3qsqj8zx9a8JqMc61mOhkH3BRucPz6pnvEr/uSSwTcOX/APL1OOvFH/qXvms4ZCQSXbSHgzEc+WD9aphwD/72x5+K/cR9JUgp2OwBBTSH+oW/ekiD6kilaQxxLHPB2ubM52PeFU00UdTATmRkg9Fw4jvRd4/OqdvCgP8AM8En+pjK947KaUVAe3Y70ZAD2juPzJLJAMBjlbEZZHFrPSO8jJ8UkYWbXgyODTlzQ93pDx7VU1sRIFRHzLRzA/Oao2uD2hwOQeYSWIIDsicd8ryxw3sfvdzafnXy6M7mSPlmijPqxhxL5Pm7v+exTMPDJiDW5OXROcM7Xd4XkIBBecmQ8nF3bnopIhHjY3udvJfG3uYJCfE/sU0byHbHHOex3VeEhoJJwB3lRNe6SWJzWERbiNx7ztPYqmGVikh9Y+5RqSH1j7kYJkRFCnj/AFSuNNafDHUPyhUfaOXZb/VK401p8MdQ/KFR9o5AdmKmd6zveqlUzvWd71URni1q/f5xUtaMAyu9L51spa57Kmr/AO1P3Ly87VKmzpcTifdkCmtqdh8wjYwOb63bn2qq1XqCSztp2w1UELnguPFc0ZHsz9ypYTmM+wkfSrncrdPebbQMjmY0MLZHsk3Fkgx2EAjPPn8y8fJLSnS3rU+zK6IpV0tNh1PWXC33A1ddT07GQQyF8rcRxvJcXd47doU1g1BcLjXPhraqmp97iaVklO5jqqPsD25d1zy7cY6q60ljrqoXCaSphjqJDE1z6aLLnBvMDD9w788hnkpG2W6Pc3jV1wmYDksdFC3PzhoI94IK/SWTVVCdOo8urBuTEXavvRpKuWSSma2Fsbswwu3c5wwgcz3A93eshud6qm1A81lmbE6Nrwx1rnkIyAebhgd/Z3L5qPJ3Tvpnso6OWnnLmESS1b5BhsgfjBee3B8Vdqix3iaaRzLhUxRuJxG0QkNHQZaStwyNoi05cp7tZqarqA0SyF4O1paDhxGcHs7FdFTWqwVtsomUrXiRjCdpkLQQDzx6ICrhbK0kAvgaOoycfQrDJKPmngNZNsH82z1z9yq5XioeGNA83jPd2OI+4KdsUVNB5ux4bkHLiRnPUqkaJI2BofS4HL+cP7FpIw2SOaNp5DsVuj9QKrMx7HS0w/8A5D+xU8tKynIa+4sjyMhu0Eo0VML44MYnM/Dbxi3Zvxz25zhfGYAede9392Fe4oz/AO81rj/VZj7lmDUnssMNQ6PiMY8wvEjN3Pa4ZwR7eZQwQk5MUef7oXyGUmSQ2udnqQMr6xSd9FUv98gH3oJD9rGAABoBBAHLvVdVjdTSj+qqFradp9C3ub7TMfuKkfK0NJNDC8D48hd9YVRGUaLxvZkNDQeYaOwexerRkwC/aSuVz1C64UzY/N4i17i5+D6PM8vcorxryvtd0ZQxxxmmpJjKWntk57hk93NXfUWtnacrZaMUgl48frl+MZBBV0fTWVtvqDVx07Z62Rkby8+lJzYQPZyJWTRU2SsNwtVNVFoaZml5aOwZJ5Kx0lxhr9US0NbDEGUET44ZHHBGQSDnuKySip4aSA09O0Ngie9rGg5AG44CxfVOn4rhRSeZwsNznlLSQcOf6TcD6VSGIwWG7s1DWRU0N3NJFPJEZ5IS4yR73jDeIOYw6MF2Mfk8963pBM6q07DNI2Rj3U7XubL64IAJB9vJaOv9ru9xujJ3ivlggpWS1EQrJHNZIXPY7bDkNwMAkDnzBAW0PJ9QSW6xTW2pjhpXyOe+GjjkDuFGQByx2DOT869HKvrs1XOJws8G0ZA2Npk2tAa2SF4Pv5Y+9fIdCIonOY0ukAIaG5JOO4KWCOd80IfC5gjJ3k4wfRI5de0KpZDTW5heTgnluccuPsH7F552I6aiLmy8djRHJtPC6Y7z7ezwUslXFGHRRSQiUcmtc7Az0VHJVvqn7NzoYicYHJzvZnu+bn7l8BrhvjYxnDYP5rbzcOo/YgPrD4ncSpY/ifHxlo92Owe9Ste17dzXAt6g8lHExrWh1PI6MHuZ6vgeS+TEXzYkcwmTBftbtG1vbnxwgKihDnTvl2hsb2jHPmefbhK0h9RExksXEZk8J5wXE9hz4+K+KJxkrnyuJ/KR+iOgzySpY19dOxwBa6JmQR7XoD5kBa4Nmc4vIzwovvcjWvAw0Nhae6PtPzryIP3M3MLQxmzORzweX0KZAQbGh+1mA7G5z3c9o6r1mc7omkE/9bJzcfcO5fT4Q9xO5wDsbmjsdjsypEBDvccSOOZGERye0fmu/wCfavhzQ6LYfzZCw+xrglQ2T0mxhpdOAzB9mcEeJXjp2xiWJsbZ3OI4hk9QkdwHsQEkL3Gpp5C3cKmDY7lkBzefPxIVHl0xL5fSfkg57Bg45KY1NSW7BK2No7BEzHLpzz9yia0NGAstmkjzhsH5jfBfQAHYiKFC+HMja0lwa1o5k9mF6+RsYy44WKa7rp4tPyCIFollbG4jtDTk/cB863Z0X6lSZqquqS4O1ta6aYwMuILuzmC5nir02vkdCydkdI5snY5rSQR17VidFp+0/ghlA+hLpJdpNQfXBJbzHuyqbyf1E77ZXUkhcYaacCPPtByPmwD8671WdN1unYclU5hmw7ZcnTyugm2h+MsLRgEdF9SR+az7B/NSklnsPaR9/irI7cC17DiRh3NPtWQwyMulDzy0u7QO1rh/nzXA6EUjOI0jOD2g9CowZHHixxlzz6MjB3O6+4r7jhqpvQcBFjk5/bn+6P2qozT29m0D0nc8Dm55+9AfMVEXEPqCHEcwweqP2r4lqmTTQsiBc1rzl49XO08h1Uchlq/530Iv6IHt/vH7vrXpAEkAAwA7s/7pVRGVKkh9Y+5RqSH1j7kYJkRFCnj/AFSuNNafDHUPyhUfaOXZb/VK401p8MdQ/KFR9o5AdmKmd6zveqlUzvWd71URni108bayrH9fP0LYq15VcrjUDq1pXlZ4X+in3n25vcWrXdyIKc8pB0e761fbY6d9FHsdHtBI9IHPaVYYeU1QP6wP0BXqwPzDUsJ9SU49xA/zX56iqGenlKmmS8UD6yOSoEZg3Fm4Ag88Z9qrYZat8LXuqGlzhn0WDCpqJ/DuEB+OHM+jP3L6hqWRRiHa8mLMZwO9vL7l+lzbaXrDwZ41soqPG1lY9uTUAe5gTj1J7ah/zYCjj9UZGPevSQ0EnsC+2TEB8s/IComLicBod2lV8k77bSMY+R0tS/ONxzz/AGBR0rG00Tq2oyOXoN7wP2lW+WV88rpZD6TuwfFHcF0opk511QRuYHkuk9N57XO55XghjHYxvgvtF3OEnw9obG7AA5dyq5x/tEWQMOgaVTPGWuHsVRMd0tIf/gAfSuVrqOtke4wiIuJ2CIiAL5f6jvcvpeO9U+5AUTObG+5fS+WeoF9LZgtVVpm2XuqllrYeJK3axp3YwCD96w6+aQuN8uNXW0xa2kp4Wu3Od2uawHaPbhVWu7Zda+vpBbYp5BwzvERPI55ZXtw1nLpq1xW4Uu+pqIWSOc88gCzYRjryWWaMg0nSVVvs0dLWMcyphc5r2uOSM8/qKsl0iqrZqeovfCc+hgYzPpdrm7X48GrIrHfI9Rw1FwiidE2SYjY45Iw1oVtvl0gNWbFNE8isBmMgOAAI3Aj6FdhNpjr9Qy3OCruMFLXOjFWWPdTVogEDX4cN/onI9I+l3Y9izLQEkDLjLB5jFFUy0jKnj8Z00j2ucQAXuA5cs9gWAVVHYbXedU2a80kzrLSy0jqaRj3hzJJQ8HmHDOAxnXCznyd1FmqLtViiiqX17aZhmqquR3GcNxG10Z5MIx3L0a1Fjq2L0XecF+c2BNVtinig2udJJ8UZ2jqfnVua48V7pg5z2uLC8+qOgHRSmeSpdJ6YjYHubtZ6xwccz83dhfPD4IJjbuaRh7D+eP2rzzsHs27vR3wv5uZ3g9R7V44bQw78tz+TmHXofajHiIN9Iugd6jz3f1T7V7IOCHPADmO9eM9jv80Aa173O4Ya2c+tGeTXf1ghaMPjjBc4+g+YnxwF66IQkxxPeZHj0nOPpRt6D2qRrQ1oa0AAcgAgEOG1sQHexw8MJUDbXuJ/Pibj5ic/+IL5J2VFO7+vtJ9hH+QUtcNtRTP67o/HB/8AKgIGy5cBtIa4kNPXHapVTlwYwuJ5tm8AR/kvsVEbvUJeejAXH6EBKo3+lIxhcWNwXOcOWAP/ALhe7pDzbBIR1OG/WQvYoXvnBeYmgsc0ND8k5x3Y9iApGTPFMJjniv8Aycee74x+75lG1oaAB2BTOip3R0zWVLyyGPh5ij3Z7Off0Xggh/8A3ko9u1v7Co1JU4I0U3BjaCRRhzek0xP0c1C9jI5BwwWskY14bnOMqNFTCgfUZ5R8/b3JUtJaHZJaPWCibjI6IkGygqLxbaWZ8VTcaWOdvJzZJQHD5l7dqq2VFsZT1lRTR09Uzcx7pWgPHIhzTnn2gq3XWqvDZ6ejay2EV7pIWSYkyzDCcn5grJfPOLLbrVTOq4HuoKQxvp4p5IXS42hjgdpycNPTtXWlYqDDZ8fgFxY6pj1LH+D2O2mXij0fZnOFk1iktNNSmkoK+mm4YM0hbK0k9mXH6Fb9OOl8ynbb30BnfKJJRLPJMOY5c9reaks1BetQTfhprbZExsFRRhjmvIP5TBJHfzjHzLdddVShsyqUsUXVt4t80b3RVsDw1/CJa7OHE4H0ntWXUNIaVrnyPBe4AED1Rj/ntWl46FkLYzFNUSUstWzMz2tjBlEg3ei1+7Zy5At5LcEgkq/5/Ai7oh2fP1+pc6lBpMrK2SSKnc+PGQRk4zgdVRxxtb6WS5zu155kqehky11M/mYxyz3t7vDsUDWmnldAewc2Hq3/ACWSki+XfzsP94/+Erx8gYQ3Bc89jW8yVGHf7VC18gMuSeG3mGjB7T1VRGVqkh9Y+5RqSH1j7kYJkRFCnj/VK401p8MdQ/KFR9o5dlv9UrjTWnwx1D8oVH2jkB2YqZ3rO96qVTO9Z3vVRGeLXM7gawEd5kYfeHE/etjLWkjiZnnAxHUvH+9n9i+TONnfySt7oZ0yWu7lNC3yuB9DlUP9rR9Z/wAlcLE/bW1kefWa1wHiFQP5TRnqC36j9yqLa8R3iP8A+JGW/OvyS1nv2qmhmSl3DfFJ8R7XfNnn9Cke0R1tZGO6Td4gH71FKziRvb8YEL7fJxamOX+mgY8+/v8AqXuZqrwqp8GePbrUz7U1HTipfxX8oIz3/nEfcFHDC6ql4TDgDm9w7h096rZnNfimiwII+T8d5+L+3w6r2Uj52yir5nVEjHZ/I/mDr7VSqrrwfyeMKj9L2L6KNR89es9RfOHfG+he4PVaMnp7FI7mKF3WJw+lQ7f6xUzseb249+Hhc7XUdLPWSIvNw6puHVcDueorLcdXWK0VHm9bdIIp+wsAc8t9+0HHz4VzpK2mr6dlRSzxzwP7HxuDgUBOiZyiAom9nzr1fLe/3lfS2YLNedUU2mZ2SVEcj+MwBuzuIcD9SitdptGpZobtWRAxRwMDWyHDcOLxz+fCkvmm6bUUsDKl8jBECW7D25ICxPVFir45KK02kVEkILoCAeRDZPR3Hs7T3rLKZlZLPBY6fzWmLjG9jJ/SOebm8/qVHfbPT1UlTcXbxU0kDWsIPL0y5v3qh0KLhFDXQXETcaB7Yxxc8gMjAPTkpbxemTzfgul4r5zWxGoDWHDY8Y5nsxlXYNprW3QRVV9v9Nqad9PHBR0VaZZ2uc9jmzyjljJyQ4DwWfeTiW1Qauc2idUy1VZSPdNPPC5hlLX5BJIGThwCwmv1C+6eU91Za4pHsntsUU7HxB2Qx/PLefIEjmst/g/3Kou0epjcJTUVNFXujjkk5uY12TgHpyC9Of8AQfclxw9j54+tG0HzMdJLvipmkPcOZ2uPM8/nX00NIy2Fw/rMlz9a+5ppWTy7jhodyL4SW4x1/wA1C18cpzw6OY/1fR/avNO5NA2GN8mROWvb6bXgFp/zXjWNZFHIJjLg7Yw4Y9LqfcvjDGlu+CRoc4N/JTEgEnHYcdVJMYA9rY5XsMILBmMub7Tn/NAfcbNg5nLjzJ6lfSjYZHDLDFO0d8TxnwP7VHLvLJpHNlj4TQRnkCfqPZ9KA+6gEwuLQS5uHADvwc/cqurLJKN0mwSBrd7RnGVpG/auu15rJhRvmjpGuLWNhB5jqSFlWlbzebRXW+3XpwfTXEObDlwL2OHLB+fkvoqyeqlS3ic1aJszuN0gYQ1sLAeeWjOV6eM71p35/q4AXxTjawxntiJZ4dn0YUq+c6EfAYTl25395xP0L4exkAErGNaYzv8ARGOXf9GVIyVrzgZ6jIIyOo6r7QHw1vDlljHYDub7nc/ryPmX2qcMILIy4NkYNsTz2OHxSpWPyS1w2yN7WntCASEnaxoy+Q7R/wA+5Ucrg+YBhyyFvCB647SqyRpc3LTh7Tlp6FUk4Ae2Vow2bOR8Vw7Qoyo+e1UbmcJ+3u7QqxRVGzhkucG47CVlOCtFJJgFj+FxHsOW4AJbnkSPmKkEry5sce5zz2NapqS3VFa6OQZgjaciQj0j7h+1XympKegbtjAD3nmSfSeVsyUVJaXEtkq3l7hzEeeQ9/VTCWKmZ5vQxMa0EnLRhjSTk+85Xw6aStzkmOHONg9Y+8/cgIBEULA5/c0cgPee4IClZaaCGc1TqWnNTncZjE0Oz1zhVRLnxvkzw4WjJeRzI9gRvDbNtmPElzjs/JsPcPaf+eSja6QEyuLnvblsrDzyO8AfV/mgPprpIwXQMDAOZDubn+8qsexlfBG9jtrgdzXDtHUKlZiNwjDtzCN0burf8l9QyeazjP8AMynB/qu6/P8AWgKCapLXSRQAsAJa97vXcQfoXlE0NqYgB3n6iqu70m0+dRjs5SADtHVUlHzqovn+orrTF051TJeFJD6x9yjUkPrH3LmzZMiIoU8f6pXGmtPhjqH5QqPtHLst/qlcaa0+GOoflCo+0cgOzFTO9Z3vVSqZ3rO96qIzxaslcBVV+fWEu4fM7/NbTWqpCfwjWNABL3SD6z9y6Oz0mT21H/E+e0r0dtZVd5WTcuG7o8fTy+9fbH8Kvo5Oj9viFE526l3dG7vDmvajIEbh2te0/Svwa1n62pSoMwXwxksgpWxNy5rpI/YOeRn5ivY3bo2n2KnqNTUenGTGqcS+TDo2N5knGD9QXr5pprryhUUKZR42UNU0TU9RfXAUcbaeH+ddze/p7f2L5bshaGAhoHcSsbsmtLbdZ+CHPiqJD/1v5x9hV5qmg1PMZ9AfWV+jtbKuyd2tQz4qK6a1NLPa17XBmHA8+4qkSVoa9mABlFqzcoxaKGERFswU1bWRUNNLPM8MiiaXvce4AZJUVBZLzeKOnqKmu/BkZBdFTRRNfI0H47nZGcdwHLqVadZO/wDYdYO4twfELZi+K2WktdG24STwca53Y4RxPrsvos76WLcb9UczFP4oV3/6iq//AJEX/pXjtH1zmkfxjrBkYyIIgR/wrQ+pdRang1FrJujrtf8AWen7TaaqC6NnrKWOCnrJAS1sUgjZuMLGuLmguILmjtysY13rWm/iPoGyyaqummKs6coZ6nUENfdGila6NpOKembw5pCB60jxjPuClOR0NTL2f7qtt7v7p708Dbt6k4w27Fsju7/tBvuHyJ22F0zjc6qV8oIc6WNjjz7cZHI+1V9l8lo0/wAYW+/V0TJcbmFkbhnrghaU0/ru21vkd1S2o1PfrbcrO6OTz+iu1dPNXzuYRE2F1dHvaJHtxwW5we/nkvJbeKqs1N5Joq7Ut5uV/D7rFe6O6VTXy0VWyma18RY0Dazc0uaHZyHZBOVv4CiWpeHfVunf1K3mPi64WCx7ly/zjuN510Nz00+GWqqWV1skkbE6bhiOSBzjhpcByc0kgZGMZHaOYu7H72gr3XJA0jdycANgJye7BHNUtBIJYGPa4Oa4Agg5BHVcLNOi0qs5lQnjjrnkda3eoVcY4ryjmfLgWPIIIBJwV6qqRgkYWlUgzzB9YcivpTODRjetXV8VqbNbpZWTseBiPtIKt9p1WbFZKequomkq5TJycPSc9sjXAHplZhUcoi7mdmH4Hfjn9yxq60tu15PPNBUOFNTGSbLW4Li1jMt59naUYRdLPqOm1C2SqgDm7G/lQ4eqSXHHt5FWUsg0xc6qpqpnvfdmDDQ31H7gQPdzUtyoINGWu8soC85kDWA8yCWj9qqrVbqi6WSqqbjI10/mhALh6W/G7l/3Q0KAwKGOj0P5R7cZZXvguVillJ29j3SMwPcFV/wdJCy/a8p+5tY1wHirBPSVGpfKrQ0FbUPbHBTVkELtvqtY9hAH+6r95CmeaeUPXtJ3h7ZP+IhenRjYVf8AVcKjhV+dePsbsLpnVVWBPI1rZAGgYxjY0946ko8SyDD3RSNHc+PK+5MNq5x12u+jH3L4jm3kDa5uRuGe8dV5x2PIWtFTG11PE3tcOGSMY78dnReU2TCHE53kvHuJJH0FeSEg1LgcOjgO33u//wCUqpoqOlkllk4UUbebwM7R1wgJHwxyEFzGlw7DjmPnUW53mpjD3bZXvB3EuIaBg9vtH0rBarUsdHU0JptST1eamNs8TqcANiOS4nDARyH0qjvOqKqnuddRxXOKnNK7gtZNPCweqDyDm5I5/WtKlkktVxhu+lhNRwU3+zPcXxztbkkEjHP3ABXjS1tumpdSUt3uMZip6XBZkbQSOYAHv5n51d7nfaxlhjbTbZbpNCx3H2/7PGBjLjIfRwewD2pdNR13mttqKGjqYXU8zYqiKWle4x+g8n1RzGQzmPvXd2za1Y7zmqMdZmc7eFWE/mytz84Xp5Dn2LB7HebzcrlRGqOY/N5JZhJmMMxIQCGkfF645KiOsay42kmFshP4SihdOwtaGROmaGgjOSHMPaB3rhdZ0kz5hjc9kJc1zHH0drvSYfZ7F9wOL4mud2lexSOjm9IRua1peXbAHDHuXkAIiZkAHGcBZKfbmhzSHDIPconMdgNc0TNb2EnDx7AVMiAgaeeGSEO/o5uXgVG4GUzQFpbIfyjWn4w/apDK2oBZEwz5+L6o97uz71WU9I4NYZyHPjOWEfmjpnvQFmiMtWdtNHuPe93JrVcqa1Q0/wCWqHcWQc9zvVb7gppK2OMmOBvEkHaG8mt957vrVO5j5zundv6MHJo+ZSCyTPuTQ4FjC6IH0pOwD9qo3xhjpeM4unB5SHmSM+iR/wA9yqsDGMcuipJvQEQcS0xOwyTI9Xp1yqQnkiIj48jjEHAcRoGTn2dMr4e4OgxG10LGOzKxpwcHsdkf89q9hic5+djxC8YkdKcZHz88r0BlPOculmewbCxrQBggdpPagI5PyjHPd67cNlx3jucF9bzgTHtBDJfud/z9y+otsUmY4ZDy2/lH8gPYBlA805/moSyRzWODWkE5OOvtQHmzB4WcZO6M9HdPcV4/MsREoMUR9F24c3ewBfckLaZ8bRmV3Ph7j6LcdepXy+QvcJ3gb4/QkHc0dzh7EBV0tS2bdC5jmlo5BxyXN6q2spTR3FkePyZy5h9mDyVTI13oyR/zjDlvt9inklZUMppG97z83onIVTgjR9qSH1j7lGpIfWPuRgmREUKeP9UrjTWnwx1D8oVH2jl2W/1SuNNafDHUPyhUfaOQHZipnes73qpVM71ne9VEZ4tTVJa25TF3qiYk+7ctsrUdd/06p/7R31lehkCl1J7jzs4OFSy4UvOma088ZafmSXLqMnvDc/OFCa6noYw+rkMfGeC1rWFxw49vLuyfoUxqKRscrXVTAGNa5ziCG4eSG/PyX4SvN2UpuKHEv1P11GWWNVKbqWwyqhcHU7MHIxyWHyUkty1vUU0shiaWYY/AO1uO3mrza75Qh1HRCVzp5QyMbW8i4scRz9zCqa8Ost3mY19U+KVruF5zGx2wHOC0vxt7favdzOrTI8oq0lDSqWuJg8bLUrWhXWsH5ll1PbW2yTEUrJZ2SNDJAAH5+buWxCXfkHScpDEN3v71iVstFgtF6EVZWz1dTHII9zozwmyHsaXdmVkbdUadEDJ5XcIPbK9okByWxv2OPivey2rS000Uy42x1uPiyem43U8J2Ek7m7ozuHf3r54jPjDxXtJqqx1RqeFC/FM3c9xhzgbi3uye0KSDU9lqqhlPCwOneCWsdA5pdgZOMgdy+OmiulRB3qipzJHuHXKTMqpaKcUL4Y6s44clRGXsHPnkAjuVyFfHgFlG3B9oH3K23qip9QRwxVUM8bYnEt4FSY+3rjtWXaFVnBrzX/8AGi2WGWaprLRPE+RkRZFSOY47nAciXLbdmvNLfaJtRTuG7GJYXH04X97HDuI+ntHJanotD2+6ah1PRV/n89DQyU5p431T8N3M3Egnt5qPV+jY6uop4qKy3Sc00O1lTT1DGl2e52Tk49q+aumu/fo1xDT4Y473sO9FVN25XwNqyaXtY07V2CkpIaC2VMEsBioo2xBgkBDi0AYB9Ins7VZrj5N7Vc/JodATVFa2zG3MtnGY9gqOE1gaDuLdu7DRz249i1HpzRNdQx0zr1Zbs+QA8eSSoZwgPjYBzyCx3XtpqXNbc7UyqoII4BI6N7+UkZPKQD5xlfNlWV2+T2VVro1VGLSqxw/8dks+jJrCytrWmi+1ubW/77YR1JUWyjrI6ZlVSw1DaaRssPGjD+HI31XjI5OHcRzCs1y0Ra7nqux6lfxYrjZzUOjEO1rJjNGI3GT0cuIa0YwRjHeuX9OWZ1eKeOpuUs0tS9mODUbg0OYXgHB5HAOemMdyza1U1vpqKlmqbe9/HOyJ76ghp9Ek5z3hXJ8qyi1s1aOzVM4w25x/8TFrk9hZ1uzprlLCUsPU2p5Ub9SUOma23GVrqysj2CJvMsjyNz3AZIAGefecBT2GN8Vqo2vMZIiaAY87cY5Yz7MLGNN2K0XBkk7bZStLHbSQ5r93vx96zhjQ0ABdrOmq8669bjVuX+TFdVMKinUj6VPURkHiN7h6Q6qoQjK6nMo+1a8vNHV6Pt7qm2VDy6uqZGOZsB2DByPnGPBbDLeG8s7u0e5Y/U17rZq23OqpNtofHJK/c3LWPDCNx5f3R860zKKWxagE8NO2um33WshdJ6uMknbk4/qhWK6X68utLq6ifFDRQ1UlOwNGXvDhzznuAGFe4LNFTXBt5qJ2tNRAMMd6LYg88uftGSvm26hoRqGaAx08Nlt0znb8cnbgWgn5woUwvXOoBT6k0lLbZ2i4wMrWSgNyW7iHDxC+vI5UvHlj1Wx3bPRxyH2ndn715XWuO1a70Vf6ieM2+oqnseCOTR5sXZKeT2Zjf4QN64RaYqqhMjC3sI9E/evTscbOpf8AD0qPnr/MvH2N81hjFU88fhu2Na7czLQOeOfLqvZIXRSNDXRuDIwzBdhy8q+C+pfG7igu2h5AG3n2L4qGwmplPEBdnmHREgfPhecdijqLnSU5mjqKmGGcvYNjpBkNBB58/eql9zo4o45HVULY5Rlji8YcPYsXvbrhTz0LPMrRMyvqTBFK4EuHoPILj7A1Y/faaeyWi0UFVVUrpLfA4PiicQ6QY9Eglp55B5LSpkkmbUTKM3G5XZtc10U4jY5jCM5YOuew59iop9Nx1NZTT1F6uDZGbpYAJo/RJGMN9DnyKxOB1oZo+/iokiEtTE6RrTLu/KcMhoPIc+XYFVV1wtLotGTGtpjURvia4h+CxgjIOf8AvcvmWoZJMjqNP01yhNHUXu4vi2/l2GoZgAc+fo9cKOa20lbaH0jr5Uytr5GxxumlbINzcu24AGcgcx3gLW13q6GiqIqSW6UcbXCcPmiJMs5Lg5rHN7XEkY2jnjKqBeKR89HVZrYZJa2OeVlNbpo2wtbC9uQNvaS4ZXWnJ7WpXqaW14GXaUpw2Z5ZbLZqKW4ed1lvdFLFw5KamiEMTmjmSWhxyeSrK2zWu4wx8N1PBMHRPZM1oy+GJ7ZAPdjHuytbVFxo6ye61ZlqIp6hjmGOK2TsBAadux2z24Oe3Kmrb9DXOtpM9bEaSkEDWx0FRzaQ3iBztnJxIbjHL0T1Wvhbf+x+TJpaN6NqR3e31MFTLHWQOjdsYHbxjON31EFVTbjTyTtgikbLO5nEaxhyS3OM+7K0Pb7rFS09vbNTV/5KoD3htukkwPNmMzhzcH0gQsrs+tqOj1GK2RtzfTCi4LGutkjNh3k7cNbj2qPI7Zf7H5MqtaN6NrNp6mXt2Qt9vpO/Z9akNDTRjdO7iY75jkeHZ9CweXynUk2Qw1ULfZb6hzvHYqJ2tbRI7dM+4Su6voag/RsXJ5Pb9nV5M6Uuh66l5mfT32liBEIdO4d0Q5ePYoY7jJUES8TaGEF0TR3e0lYWNdWgDA89A9lvn/8AQvBr22RSF7PO3Bzdrmm31HP/AIFlZPlE42b8mbbsowrXmjOWbaffESGtYfR/unmF9NlMn83G94+NjDfErEGeUSzmCn3Cr4wbtc59uqCRjs/MX07yiWiVrGyy1rsPOf8A2fUDIxy/M6rp8Lb/ANj8mcdLR/cjKnPxkPm5j8yAbj4n9i9JELv9ma1o4bX8R+XvIOeWSc9yxhvlCsbRgOrAOgt9R/6FV2TVdpvMsVNSVjH1UUAEtO8FkrQMAEscA4A8+eMKVZPa0q9VQ0vBhWlDcJl7mjbVPJk5mSIFuee3IwcdF9NeXSRyHlxogcdCO36x4L4jOBAe/DmO+Y5+9eNe3gl7SC2GY+kPiuHPn7CfoXE2VKiqDthL/wCjIf79pz9yecRn1Xb/AO4C76l9YlkBDYH4Pe/ACAjeCKmeFzsNe4PjPxCR9RIKbyDxdmXNy2Rnxh3j/n716+La+GJ/NzoBuI6jHf8AOvcONPFVdrtv5QDvHX5kAj/Ju4e7cwjdG74zf8l4GuZUx4P5N7i4jo7afrXgbnEQIAJ3xO7g7p7ivWvc+SHMb2+kclwwM7TyVRGVikh9Y+5RqSH1j7kYJkRFCnj/AFSuNNafDHUPyhUfaOXZb/VK401p8MdQ/KFR9o5AdmKmd6zveqlUzvWd71URni1HXf8ATqn/ALR31lbcWo67/p1T/wBo76yvRzf+ao83OP5aSlro4JqakMskbWskMTuNFxI28QH0yMjm3h8j/WX1SWtm+sdQVcVWM078uc1riWOOckDt7OarbbM+OoDWktDmnJBxk8sfUVc3SvdJtc9zgRnBOV4WcM61ZJbV2CpmH9sfq3TxPYyHI1b2FFrOz0wLHHHUDVVDWSwyN4dQw8OJxkB/Izd+Bk/tVWaepafwY2G7jT+0u4ZovygO7PD3Yzt5nn292VeLXUvhnqoy5xYdj8Z9Xu+4/Qsj7V81WeIpoijYtuqJ1TOKOleSxXUr231MA1VUR26/wyU7DG2RrKswzZEZk7AS0yNGRhU1RaoKvSturZ3SGUOmgPBczhhskxcXOcWv2gbRzHMZwtkOJ82pX5PLdEfeDkL43O6nxXvWGVX7KipLcefaUXamjAtN08VNJd3sjqHiTG6S2kvyS4uyHgA5GcdiudPTudf7ZUMgvJiibIxxqw92C8AA8+QAGcn3LKcnszyRbqtZbZEowJo/5tvuX0viH+bavtee9Z9a1H0XuI5ntUT4mSY3sa7HUL7RAQPo4JI3xuibte0tOB3FYZf9CuuNhqqCerlqxLFHSMDGCNzINwzzycnHf7OxZ0vn0Wytc5m9uOzGUI0moZpzTX8HG3WbVEN2t2oKv/YHiaOKWnY45w4AFwIBx/dV1v2iL3c9EUFusoprbd4gH+duqpIHxv3el/NtOctGOZ71ty2yCSqqMNLQI2cj73KhZ+eOj3D/AIit1VNtVM+ezyWys6KrKzUJ7jUHkx055SbDqB7dS3PzqyvjeXA1DZiZBgNwSNwHaeXTmsu8perLto6y09dZ7O+5zOnDJY2xvcI49riXEt7OYaMnlzWZojrly0SnJnZ2Ts6K34vFo1t5MfKhV68q6+hrrE+21NFEyRzjISH7iR6paC3s6nPNZBrHyg2PQnmP4Zmlj89LxFw4i/1du4nHZjcPFZSrLftI2PVHA/DFspq0wZ4Zlbktz24PzBJpdUtYFuW9NjdVSde9r2XcU1h1hZNZUk1VZK4VcdM4Ne4RvZtJ7sOAKrZnU8tXFRyyMBrI5KdzCRu2OaeYH94NHzqlsOj7Jo+iq4LJQtpIqh/FkaJHP3Oxj84nHuHJa48r/kuptViS/uuM8FVSUzYWxhjXMIDyfYfzz3qpUt68DNdWUUWSd1VV7k4XEye5Wm6XTTsdNLOySoDuJO93ojhtO0AcvdyVTWacpa2zGjoBBHPURwRSSjsMm/tOPYVh3kyjqbHazYG1Aqw5hqX1BaWuBOSG9p/qrGrtbfKrpme4VVlfC22RyCUxEwuIc3BJO8Z7c9h7kVKbhMWmUVWdkq6qHL2LFouuujWMuukdOPdFwm1Iw/B/nOC+Ht6eivdKUb7D/CCoKOVzXSfg18LnN7Dhrf2L5u9wqqrR+lb3eWw/h2S4wSuIaG4JnfuAA7vTC15N5XXUXlgp9S19scW03EhLIX44o5jcMj3L0Mni7D/tqXE45RlVnZU012jiWjsWvBNTVAdvAYR78vUmR5zMB2HDvEf5KnjqRX01JXtLYRWUzH8KXm5oI3Y5dN2FhtH5XtDV0jY4dT08U7QInMmhfGS4d3pAA/MvPSb1H1VW1nRF6pKdUmZyU7JjHE+ON0kTjLTue0HDueR7Dgnn0PsX2GMli9Boae7I9UrwtY8tz5w4ggjkGjK+pA2WRzjS4c45J4h5+Ch0PC3JJ82hD3d7nlzR7dv/AD71BURMYwNa1skzzgFwHL5u4BT8Mjshh+dzikQDmzngxscz0dzO/lk/cgMb0Da4KmtvV/naJq59ZNRwyPAJghids2t6bnNc89dw7gMV2pvKZpbR1eygvd081q3xiVrPN5X5aSQDlrSO0FfPk5//AAW4fKtd/wD7D1g/lU8qV8sV6tNssGmNY1JpLrTPuNTQWR9RTz0W0mVrJMEE82jlg5BGQvpyiare59vtq4HOzws73dJl48sOiXafjv7b7G+0SVD6RtQyGVw4rRlzCA3IIwe0L2m8rmka2st9PS3WOeKvts13jqWfzTaaLZve4nm3lICMjsytJ0OnNT3rSVNDpu26rp2fxorbldaaZ8lgnqoKgSvaxji7Lmtc+ME55lpVppfJhV2XXNJaaSzW201N20VdqKmt8LWPmj28CKMT1YP5cnLSCQNgJbz7Vxur7xq77jfqbTx7p1/+aXobtrvL/oCloI66mvkdypZI6mRslvbxmkQNjdIMjkCBKw4J55VVcfLXo61aoZp2prqlteS5r3No5XRRvBaAxzw3G47uWMjkea5i8t1v1UzS2iqK/WaeyfgKxVtKJKOtZUCpDYqRh4hDMMa4g+jnJA7Qtn688mV7d5WodV23T10u0DZqSuNRTvoWTsfCzaIIp5pmyRQux6cYaQ7c7mMlbqooTWODng2k/IwqqnRMY/4wNsQeVvSr9JWzVNRWy0touUj4oHzU7y5zmuc0gtYHEc2OU+nvKnpLVdRXU9mu3nlTQ0/nU8EdPKJGxZI3BhaC7mCAACSVovR98n0H5AtOS6Hs17qLteK7zqaShsxqHOaKwNqDKGcRrXcIOaDnnt5c+y7eSyfVOqNQV97uxvtr1hW2k26Oa56ZmipqZkdXNIxxcdjHlzJIxtBafRPM81HRi+6V5LipwNOrBPw4vhgbQh8s+kpbhbqGR98pZ7jUspKY12n7hSxyTP8AVZxJYGtBOD2kdiqb95WdLaaudRbbjPc2VVPjiCGz1k7BkAjD44nNPI9xK1RPp/X3lA1/cNN3LW1qZHoqrt13gmhsJHnEr2Sloc3zjkG4PfzyOzGDg/lNsVfU+Ue63Gis96utPdbiyKqbWaSuczLfFHDsL4XxVEbZml8YOABniZzgc4qU2l4+0er8g21P29/aDomfy0aEpKC2V9RqKngo7pQS3OklmjkYJqePbvcMt7Rub6HrHPIHBVPafLnoK+R3uot2oKWqt1lt8VyrK6FwfDFFJvwCW5IeNhywgOGQMZOFpLyl6SsjvJlpK7zU9zovwfFJp5sbrcy10kEM7gHyzNrIqh9NEOG0CQZcC4YPPlavJ3arcyweVOpprjVT26gsFJTy1NFXR36mnhgZM90UU9REI3OLctdFwtrNwIPpYFdNN2qrdPvHC7xCbbpW+PZPje4HQ1P5ZtKVTqaOKa4CpmuNPazTVNvmpZ4pp2l0RfHM1jgxwaSHYwce9XHylUDJdJ3G6xgR3OywSV9HUAelG+Npdtz8VwbtcO8H3LkfycUttb/EOKl15p6pnqb/AElzq7fJLBFO3bI8QsAjp2OfII3xtIc8gO3BoAwuxPKFy0Dqr5LqvsnLvY0XLWiP7o+2HvP2OVVV6lru4y/aD7tdXDW0LNtO55kYZmcbB3HHsVdBPKXsDnscyVpwAzABxkY+bP0KyaaDm0VCxrHl0YGeXcR/mrsWCme1sksbGRPy0Dm7GeX0cl8lShtHZag2SeSHiOmcDtyWMAaM9OzP0qiJfM3LtpB7N+XkeJU9XPDTTuiEc0mfSwHYHNQ8SnbFuFA4YHYZTn61hs0irhOW0fsgI8CAvAXjzd0bsGNzxjPJ3McipZeHHUAAhjIWbMZ5ZPP7vpUfqxu6RSh/va7/ADP0LRD1oa1wMNOWuByDMchnsaAf2L7LXceFz5HvcXHtPL1T3dikXy7+dh/vH/wlVEZUKSH1j7lGpIfWPuRgmREUKeP9UrjTWnwx1D8oVH2jl2W/1SuNNafDHUPyhUfaOQHZipnes73qpVM71ne9VEZ4sbk0TSVrRUNqJo5JRvI5EZPNZIsfh1pQ0/8As8zJWGH0C7GQccu5dbLSLGzn7HG10bhWkfcsNz0nLZRDWCqbLG2QNI2bT6WR1PeQqV/KSM+8LJL7qG2XGy1UcdS0SANe1rgQSWuDvuWNzcmtPcCCvzOe1W8oVVeto9zNdxWN2jUmSU5xWlvdJC4eB/zWTxP4kTH/ABmgqxWmXgXm3vzgGQsPuLSPrwr7E3Y0s+I5zPAkL4Imypfe16MtuotG96RcLZGyogqIJG5aH5x718zWeVmTBIHj4knI+P8AkvbQ7FTO34zQ77leF+kzfW3k9J5ttSr7MXkD4XbZmOjcezcOR9x7CiydzWvaWuAc09oIyCrdPZ4nZMDjEfi9rfDuX3qvefO6Nxbof5sfP9a+146KSlwyZu3nycPVPzr54gz3gdme5catZ2Wo+0RFChEXjWNknja4ZBzyzhAVVs5VkvtjH0H/ADVI31pP+0f/AOIqvpImRVzAxuN0bs+Lf2qge9rJpmk8xI768qvURaz6RfHEb3ZPuC9DwTjBB9oUKfSIiA+Jf5p/uKttZFHNSlkzA+Ilpe0jIIDgSFc5P5t/uKtddJw7dUyAZLYnOx7gqiMtklLQU1RVtstJHJWwkxPjaNhc8vy1uTjsAVY57q211E1XAI3Th0kkJIcB1bkdqtN5uJ03LV3RsYkkq6mKZsecYy0csq8ytMlolGMF8Dj7iWqohrjyzWdsFtt76aKNtPbrnGdnYA0yNdgfMVj18ltNR5Y9Dz0LI5LfuET2GMhvpCUYLSOfM9FW63kr6LyVVU1W2TzuCrY53GyXYJbzOfYVSa5raSTykaJqaN0Rj3UjTwgMB28A/wDjXpZHjdW+96I4Wu37HSNXA4TU7ood0cbHMw0gYyW47fcVrSfyA6NlnpquGztp6iN/EcI6qQNJByBtJcMewALZdxY18tMHtDmkuGDzHZn7lRSQ07ZNpjgjAGdzo859y89VNamW0sbO1jSUpxvIL1RyXO019vNZTUb6uCSATCTc6MuaRuAIGSM5WmNM+Q+86U1Fa6ym1tLU22jlY98MLJY97QclmGucMew8j2FbxbsYMMmy3/4UP+a+vRPquqnewhrfuVprdKhGLbJbK2qVdaxWrFlh1nU3uHTNwdpyWGS97WimjkY2MklwB5yODeTcnn06rAfJ1qLypVOooLbqixRQ2mbivlq/NyC0hpIAe123m7HaDkdi25tafWp53f8A8uPqKr6QxCiidG3ZEWbgCc4B5oqoUQLSwddorRVtRs2M1LqXWjfI1X1VwroX1enLrUB7oYMcemqC3BLQSA5jgzJGQQQTzzgfen/4RmmdUVUlLabTf6meNnEc1tLHybkDPOTqQs/raKjuHFM0cr2TDBaWB7XD3YVuotLWahmqammo6emndGOJLHStje5o7ASOZA6Lvp6Kl9dEvfiR2dqq/oqindHuWG7+XOyWBkUl0tN6o2TO2sdPFC0OPTJlUlr8tdnvjZHWy03itbFgPNM2CQNz2ZxKeh8Cmr/Jnp7XMNNFeI6+QUxc6IxHh7S7Ge7n2BfPk68lVi8n01yitT654rwx7hWOY4jZu7C1o5en35V0ljH5MfF8zMZRpYwub9pWHytUgc5p03qLc04I83i5H/5i9b5WKZwyNNaiP/8Abxf/AFFhXlW8md+1JqGmqrBfTajFT4LWySRl5J5ncz2Y7levJlpS+WOy1NNqa4NuNwNQXMmbI6QCPa3Ay4A9u7u7/csO1skpuLzfM6UaWq0uOlpb8IPNQ/wiNNaTMAvVrv1Fx88My0zMPx24Ik7sjxVJZf4TejtRVXmtso7zUVGM7G08Y9nfItQeXu06h1DrCrslHYZq+iZBGaV0VO+R7XbQ5zm7Tz5kjs69Fldi8mTNA+TA081vjZc/PWvkqzjiSPxjA7w0YAGDjtPeV9VdtkFGRO1dL0sxEuI3zs8Nc9x1yaxtbfK9CnFKiW1Dl6onX4o20fKzSg4Om9RA/wD9PF/9ReN8rVI+Tht03qIvxnHm8XZ/8xTUdEJqCilmb+XdTsMgx2Oxk/SSq2gtUbTHUFuDI/GejexfFTb2dX/5rzfM62lg6F+b0LePKrTuLANNai9M4H+zxc+Wf6T2J/KtTbS7+LWosA4/6PF/9RXcW2KKJp2/zMgyfYHYP0ZX2+1RAVMYb6rjj5+f1ldNLZ9mvN8zhdq/uLMfKrTguB0zqPLe3/Z4v/qK036/XLXtFJZqa0VVstNQ0eeVFeWcSaE9sTGMc7k8AtLnEYBOAcgjMW26IzZxniMa/Pt7PuX1TUbIm0uAMCN0RPUtOPuKqt6aHeooSe/HDiHQ3g2fcLHvbt4j2h0YcwNOAMj6eaqARJK2QjAniDse0cj9bV8sGxsJ+K50RPXvH3oTsjJ/oZc+0td/mfoXzHQimHoQP725hd83Z9CikztBaMuaQ4DqQc4+hVLmOniqdjXFuA9pxy3DoqXiNxkke5ZZpFXI5vH4wAex/wCVb/WbgAhAGxTtY70oZBtz8ZjuzwPL51HTh5YYi0tIJfET17wvtrG1DOCeTSC5jj+Z1afYtGT7ZJwm7JCTIw7CO9x7vEc16AeNFxHtEu7+aByWjae32r6e8cI8NwfUxtDXS7eZHeQoI2DziFzGnYCQXH844KqIy4KSH1j7lGpIfWPuRgmREUKeP9UrjTWnwx1D8oVH2jl2W/1SuNNafDHUPyhUfaOQHZipnes73qpVM71ne9VEZ4tR13/Tqn/tHfWVtxanurQ24VGBj0sr0M3v62u487OK+mllC8bmOHUK8xu4tCx2eZYD9CtCultO+ia0/mkt+leN/VVnhZWnivQ9D+nbTG0o8H6k00hjhbMCWmNzZMjtGCD9yyikc5wl3OLnB55ntOcFYts41I5h/OaWrILRNx4Gyf0kbH+IX5WhzQ13r3PcypfVS/H2LvQO2V8eTgOY4fUr8sRr2b4mg9mefu7fuXkM81PjgTPjHxQct8DyXt5utlTZXXvPKylxUjL0Vghvs7MCaFsg6xnafA/tV0pbjT1fJj8P72O5OC9Sm0pq1M4pplS9jZGlrmhzT2gq3y0Zp2nht3wd8faR7uquSLZTHnBjXt4TsxvBIHReqtuNCedRTtHEHrM+OP2q2NqWPwGh5efzQ0k+Cw0aTJl8uaHEHmCOiljpaub1YAwdZXY+gZKm/BzYxuqqvAHaGAMHicn6lYEkdvG24s5k5ieOZ9rV9SytgmnyzcTJgY/uhfcdVQ0hJp4nSPxjc0ZJH94/tVE+fjuDyA2R8hcWZzt5Y+5NSITGpld6oaweJUZ3ufve7JxjswvUUNQEREBFO7bGQO13IKgrWg0VQ3uMbh9CqpnbpcdzfrVPVDNNMB2lh+paRllqrrfT3CrrqSqbxYqFrHx5OD6rQPrV4MYMZj/NxtWGz0VXcLlpquDX8NtPG2qI5EPA/O9pWaogzWmv7tSaj8nOoZo2nayNjXcQY9NoaCfoWu9W2mWy3rSNZI+N0c88dTGWdzfOIP2rY+rNMst/k81HTwyvk85oRV+kANriHEgewYWsdc3Wqntek46uKOOS307GhzDncPyLwT4L0shxqo8X6HC21P7HWlecmjPWQ/8AgcvlKk5go39HA+LSPvTI6rzjsEUfFDnFsYMjxyIb3e89gXuZv6At/vPb+1AfNSXcItacPeQwHoScZ+nKqa3EVIImejuxG0DuH/2UUMEslRE9/DDIyXei/JzjA7vaV91zHvmhLQ1zWZONwBz2D6MoD5AAAA7AoJ+yo9sDvrUmZf6An3Pb+1fJimk4pMJaOC9vNwOScY7D7EBKvgHZV0zvjF0Z9xGfraFTPmjibHxamcSOaHFsbG4GfeD9a+DXUzHNe9lVIGEOy5zQBjvwCEBSX692ynuEbJa+njmiBZIx7sFucEKAX220hikmroGRzM3McXcnDuIUOo79I6uhgp4KvzYCVk7vM3kZ2kNOdvZnosddba2poaSkMdfNbG0ojdHAWQl0gcc5EmDjsXz1PHA9KysabivYff8Agu+paOz3eSUz3DzGsoC1nnG/ZjeMhue/kseoLHaqeuYLjf5bg1h3iMk8I4G7JJJBGCvultlVDW1/DpIZJKetpp3wwBsQxwyTjc7GefXmvIoZ6rWRfK2alMwezhgtOzMQG4Ed/wA+FyqopbvNH2WddVFLoprwidm7zRkY1LaaprmRVzHOMohwO0kuDcgdMntVzrNRWqle6lkrYIpYyMte/GOwrAXWeekjbDU0kscU09NEeFta04k9dwEjjudkAkYV+vV7rhVikpaSpNMyNkW7zWQgEZDjnHMdnNdaamkfLa5PRVUlS5+/h3GYCeCsjmfFKx9PUMEjHtPolrhjOfeCpeKXPjmLXHjxNd6LSRnv7PerZp9jYLLaoI42F8UJYZJ4XAgtwOw4PeVdXzSyxwtdIWne5r+H6Oe8e0dy7pypPNrpVNTSDKefhQObGAWBzSHuxyzyPevOFiDc6WJhZKXDB3AZHZ78kqN7BwpmOy8Mex4DyXcicHt+devaGNqABj0GvaB1aT+0KmT6bwDFPl8khaBIfRLMY6L7p3mN+xsTIxK04IO47gMjPzZTDTNEHc2SAxkdcj/JRBzhBHIfXiIJx1HaPrCA9ZI8CKd0j3HI3gnl0PJexQsjL49ozG7AOO7tH0FH7I+PG4gR9oPsd/yUmY1lPHUzh5OxrXRNGNx7s+KAwO83q533UElotUwggpztlmzjB7Cc+/krTcIqm0XSko75Ka61SSYEoeQRnGTkdMhVV1p6jS17r611BJNT1xEg2H1CTuIPzqkY2TWc9FTsoZIKaGXfLK88tuACPfhq+5JR/wAYODbnvPqKvsNp1vZha6idkMRqI6nIlc3d6rWnI7c5WSW6sil8o9ypmxsY+GEHPGcXOBA7GdmPb7lcdN0tXQ2iBlwDBXOkllmDCHDL5XP7f+8vu10NxhvBnqrpHVU79wZEKURuZ09ME5wOWML4U8Tu1gZKpIfWPuUakh9Y+5VkJkRFCnj/AFSuNNafDHUPyhUfaOXZb/VK401p8MdQ/KFR9o5AdmKmd6zveqlUzvWd71URni1jqCMtrQ89jm8vmWzlrnUcfOGTPVuF9GSVXbelb5Piy5TZFhVwtRIbK0/G3D5//sreqq2OxUvb8ZufpWf6ks7+R3v7Wn7e4zDXdyq7vT5+xcouW8dHFXPT7sRBp/ND2eDjj6FbG8pXjqAfu+5Vtkdtqpmf18/MWj7wV+Es9qP1mUqaU+8vdUziQlvUgeJwqVjtzGu6jKrJ/wCZkI7QCVRsAaC0djSW+BwvSyF/TUjxsrWpn0vC0HGRzHYei9RfafEV1Ld56bDZPy0ft9YftV+pquKrj3xOyO8d496xNfUUslNKJoXbZB4OHQr6LO3awqNqveZHWyz0r2zs9OnH84zHMDqF5VVIpqU1NNEx4cQSRyGOpUtFWMrqcSNGD6rmH809FTtjFBNw8ZpJzgA9jHHu9xX2zJ1KJ1TUzjLpyGnujGPpUQiZu3EbndXHJ+lfU0Boajg/9S/JiPTq35l6sG0EREAREQBePcGNLj2AZXqp6kkuazPokZI6oCJucZPaeZXkgzG/3FfS+ZHBsbyewAlbMFqtl5gvUTZ6RpNOJGtkkdy9IR4Ax71d1gtXDN5PqN1PC5tRDI+N8r3txsz3gezB8Vm8UjZomSNOWvAcPnURWYIyoqrnRawoatxfFDFJDCCOQYN/Jar8odRTV3k50Pc2RsjqJ6dsMu05JIjxz/3VveorKSqgqaKGRjqqGGpEzAObdz8tz8xXPmubbT0/kk0bX08W1z6oQyEEnLm7xlejm/8ANT/29Uzhban4HXxjbJQRB7Q4BrTgux071TTUgNJMQ1gc1u5pY4k8ueMquhGaOIDOdjezHT2r5MZkY5j+JtIweYH1Lz2dihD5JKd8O1rTtErGx8sgEZb9XivpsML2ghjXAjIJGVHSkulZtLdkLnFzi78zmBz7/wDJfEczXUuInBxJLGAHuzgfQgK22Qsa2WVrWtEjuWBjkOSp2NZUuknexruI4lpIz6I5D6s/Oq2VzKOlaza5wOGBreRJ/wCclULWxxtDWQVbQ0YAEjT9bkBJ5vF/RtHuC+YoS+WXhNIaI3MJzyLjjH3rzn0qmj/uH7147hmJkRNS1gcXufsBJJzyIAPXp3ICGtpKjjNLYXPaI2jLeoCpX0VS9rmmmlw4Y7Aq1zKY8zU1QwO3gkY/4VUSsbJNJ6TnbcDt7OQQH3XCSSzvDwxsxiBIecDdjsWHC6V8Fwt1PVRUZjrHuZmGRxc3DC7PP3Y+dZfFSwXCikpaqJs8bHkFrxkHvH1hYdRaaioK6tlktbZSajfT7I93DaAPDmudom2oPryeqhJ3ijq9UVEM1X5tSwTCKV8foNky4tJbgkDGRjC+LtrQUNRPC6iDnwGHB3gEh4Ge3mO1eHR1XVV3nr46WNxLy+I0Eu2RzjkuOH9uc+JV1vGnK+q/B00dNHK1kW2VkLjTvLweR3YccAdxK5Xaj7L+Tylr4Fup9XubZY6t1PSxiWtlh/KHIjawnDj3k8gM+1XjT2oZbxWxxuZG1opOIQx27Dt5GM+7BVvqdHXOXTtNSGnpZqh1TLJM0v2nhPJO3ftzu7PSwrnYrBV2+4z1cltZT8SPZtiq+KO3PYWtx4rdN6VJxtdBcqdOvHaX3sY3pHOQT7HAn6yEdybMPiSMkJ9/L7lJ5tO9k4LWM3bHNBd2uac88dnYF4IXv4wkfAwOZjDZM889/ILsecePYXySRt7ZYXN8Oz60jcJHwPHqytLc+wjI+nC+WVETZonyVdI0MPMCUEnkomOgayJouVJsjIIwQTgdn5yA+92KZru+LB8Dz+gL7c8RzTR4JLnbmtHaQRn68r5c+GRz+HVxuikd+Ywuxn2jkpZHEyuhjBBjaGPmONx5ZAHj2/8AIAjfJTUsDXVLmtlhYfT7Qwd2Vg9bqWW3mKen1HHUyPqIWyQebj1HPAcc92Bnn7Fm87qempzxnsjh7C55wPEqwQU9LPeaqtZU07zLBHEY2FucMLiXdv8AW+hVNIFhvmpqiC7XClpayNrGFuOKGDduaCD6Rye1LjqqptlptjYRTvrpBE+oLcFjGOeG5GD2nPL3K512naK4VYqK68zwtjkD4jmLawnkG82+3Ayvq7aUtUtPLTXB8b5GGMmrk4bX4DwWgnAAHdj2rUrAmJj101XVQ3eqjjq5PNnS7YOCxpG0MGck+0O+hZ1ZqptbDRzNeX7hzJ7c4OVjFZpiC6Tz1dNWUTqeqA4TnxbzCANuI3B4AGR07crKbWI4nQUzJxM+nAY87gXZ2nmcd57VHGEDEvakh9Y+5RqSH1j7lGCZERQp4/1SuNNafDHUPyhUfaOXZb/VK401p8MdQ/KFR9o5AdmKmd6zveqlUzvWd71URniwLUDGmiDiPSa4YKz1YZdIzJQztAycZCiruWtnV3nz5Sr1m0YcpaN2ysiPcchRIHbJIn/FeCvWzrZ6TI7WnunyxPOzZXcyuzffHngXw8pm+0FVFudsuePjtafA4/8AMoJO1h6OXrHGOtpnjvJB8M/cv5nQ8T95aqaGZW4bmkdRhW+PsI7+RPvIB+9XFW8Da9w68/pLf/KvQyF/VUu7r1PFypfRJ9IiL0TzwioLpXPoY4nsaDudg56KshlbNEyRvquGQutVlVTQrR6mZVadTp2oqqKrNDUCUZMbvRePZ1+ZZO5sdTCQcOjePELEVdLLXcN/msh9F382fuXWwtP9rO1FWwr5YPOoH0srvyrObX9ejvuKtUbnHcyQYlYcOHtV+njLg17P5xnNvt9ittyhD2NrYhzAxI3vI/yX1NHVMp0XjXBwBByCvVk0E7UWt7xcLtqzVNXYbXWCioaEHzibdt7CA4k+wnGPYo3ANkEEdowqeo/nI/cfuWtZ5bt5PLrROmuguNjq5Nm8v3Y6+4jOVsqp/nI/cfuVpeJHqI1RXiXgWmulHbHA93g0qtVNXSQxU+6oc1sBexshf2BpcAc/MVsyUO+K+UcNTVMimdLSwSSsI5FwyTyVdRy07uLFSvY6KF21u05wO0DwK8kbA+8uqKPhGhqKZvDdFja4tcQfrVi0/Z57Lc7k7czzKpmwxoPNrsbufvH1KFPuksUlNf7ncTI0xVodEGDtBDWlaC1beWT+Sq12XhOEtDcTIZO45mcMfSunWxYLsuJBfxAOh2hv1Bc5agsNOfJld7l6RqKa8iAjPo7TUA/evvyHCpf9qfc42ur7M6wkDTbPSkMbRGCXDu5K38Cm7TNBn+tDz+tVkhMlkJHMug/8qjNTCACZWAHq4L4WdSEu9Jro6qJrgMAtyOXTsKqqMSTTb5nU0oYOTmN9IH3qIVlOc4niOO30wqmgAZTvndy4h35/q930c/nQHzXAvmiLZIgIwcte7HM9/wDz1UWJ8bt1Ns7M7z+xfMTRKHSvaC6Q7uY7u5fDGtdHTtwMOdJL2dozgfQQgPvdJ/SU3++f2L3Mp/Ppv98/sX3w2fFb4L5e2NjS5zW4Az2IDx8U8sbg19NjvO88h4L7Mb+LLJG6J8ZwSd/ZyA+5QiIACItAL/ykox4NXjmMY2cNaA0mIEAf1/8ANAVlA1zXzF23DsEBrs8+/wC5RTirbNIQ+XhZ5cPacfNjKQBsVbHtAAe0twB39v3LyriBrgTkbmciDjmDz+sICMSZOHVEod0cdpUhj3NAL5MDpI4fevCJcY4ge34so3L4w1v/AFckXtjO4eBQHnN8FMC+QjiyN9c5IBdjJzz7FSTNY2Z7Xlzm8sBxLgPFV2wF0TYHslEUeNhOHDqce3kvXSBnORjmH+s37+xRoqLUYKZ5B83Y4g8vyY7VMIdvZTke5quLJGSDLHtcP6pyvpSBJbd2DgtcD0wvDIBkYO7oQrmiQWSnbARSmMesWnxKkDw6UPHqzxtkHvHI/RtUip/VjHWGXxa/7sn/AIVoyWzVMNZU23gUdIyofI4Bwcxj9o6gOIGVg+m7VK6CrioKB8dFVRTwOlqRG57JMFpIc052lw7MFbTVsNPBbi1kLGxQvJcGjs3E5PjlW9CESzXFLaHz241FNHRMDmxU4jpjhznRy4c8ggc+RVXfbfcbldru/wA5YZc05fSxnHoNeNpzj0uW4kd2AsvZp+ic7iR2ppOd2QzbzznPPCkqbBDW1DZprXvlJaC92OwHv59Fb5Lpr2ktskT7Y6ppAWsoi38pRecgO4zzgDPI471lukRKLzduM8PeZmHIj2YHC5DGThZRNb6ei83ihgcylYxw2xgkAkg/tXxTw0sdTvp2xtklfmQt7SQ0gZVvSILkpIfWPuUakh9Y+5ZYJkRFCnj/AFSuNNafDHUPyhUfaOXZb/VK401p8MdQ/KFR9o5AdmKmd6zveqlUzvWd71URnixSZhkikYO1zSBlZWsXK+fKXCTOdeJghG0kdF8SeofZzVTWRiKrmY31WvIGfeoCMgjqv1LStbON69TwKanZ1p7mXrdvp2vz3B33pUHbwn5xte0/NlR0J4tDGD8XaV9T/lKRx7CW5X8nau1Q9h/SPzLxMviO6Nh6gKkkAbMepLh4bT/5ipqCQS0cLx3tXxOMSE+1vgQ7P1Bfbkbi1jenzPHyhTZs+ERF6h5Za71QyVjIjCzc9p588clV2+GWnpI4pSC5vLl0VSi71ZRXVZKxepHNWSVbr2sLwjPYSD2gjtBXqLgdDJLXXeewYfymj5PHX2j2FTuaI5CcZjk5OHQrF4J30k7Z4+bm8iPjDvCyuKSOqha9h3RvGQvQsrS+u87UuUWKaA0NSYefCf6UZ+sfMvVcqyBlRA6GR4bIwb2PJ7Md/wC1YzW18IoSZKt9Od+wyQs3kEdOR5clt4HRF0Wu6+Ot0Rqe63yntctxobnGQ8RSmN0T3ODnZIBPaCR7+3krhTXmea6mnt1zlrWyUswYZosME4xsBIaPblW0agr/ADaqY2umDwyIATlgk428bg3H5uFhmoKWoqK3ylyWun/A09BbqOYySzzVDpdwOMtBc0c+Xt7Vsep/nI/cfuWI3zUkkN1rY33GRrG8qcUz4yAcc9+7nnKmuF/rhb7S+mDKiqaGT1skODFsHrtDjyySMAKpwyNYGSq0amt093tE1vpdvnFS5rWbjgciHH6GlQ6erauuqLlJUuDoi6Mw7M7QC3uz9KvrctnglB9KJxcB1y0t+9b1ozqMa026a3Q2201MbeJHSulbIx+QWufnGMKqv16hsdBUuLouM6eGUNd6zm5w7HzA+Kro7bGyalnJJnp4uCHdW+5WTUVlkrbvb68RRzQUY3zRP/PZuGQPFNgLpe7w+z0EVZDRvrg6aKPgxEbnB7g3LckDPPvOFqLWdmnborUUEdpuNDTOqY6rjVs0fpSuqGZZw2Pd2DvxlbQ1aZqOwVgoXAOglYYHEkNLQ8Fo9w5D5lri6m6+Z3s3k1jqIUb3S+cAjErXA+hyxjibcZ7l6WR04Krv9IOFo8YOgqeR0VihkacPbTtIPt2qN1VMyCMh7SSX+k5oOQPdhKY7tOQnrSt/8KgkOaOLu9GQ/SV5zOxWcWqfUwxsZGGFjXyOLPacgc+3kpbk4tpw1o/nHtYcdwJVTF/Ns9wVNcHR8ONjw4lzvR2nByOaAglfwYXvA5MaTj3BeCPhytizkQxMYD7e/wCoKM7HDaXVAZ3hzAcjplSsJe+WQgjiOyAe4YAH1Z+dAfa+JBvfFGex7wD7hz+7C+18ZxU0398/+FyAjc/bNKRLTEucTzkII9nYvrhOkp53cSDJewjD8tGCDgnCQD8kD1JPiV8OAMFSMDnPH/5EBI2OZ8sLs04DHAktkJOO/uVRXROPDlaWAxkg73YGD7cdcKkmp43RvxEzdjI5DtVTUnjWtz+0hgk95bz+sICma9r3NBrGYcQMQsyR85yPoQtYXva2GSUsdtJllw0/MP2L2pyad5b24yF7OInVEgbAXHk4u4hA5+xAGNkYCIxDAD28KPn4nl9C+yJRjFRKPA58QsV1Ndq2ifBbbeMVlVl/I54bOzkT1IKsHB1HazJVMqpJOEcua+Tc1wxk4BPYvEyvPdNhauyps6q1T+ZrZhP3w1n32OQO0oVbqSnVJsZ7ZJDmRlNKB3OjwfHJ+pfOxo5mlc0dIZj9XIKG0XFt2ttPWNGBK3JHQ9hHiq1exZWlNrQrShympX3PirpdFTpq1ogy0etJVs9hY131ArziAf8AvTGj/wCLEW5+fIVQvgtM8rYRnBG55HcP81syMVA5mKN7SMtMcmSfmIH1qN7QXZfDOHcsgNcQcdmccigDZpXmVuJByDTyLQOzC+xEWjDZJGjvAdnPigPHl4xxHCFp7B6zz7gP8140Fp3RRiM98svpPP7P+eS+JJI6fIY3Mh9uT85VO/fL/OOyPijsUbKkSVEzAx3+0SPk/vnHgOS+WGJwHDrJo3H82clvPoCMD618gADAAAXpGRg9iklgrPOqun/nQ1zR+c4cj/3h2fOAvp1Rx5IMw7XBx9IEEH0T39v0KgjL4DmF5Z/V7W+Cmhka+oi/JbH5Odp9E8j3dxWkzLRcVJD6x9yjUkPrH3KshMiIoU8f6pXGmtPhjqH5QqPtHLst/qlcaa0+GOoflCo+0cgOzFTO9Z3vVSqZ3rO96qIzxYuVlC1vq+sudpq4Y4pH01I8OPnEVOJXF2PRYQezPNYrsHbRTSzlaOFJR3qPh3CTDcB2COXby7fFW9U7rrfamW6NmpT5zFHSPLdodwNwHE2s7ztLjj2KOnuVxlq6OOGUy0zql7S98DQ6SIR5y9gLQMHPPl2di9yxtbllTS8YR5deTOqttPWX21OzA9vxXn9qqwN0bm+8LDZ7/d6al1A+OV0kQhlko5BBwsMazBkaS7uI7O05ys3kP5eTPacH6AvwudsgeT1u1nCpuPU/X5BlWloVnGKS5F2sMm+2xf1cjwVXUj1v7mfBzT9WVbdOv/I1EXxJDj3K7StLy1o/ODm+LSF8WS1Ra0vvMZRT+ZFKi8adzQR2EZXq9g8UIiIAi9a1zzhoJPQKsitsjubyGDxK1TRVVqRUm9RRK5Was4E3m7z+TkOW+x3T51UR0cMXYwOPV3NSyRiVhae/v6L6rOxqpcydKaWsT7vMFPLROfUOYxsRDw54yO3sI7wezHfleuZSMp/OjS43AOLdgDh7x1VHbIHVjxJVSGTzN2yNhyQHAeuSe13P5ld5XNjaS/8AmzydnsC+g6GM2mKjs8X4PftkqZ5ZJmMkcxrjudnkNxOB1AWHQ6QoJXecvvUJgklkeHSBoyWvOW+ljsPL5lcbxZKs3mpbHQSSvnqqaenrg3IgjYWlw3d3IOGO/csffb56yy0mKV9cPP6p7WsgB4LeJKCCe/JIPzLDfcaSMzqqe03espZx5g11ud5zNhwIcC0gbnYwBzz3qC8Ms9+iooqO4UUVTBOXMhgmjla/0T2s7+3xWKUVLM2nvVLKW0DTbYqd8cseAXlpAJI7Bn2KexyilrqYyVUDmtIbymB7sdnCH1pIgyW02ltqbUYlMj53h7sNDGg4xyaOxXFCMHCLpqMBRvjLy7nhronxkY5+ljmPdhSIgLXqKsqaK1OkpJYoZnTRR8WWMPaxrngOODy7CVrO93hl4GpbNJqG310clNM6mpjNFLLPKGnG17X+j6QyGEEgDuW1LsM22c+cUlOxnpyTVcPFYxg5khuQCezGVhkt2gko6mGHUEtVBLmB8PmscTpWujc7Mbo9rm+qRu54K+7JWlTMdeRxtNZtWhB/i3TA9vmrc/7oXzHxGebxbYpGudyDmnIHae9fVtcJNOUjhnDqVp5/3QpLe3iy8T82NoaPee1fE9Z2LkrdUzPlmfE1sTo48A72k5PTt93iq+SRsUb5HHDWgkn2K2QNcIw54w9/pO957vu+ZQHw5sbGkvp3saOZ4MhwPbjkkvBhkEbWVcnohwc2QEEezLuaqFA9oYwMcC6nHZj1oz1Hs9iA8BGAS+qjHtY131ZUkGySeMecAkZLWvjLHE4I7+3lnuRji1wjeQSRlrh2PHVfHpNkc/hcRzTlhL8NAx0HegPGyCFojNRSZZyOZcHwX0GsdSyv85gw6VrtzTuaCMcvoX01wMpLRhsoEg9/evg/zc56Ts+pqA84g/8AztN/uH9qrqAM81axsrZQ3LSWjkoV7RHbU1DPjBr/ALvuQFKxkbWtp/On4b+T3NiPaOXaQQpRjjyBuS1gDCT3kBQ1r203n5e7Y1v5XPxQR2+IKwe4aj8zkjdS6kdVSGePfEYGDLC4bicNz2ZXC3yimx/N7e7OlnZO01F41Va6g19vvFLC6d9IQ2WFva5gOeXifFUFfqdl+pKi3W6zzCrlG3PL0Off0VNfNU1cF5raamrJYmQua1oIhDSTG13LdzPb9auFdqiptunrHUVNQx1VUysbM4AOGzJ3HlywOXNeBb5PRaWtrVZWjopq/NgnP+3BvU2ejZ2jpooVdMtasWu/FbTIbJb/AMEWmnpXOBMTcuPdk8z9JVaDNsbIYSY3DI2nLgPaFg92vdY2XUDKerkL2TRMpY44xJvLoQ7Zz5AEnJPRRWa+XWsu1sbJPLNKYZRVtcBHw3gt5DAwcZOOuV6tjlNlYqmwoThYL7OD5K7KuubSp68fczw1DNhc07ueMDtz0968dEdnBa+OSckOmYHYceXYPZ2Kpe5kEcMkzBJUD0Q4NG7PfhUbYQWbWfl2gk4PKRp6+/2r0T5T0PLvQc0yBvax/KRvuP8Az71j+urrV2LRt1utuqcT0zWiPiRglri9oOc9vIq/8QO2MkzKNwaM8pGE+1fc1HHVOjpSI56Nu5s0T8O3HtyQe3ngrNSbTSOOU2ddrY12dnVdbTSe5ta/saf095RLpU6hoqC7xNo2xUT56zLBiTADmyNPQt7lbrT5Sb7d66qp2T00ZuMMr7a3hDMUjHEtY4kc9zR2+0Lbht1LISZqSB72AxZdG0nYDgN7OzGOS+/MaXEA81gxTkGL8mPyZHZt5cvmXz6Ov+48P8JzhCpeVOE29stOFGvYk/u5NVU3lLut8L3W9ohZRW2SWqzED/tAY4459gG3Ktls8o+on2Gvrp6pr5WUpljzFFtDtwGcA7vEBZTU1EWkdX1tRcWEWitift4cILdzhzBA7fnVvu1ytWoIqO26Zp2tqpZsTPbSBgMftznks3K/7hVmfL6rreUuYc60pepxOz1lly8nurbvf7tW0tY4y0kFOyQyTQNgkEjscg0E5b281smD/pEXvP1FUkFHDTuZI2KITtjERlDAHFoHZnpyVdSxl0od3M7/AGr6LKl0qG5PWyDJrXJrDR21pfcvF/zu8WV6kh9Y+5RqSH1j7l1Z9ZMiIoU8f6pXGmtPhjqH5QqPtHLst/qlcaa0+GOoflCo+0cgOzFTO9Z3vVSqZ3rO96qIzxfEkTJQA9ocBzGe5faKkKB1lt73yPdSxl8oAe4jm/AwM9eXJQwaas9K/fBbaaJ+Mbo2Bpx7wrqit57yXVuLe6x258ckbqSIxygh7SOTge3I78r7NooScmmZnqq1FmpKr82JpNrUUcdro4S4xwNYXdpaSMqTzGDl6B5c/WP7VUIs6OjcW83tKUW6lAwIuXvKfg6m/ov+IqqRahGYRS/g6m/ov+Ip+Dqb+i/4iqpEhCEQNo4Weq1zfc8j719ebR/1/wDfd+1SoqCLzaP+v/vu/anm0f8AX/33ftUqICJtOxgw0yNGc4Ejh969dA1wLXOkIPIgyO5/SpEQETKdkbQ1hka0dgEjgB9KijttLC3bHEGNyThpIGScn6ST86qkQFKLdSiR0giAe8AOdk5IHZkr68yg27dh29uNx/aqhEBTeYU/9H/xH9qeYU/9H/xH9qqUQFN5hT/0f/Ef2p5hT/0f/Ef2qpRAUc1qo6iN0UsDZI3cnMfkg+8FUbtKWR8zZ3WqkMzc7ZDENwz24PznxV4RVVNamIRAykiiibEwObE0bQxryGgdMZ7F9Mpo4xhm9o6B7h96lRSQROgY9pa4yOaeRBkcQfpTzaP+v/vu/apUQEXm0f8AX/33ftXnm0f9f/fd+1TIgIPNIsAYdhvMDeeX0r682j/r/wC+79qlRAQ+axDHJ/Ls9N3L6U81iwRh+CQT6buZHf2+xTIgIvNo/wCv/vu/avBTRtduG8OxjIe7P1qZEBCaWNxLjvJIwSXu5jxUQttK2Z04hAmc0NMgJ3EDsGeiq0UeJSimtFFUvY+anZI9jtzXP5lpxjIJ7Dg4XzJZbfK/fJSRvfsMe5wydp7R7vYq9FLqewSy3x2S3Qxyxx0kTI5QA9rRgPAGBkd+ByX1TWago5DJTUscMhGC6MbTjpkKuRLtO4SyI00biCd5I7Mvdy+lfJpIi4OIduHYd5/ap0WiEJpoy4PO8uHYd7sj6V4KOFrtwDg4d4ec/Wp0QEBo4SSS12TzJ3nn9KeZQ/Fd/vn9qnRAUc9qo6pmyeBsrPiyZcPAqOlsVtos+a0cMGe3hDb9SuCICAUcIOdp/wB4/tUzWhoAaAAO4L1EAUkPrH3KNSQ+sfcoykyIihTx/qlcaa0+GOoflCo+0cuy3+qVxprT4Y6h+UKj7RyA7MVM71ne9VKp3jDj71URnyiIqQIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAKSH1j7lGpIR6RPsUZSZERQp4/1SuNNafDHUPyhUfaOXZb/VK401p8MdQ/KFR9o5AdmL4ezd719ogIOG7onDd0U6KySCDhu6Jw3dFOiSIIOG7onDd0U6JIgg4buicN3RTokiCDhu6Jw3dFOiSIIOG7onDd0U6JIgg4buicN3RTokiCDhu6Jw3dFOiSIIOG7onDd0U6JIgg4buicN3RTokiCDhu6Jw3dFOiSIIOG7onDd0U6JIgg4buicN3RTokiCDhu6Jw3dFOiSIIOG7onDd0U6JIgg4buicN3RTokiCDhu6Jw3dFOiSIIOG7onDd0U6JIgg4buicN3RTokiCDhu6Jw3dFOiSIIOG7onDd0U6JIgg4buicN3RTokiCDhu6Jw3dFOiSIIOG7onDd0U6JIgg4buicN3RTokiCDhu6Jw3dFOiSIIOG7onDd0U6JIgg4buicN3RTokiCAROPsUzWhowF6ihQiIgPH+qVxprT4Y6h+UKj7Ry7Lf6pXGmtPhjqH5QqPtHIDsxfLnhvvX0qd5y4oD64x6BOM7oFGipCTjO6BOM7oFGiAk4zugTjO6BRogJOM7oE4zugUaICTjO6BOM7oFGiAk4zugTjO6BRogJOM7oE4zugUaICTjO6BOM7oFGiAk4zugTjO6BRogJOM7oE4zugUaICTjO6BOM7oFGiAk4zugTjO6BRogJOM7oE4zugUaICTjO6BOM7oFGiAk4zugTjO6BRogJOM7oE4zugUaICTjO6BOM7oFGiAk4zugTjO6BRogJOM7oE4zugUaICTjO6BOM7oFGiAk4zugTjO6BRogJOM7oE4zugUaICTjO6BOM7oFGiAk4zugTjO6BRogJOM7oE4zugUaICTjO6BOM7oFGiAk4zugTjO6BRogJOM7oE4zugUaICQTHvAUoIcMhUykhPMhATIiKFPH+qVxprT4Y6h+UKj7Ry7Lf6pXGmtPhjqH5QqPtHIDsxUzvWd71Uqmd6zveqiM8Wnta+VnUOntQTW602KjucLJREXB88b2O2g7TujDXvPMhkZe4hpOFuFaD115MK+56rr6iCx1j7c4tnpZbY6nkc+Ygl5l84mbt9InHDHZ35Vp/NjqJV+XAvGoPK5qCgsllntdloLjcLjxHbITVuIbG/EmIuAJMBuMudtAcex2MG+aF8plTqzUM9vqaGGGF9BTVtO6j4tQzEnEJ3zbGtHJrcBzWnJcOeFi990BfBobS9PSabts14iq+NcadrI6hjGujc1+3jysPPEYIbIOztI7ch8mFn1JRX+/Vl8tUFvpZaWkp6RsVPBANsfEBaGRTS4ADm4y7HPkBzXWKYfc+vbyZzmrB+Hry1lq1H5a7haZrtDDYYGR0zq2OnqpasvD3023dujDQQDvH5yjvvljvVt8n9w1LRWy2VYguFTRtldUbI2tZLsjIbkmUuHc0jr2KfV3kZluEN6q6G6VNTU1HnclNQStjbEySoLeJ6WNxHoDtPcooNCXq2aRmNvstGL9R3mqqKLeyF0kNNJOSXQb/AMm15Zjbu5Dvwsq7C3xj4yvY05n7vyhwXy1eUC+Xe40tI222umkq7LLc4g+sMsZcJQ1mZmDGwtOThpIWHUvl5vFTZrjX/g21uNPVx0bPNHz1ABMrG8R/oNDInMc4seXDcQBjuV50RoCpZfa83SxPbYZLe6mEV1homyOe+Xe9ojpPyZYcZJcN2cY5ZWMS6R1zV2uIVFmuLrzmKB5MFmFH5uycODMgcXYGtBA7nALdKpvKesX7YcTDvXX1sXvPobH1f5T4rLcoaCzUr7zU0kzX3aOkjdKaGl/Pe7b+f2YZzJGeXJZjZr9bNQ2yK5Wuugq6CUEtmieC3l2g9CO8HmFgmrtAV89wtVRaLvqMNqLm11wbFeJWMjpi15ftbvAADtoAbz6LK49IUMOnq6zTPrLlSVYdxG3GpfUOdkeruec45dmVycKidvXDd9zrjfjZ1x9oMHk8tVNU6YlvdE2lEdNfI7ZOHPdKBA6UN4w24PpMO4do96ax8qt9tl0FHpnTbLzHE6OSaQGpDhEWhzuXA2g4PLD3u/qHsVlobBerhQWNt68m88F5opKYy3C3VFBEC2GUPYzAkzwxtHogdg5KyeVHyRXbU2tq+sobFttruDNHNRUNBI6aXGXl5lmif245HcHc8raVKaT3+y9cfQxLcvu936Yepu/R2o3apskdwfBHTyl7mPijMpDSD/8AFjjd2Y7Wj51f1hPkp0w/Sej6aiqKJtLXGSR8580p6d0hLjtLmwPezO3aM7icAZx2LNlmtJVQi0NtYhERZNBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAUkPrH3KNSQ+sfcoykyIihTx/qlcaa0+GOoflCo+0cuy3+qVxprT4Y6h+UKj7RyA7MVM71ne9VKpnes73qojPERah1/8AwgLNou6S2qkopLrXwHbOGSiKON3xd2DkjvAHLrlbpoqrcUo+bKsrsclo0lvVdRt5FzlN/CpME0kTtJAujcWnFy5ZB/7JfH4136I/rL90unw9puPN+Yc3dpwq5HSCLm/8a79Ef1l+6T8a79Ef1l+6T4e03D5izd2vCrkdIIub/wAa79Ef1l+6T8a79Ef1l+6T4e03D5izd2vCrkdIIub/AMa79Ef1l+6T8a79Ef1l+6T4e03D5izd2vCrkdIIub/xrv0R/WX7pPxrv0R/WX7pPh7TcPmLN3a8KuR0gi5v/Gu/RH9Zfuk/Gu/RH9Zfuk+HtNw+Ys3drwq5HSCLm/8AGu/RH9Zfuk/Gu/RH9Zfuk+HtNw+Ys3drwq5HSCLm/wDGu/RH9Zfuk/Gu/RH9Zfuk+HtNw+Ys3drwq5HSCLm/8a79Ef1l+6T8a79Ef1l+6T4e03D5izd2vCrkdIIub/xrv0R/WX7pPxrv0R/WX7pPh7TcPmLN3a8KuR0gi5v/ABrv0R/WX7pPxrv0R/WX7pPh7TcPmLN3a8KuR0gi5v8Axrv0R/WX7pPxrv0R/WX7pPh7TcPmLN3a8KuR0gi5v/Gu/RH9Zfuk/Gu/RH9Zfuk+HtNw+Ys3drwq5HSCLm/8a79Ef1l+6T8a79Ef1l+6T4e03D5izd2vCrkdIIub/wAa79Ef1l+6T8a79Ef1l+6T4e03D5izd2vCrkdIIub/AMa79Ef1l+6T8a79Ef1l+6T4e03D5izd2vCrkdIIub/xrv0R/WX7pPxrv0R/WX7pPh7TcPmLN3a8KuR0gi5v/Gu/RH9Zfuk/Gu/RH9Zfuk+HtNw+Ys3drwq5HSCLm/8AGu/RH9Zfull1l8u/4YtkFb+AOFxd3oeebsYcR28MdF5+csssc2WStssqu0txMN4w3sncz0s15VZZ2tnYZE79SUxisJS2wtbRuJFqz+WT+xf8V/oT+WT+xf8AFf6F4vzdmftv01ftPf8AwLL+z408zaaLVn8sn9i/4r/Qn8sn9i/4r/Qnzdmftv01ftH4Fl/Z8aeZtNFqz+WT+xf8V/oT+WT+xf8AFf6E+bsz9t+mr9o/Asv7PjTzNpotWfyyf2L/AIr/AEJ/LJ/Yv+K/0J83Zn7b9NX7R+BZf2fGnmbTRas/lk/sX/Ff6E/lk/sX/Ff6E+bsz9t+mr9o/Asv7PjTzNpotWfyyf2L/iv9Cfyyf2L/AIr/AEJ83Zn7b9NX7R+BZf2fGnmbTRas/lk/sX/Ff6E/lk/sX/Ff6E+bsz9t+mr9o/Asv7PjTzNpotWfyyf2L/iv9CzHS+saHVMbxAHRVMQy+F/aB1B7wvryP+oM3ZbaaHJ7WanshrylI4ZRmvK8no0lrRC8U/RmRKSH1j7lGpIfWPuXsM+AmREUKeP9UrjTWnwx1D8oVH2jl2W/1SuNNafDHUPyhUfaOQHZipnes73qpVM71ne9VEZ4vzouHnHn9V55u874ruNu7d+Tuz7c5X6LrUOv/wCD/ZtaXSW60lbJaq+c7pyyISxyO+NtyME95B59Mr6cntaaG720/Of1Dmy2y6zodhi6Zw1TMcjka4//AIhV/wDav+sqmXSc38FYzzSSu1aA6RxccW3lkn/tV8fio/pd+rf3q+v4iz3n5Or+ns4tt6LjTzOb0XSH4qP6Xfq396n4qP6Xfq396nxFnvJ8u5x7LjTzOb0XSH4qP6Xfq396n4qP6Xfq396nxFnvHy7nHsuNPM5vRdIfio/pd+rf3qfio/pd+rf3qfEWe8fLucey408zm9F0h+Kj+l36t/ep+Kj+l36t/ep8RZ7x8u5x7LjTzOb0XSH4qP6Xfq396n4qP6Xfq396nxFnvHy7nHsuNPM5vRdIfio/pd+rf3qfio/pd+rf3qfEWe8fLucey408zm9F0h+Kj+l36t/ep+Kj+l36t/ep8RZ7x8u5x7LjTzOb0XSH4qP6Xfq396n4qP6Xfq396nxFnvHy7nHsuNPM5vRdIfio/pd+rf3qfio/pd+rf3qfEWe8fLucey408zm9F0h+Kj+l36t/ep+Kj+l36t/ep8RZ7x8u5x7LjTzOb0XSH4qP6Xfq396n4qP6Xfq396nxFnvHy7nHsuNPM5vRdIfio/pd+rf3qfio/pd+rf3qfEWe8fLucey408zm9F0h+Kj+l36t/ep+Kj+l36t/ep8RZ7x8u5x7LjTzOb0XSH4qP6Xfq396n4qP6Xfq396nxFnvHy7nHsuNPM5vRdIfio/pd+rf3qfio/pd+rf3qfEWe8fLucey408zm9F0h+Kj+l36t/ep+Kj+l36t/ep8RZ7x8u5x7LjTzOb0XSH4qP6Xfq396n4qP6Xfq396nxFnvHy7nHsuNPM5vW3NF/Bmh/7/AP43LNPxUf0u/Vv71ZdZfIR+B7ZBRfh/i8Ld6fme3OXE9nEPVfkP61yO2znkFFjkdN6pVpxKWF2pbY3o/b/0FktrmnONdvlquUuhqcHjepeyXqTNdotp/wAjf9tf4X/Wn8jf9tf4X/Wv5h8o547H9VP7j+ufjuQdpwq5GrEW0/5G/wC2v8L/AK0/kb/tr/C/60+Uc8dj+qn9w/Hcg7ThVyNWItp/yN/21/hf9afyN/21/hf9afKOeOx/VT+4fjuQdpwq5GrEW0/5G/7a/wAL/rT+Rv8Atr/C/wCtPlHPHY/qp/cPx3IO04VcjViLaf8AI3/bX+F/1p/I3/bX+F/1p8o547H9VP7h+O5B2nCrkasRbT/kb/tr/C/60/kb/tr/AAv+tPlHPHY/qp/cPx3IO04VcjViLaf8jf8AbX+F/wBafyN/21/hf9afKOeOx/VT+4fjuQdpwq5GrFl3kz4v8cKPh52bJOJj4uw9vz4WS/yN/wBtf4X/AFrMdL6OodLRvMBdLUyjD5n9pHQDuC9XMv8ASmcbPLbO2yim5TQ09acxjGDes+HOGe8kryeuzsqrzqUant8UZEpIfWPuUakh9Y+5f1Vn4omREUKeP9UrjTWnwx1D8oVH2jl2W/1SuNNafDHUPyhUfaOQHZipnes73qpVM71ne9VEZ4iLgaXyj6wmlfI7VV7DnkuIbXStGT0AdgD2BdrKxdpMM8jOud6M3XL9Ld6dXdHM75RYF5GLrW3ryZ2KtuNVLVVcjZmvmldue8Nme0ZPaThoGe3ktL3W5zwanrYZI75UC4VtRURyGnrYeFG55MbDG+2vd2HaNrnDl3DCxd+p0no2FsraxptlqqSfmpOpUWldfGS6R2KsZ5/SUPmvEg3isY41kbhw2GESwOYO073tGeWSzlnIvJxDcv4w6nqJJ4fwbI6lAjO+eQzClhDiKgvIe0EOb6pORnd3IqcG93XX23nS9iuuv87jZCLQN6usVi8oNZWTakitDXObFFQXfUsjjTybi7zt9I6sYx0DwS0RgggNa4N5louVn1PXWvyZaZoqeV2K2ie6e7Udvq61sDdx9ONsEUjHEjJG+Ru3kfTHJSPpVSLP1Xeuut07sRa68mOppa+nhtLZKy4W2mpGOortU26sppKqIYDeJxYGxl20t9Jrzv5kMaFZdX3S96XvesKqG+VUzvwEKmmjmDeBROMz2BzGAAHaMEl2ScczjAFdLVV3fPBN+xFV9N7dHFpe5t9FpjUdXc9D1VwsltvdxlhraSjcKitqDUzUcktYynkma5+cZbJuDT6IMfIAEhZdpMTWbWF/02yurKy3U1HR1sJral9RLE+V07Hs4jyXFv5FrgHE43HHLAC7KvLVy1idfd7xHqZyi1pTeUS4za9Om3MoBbxVyRi7BknClIZu8zb+b5y3td6RGGnA3bmsn8qlsimoYZaesu8V+rnNt1uio7rU0sZmeSQ9zIpGhwY3dI44ztYfYsw8HvLta3GxEWlb6+50du1xchf7o2p0TTRx0DfOXNZM6OkjndJOwcpjK55Yd4ONvo4OSqq/3DUOltQXCvluTWwXGCd0IfUSTMgiaInccwbdrBC0Sg7ecjpIwebuWruMJ9bPPhtJLhNrrb5cTcCKz6XrqKvslK+31NVUU8QMBkrGyNnLmEtdxBIA8OyDncArwstQ4KnKkIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIqS50P4TttZRec1FL5zC+Hzilfsli3Ajcx3c4ZyD3FRlRVouSKy46i0PS+VXUundXamu9m05RttNLJea7ztr690jONMwEBuIQQ3mDz3do5LZOhJLvpHytDR79S3bUFpr9PNu7pLpOJpIJxMIyWvwCGPBzt5gEcl0VExD18m/RT9/Ew6o1rqUvVx5m7lJD6x9yjUkPrH3LmzZMiIoU8f6pXGmtPhjqH5QqPtHLst/qlcaa0+GOoflCo+0cgOzFTO9Z3vVSqZ3rO96qIzxa6l8hHk+mlfI7TzQ55LiG1U7Rk9AH4A9gWxVbH6ks0b3Mfd6Br2nBaalgIPTtWa8opsfzV3Z74OdeR2eVYV2arjepgkslkoNOWumtdrpm01BTAtiiaSQ0EknmSSSSScnqsUr/JdR1sl0bHeLlS0F1kfLV0McNI+OZz/Xy6SB0nP+/wAvzcADGbQTxVMTJoZGSxPGWvY4Oa4dQQrVFq/T1RU+bQ322S1OSBDHVRueSO0BoOc8jyWlUn9UzJtUKhKhKEuHIq/wPQOpKalkpYpoaZgZEJ28QtAAHa7J7AOaprNpqgsNbdqqiYYzc5mTyxgNDGubG2MbQAMDDATnPPKt9R5RtIUNNQ1Fbqe0UMdfTsq6fz2sZTulhcMteGvIOD7leqK70Vxqa2npZ2yy0T2smDQcNLmNe3n2HLXNPLPatO9jJFChIx+t0RNLe7hdbfqe9WqW4cMzxUjKR8bixu0H8tA8g4HccK6WTTVNYtNU9hhmqJaWGEwCWYt4jgc5JIAGeZ7AB7FTxa+0nPd/wRFqeySXbiGHzJlfEZ94JBbw927OQeWM8lc7feKG62yO50lQ2ShkYXtlwWgtGcnBAPcVl/lx1Gl+bDWLLaobFZrdaqZ0j6egp46aN0hBcWsaGgkgAZwO4BQ1unbXcamqqKujjnkq6Q0MwkJLZICSSwtzjGXHuzzVdR1cFwpIKullZNTVEbZYpWHLXscMgg9CDlfIr6R1ZJRCqgNZHGJXwCQcRrCSA4t7QCQRn2Faql1N1a+pM0wklTq6gx2h8nGnKChuVEKKeohuUTYKl1dWT1Uj4mg7WcSV7nNa3cSACACSRzJKuGntKWvS7KkW+OoMtU4Pnnq6qWqmlIGG7pJXOeQByAzgc8dqqbZqC03qlmqrbdKGtpYHOZJNTVDJGMcO0FzSQCO8FfdpvVtv9G2ttNwpLhRucWiekmbLGSO0bmkjISXiIRW7R0HblUVRZ6KqulHc5od9bRMkjgeXuxGH43ENzjJDQM4yBkA4JyqL1baRlY+ouFLDHQ485fJM1rYMgEbyT6OQQefUKmrtV2G2VLqauvdspaloBMU9VGx4B5jkTlQpQXnQGn7/AHQXKvpJn1BEYlayqmjiqAw7mCWJrgyUNJ5b2uVwrtM2q5SXF9XSCZ9xphR1Bc93pQ+l6A5+iPSJ9HHPn2gKmq9c6at8NRNXX230UFPU+ZvlrJ2wM42wP2BzyA47HA8s8vcVUW/VVlu81HFbrnS1prYZainfSyCWOWON7WPcHty04c9o7e0+wq4xHWqfTgSVPWz+SezWWjsFCKOhZI2He+QmWZ8r3ve4uc5z3kucSSeZKuCs901XY7JVx0lyu9FSVMjQ8RzzNY7aTtDjk8gXcgTyJ5dqrKW60lbVV9LBMHz0D2x1DNpHDc5geBkjn6LgeXVR72XVgViK1DUtqMFpnFYww3f/AKG8A4m/Jul5HHL0Gudzx2K4UtVBXU0NTTTRzU07BJHLG4Oa9pGQ4EciCO9GmgSoit1dqC02yupKGuulDS1tadtNTz1DGSTnoxpOXfMgLiit9RfbVSXOmtdRcqKG51TS6Ckkna2aUDtLWE5cBg9g7l9wXe3VVdV0MFfSy11GGmop2TNdJCHDI3tBy3I7MoCtRRU1TBW08VRTTRzU8zQ+OWJwc17TzBBHIg9VKgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiALxzQ9pac4IxyOD4r1fL3tjY573BrGjJcTgAdVGUsFr0Np6z6WfpektkQsMjJI5KSVzpRIJCS/c55LnFxcckknmqLRnkx0t5P5KmXT9sNNPUsZFJNLUy1EhjZ6rA6VziGjuaCB7Fe7RqKz6ggmqLTdqC4QQu2SSUlQyVrHdCWkgH3r6s+oLRqGKWW0XShuMUL+HI+jqGTNY74pLScH2LU1SzOBcVJD6x9yjUkPrH3LLNEyIihTx/qlcaa0+GOoflCo+0cuy3+qVxprT4Y6h+UKj7RyA7MVM71ne9VKpnes73qojPFy29jo3uY9pa9pwWkYIPRdSIvzn9Qf0/wDjGj/1LlydkzMd63HrZrzp8Bf+i9ejbGqe57zF/J2xzNG2wOaWnEhwR3GRxB8FjFujvA1ZddRS2WVlovMDKGGMxhtRSthEhZNKwnIbIZZOXrNAjyPSds2evHNDgWuAIPIg969rJcm+HyajJ05u0qmfBRJ51va6a1qtWolzxk0PRUlRDpDSZtz5ai4X/S9PbfMIaVkkjoo4i58jXyTwsaRxu8nPLAPNZn5OqCKTUOqbpJFcKSvM0FNJQ1U/8wG00IG5kcj4S4gBwc0kgOwSOYWY12mLHc7XDaq6zW6qtkAaIqOemY+GMNGGhrCMDA5DA5KotVnt1iomUVqoKWgo2ElsFLC2KNpPbhrQAF91Vcup7+c8z51TqnZ/HJeRqO7U91q9cXp8NLqC5VsXCpa02t8NNTikP5SOFjnVkUrZRuLjKDjL3DaRgCK0Wq83DyYaZit1MKixxUMhrKOruppH1A3H8nIWQTOcAAc7ZW7skO3ArZ130HpW/wBY6tu2mbLcKxwDTPV0EUshA7Buc0nAV2pLbRW+gjoKSkp6ehiZw2U0MYZGxvxQ0DAHsXNt3IRpfmlmA+SumvDaKlrfMKa3aZr6KOoo7fDdDVNpd4Dw1jHUsbmDDsbeI5rcANaAsc15YYaW8a2baaDZU1+mzNUilZ+WqSZ37+Y5ucWAtHf2Ady3NTU0NHTw01NDHDTwsEccUbQ1rGgYDQByAA5YClWqqlfvU9/FNY+fAzSncuvu4NPDyNAaor7JqOa+VmlpaWu07BbLfHdH23EkD4WVjXOiOzkXNp+NlvaGvAI5hZ3oi5Wy9661VctP1NNV2OSjoYjU0T2vp5appn37XN5OcI3QAkdmGjuWxAMdiAADAGAl7COsesBd9uEavLE0derVQXjyU+UKsuNLDVV0FdeJgZmbjFJG6SKJwB7HCJsYB7h2LKNVRXi4aotNwt1mkqKLTTjPO17A19c6WIxlkBcQCWMe5xJw0u2tBzuLc7oLRS22quNRTMLH3CYVE4z6LpAxrNwHcS1jc+7PblVyXoS7o9EvIrUtve3xcml5ttPJf742pNPUWrVPEp4nU3G40s1HDTsiLeIwAkzAZLgAcZOMqbStnkm1tZBdaS7W2voqa5VjIppWQipdPVRSPcBT1EjSxhdt2SOOdzTg4JG1pLRb5YKynkoKV8Fc4uqY3QtLZyQGkvGMOJAA555AKlsmlrFpoTCy2W22wT4MnmVKyHfjsztAz86tNcR3JL9MddMjpme9t8ZNb+VWkko7hJHb5JXy6mpXUtdRN4xM8MQwSzhUlQ5pxNtJ9EYIxz5iGyvq7xbNZVV00f5zPca6BlJbp6eV8Er200bGvJqIIyGAtO57mDGDjccA7Ovel7HqZkLL3Zbdc2QEmJtdSsmEZPaW7gcZwOzovbJpmyaZilislnt9simcHSMoaZkIeR2EhoGSspq60zTmU111tNU6c07/ABSf5PLFLaKeirLddH081VTQMjiuG2gqA2cbe0uaBuDuYdkcxgnM/JYT/FCaFrXebU9yuNPTBpx+RZVzNYB2YAaAB7AFmNRRU9W+GSaGN8tO4vhkcwF0Ti0tLmk9h2ucMjuJHeobRaaWxWukttDGY6SljEcbS4uOB3knmSe0k8yTlarrvzJFTER1r58DAvJ9bNXUV5kk1Ea99uNM4W5kta2U0kfEzw6nGOLPt2Yky8ANIznL5bH5Y7rYYaO+2+O42ajvlRTMNVRVkBbWXWIAmOKlkJ9fO5oc1ku1x9UFbmXhAJBIGR2KXsU9xY195oXU1fbYo9e2ysdFFrG63ClltVLKR51NiKDzd0QPNzY5GvJLeTS15OOZVx1naYLVc77PRT1dVDBST1lzhL2xxwwyubI6APDNwMzoxkncWRh+MbmrdWBnPeiXoiNmHBL2x3kjWn11OG4sGjL+zUunqe4Mgp4GcSWAMppuNCeHI6PdG/a3cw7Mg4HIjkr+gAAwOQCLLcuQlCgIiIUIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiALHtdPtEekLw6/0k1XZBTu88ghjfI58X5w2s9IjHbjuyshRR6ipw5OKqsMvtJ5Xm6GbSXKGqsFNwqvTtIaajZC2TnTOh5l1Rwi/wBLcSRy2t7FsvyU1Gmrx5ZH3Hyew0zNLQ6Yip7g+ig4UJquKDExwwAZWx7s9pA5FdEgBowAAPYvQAOwYXVWkdeKx8/PHazm6JXX/H9oUkPrH3KNSQ+sfcuTOhMiIoU8f6pXGmtPhjqH5QqPtHLst/qlcaa0+GOoflCo+0cgOzFTO9Z3vVSqZ3rO96qIzxa21p5ctK+T+6x26/R3immmkEUEn4NmdFUOwDiN4bh/rAcs81slcc+Vai8w8pF9hF3s9HVB7KqplvHmlAK2F7TiBhl3GpiAwHNOGFzRnmMq041JPUR/lbRv/Unls05pS00NwuVFqGM1pHDpvwPUCYDibCXAtAZjm7DiCWjLQcgG9aV8o1l1lda+32sVvEo6eGpc+ppn0+5kpkDcMkAeCDE7O5o7sZyubdU26yWnyW+Tu8V8dTJRVVbNDT09VSUNKyljqInOduhqd0JIMXouLmYD+RxhpzL+D/Lpq+a4v9RaqB0jrHaLZb6asrKulrZmsAnacSwOe1pIA3AP54HogALto1FXc+vbzZzvuE96XrHM2PePLlouzVF1pJLhUzV9sExmpoqOUFxhxxGse5rY3FuRnDu9U198vekdN6aqr9cDcWU0FZU0LYWUpfJLLBJsfjaS0DPYXObkeC1H5S/JpqyODU1wmoaQWenN4qYp4KovnmFWYtg4Wz0S3h/GPb7FF+DKOm8jlybd5rtcLFS6or4KuhFQ2KOtzWEMfV1GwvZE0jc5zfnBWFSrtL2tJvxlL3NttPulr7JNo3nQeV/Tl2mbHbWXOtD7bLdGGChkLnxxyCNzWxkCQv3HkA3n3FWiD+EBpOpttfcI6e+CnoZm00hntskAMxmZEY2mQNBe1zxubnIAJxyWvfJeZ9S6yqaKLUjp6i3adfQwXe03OG6thY+cOa11RwI2iVoaAGujcXDLi7IWpxcNJP04JaeksrIamrpv/ZsuqmVFxjqW1TGOqmwvpC5kzw38pteA5hJIzhdKbNOumnf+5r04nN1tUt7uS95+x2Tq3Xdk0Q+ysvNUIDeK5lvpuznI7OCckYaMcz3ZHVZKudvKJpW92nU8NdPR6Run8briLKDc6WsqXQQSNe8NAdUFrBiMZEbWAnn3Lb3k703cNJabjtNwqKWYwPdwfNn1DmsjOMNzPLI/kc49LAGAAMLkknRe62Yfb37jo21VHW3H2ju7zK0RFkoREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAFiVt8qGjrxLqBlDqGhnbp5odcpWv/JUwIdzMh9E+o7OCcY54WW9q5e1v5NbxXz+Wu26fsUtNS1dBZ/wfFDT8GCqEG58kURwGnkNpA73AHtRQ5kse3rBvDTXlW0Zq63XK4WjUFJPSWxnEq3yboeAzbuD3CQNIYQCQ7GDg4Km0Z5S9KeUHzsabvUFe+k28aNrXsewO9V214BLT3OAwe4rnzWumdQ+Vyq1teLBp682imdpWG1RQXKkNHLWVDagTOjax4BIDGlmezJwDjms90H59q7yvs1ZBp28WOzUGm2WuRt0onUj5ZzMH7GscAXNYARuHLJ5LrcXWzX6wo8fCed5xPX+30l+XlvFSQ+sfco1JD6x9y4s6EyIihTx/qlcaa0+GOoflCo+0cuy3+qVxprT4Y6h+UKj7RyA7MVM71ne9VKhkYc5HYqiMjXhAOMgHHNeoqQIiIAvA0N7AB38l6iAAADAGAiIgPC0OxkA45jPcvURAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBSQ+sfco1PGzaMntKjKfaIihTx/qlcaa0+GOoflCo+0cuy3+qVxprT4Y6h+UKj7RyA7MREQHm0dAm1vQeCbh1Cbm9R4oBtb0Hgm1vQeCbm9R4pub1HigG1vQeCbW9B4Jub1Him5vUeKAbW9B4Jtb0Hgm5vUeKbm9R4oBtb0Hgm1vQeCbm9R4pub1HigG1vQeCbW9B4Jub1Him5vUeKAbW9B4Jtb0Hgm5vUeKbm9R4oBtb0Hgm1vQeCbm9R4pub1HigG1vQeCbW9B4Jub1Him5vUeKAbW9B4Jtb0Hgm5vUeKbm9R4oBtb0Hgm1vQeCbm9R4pub1HigG1vQeCbW9B4Jub1Him5vUeKAbW9B4Jtb0Hgm5vUeKbm9R4oBtb0Hgm1vQeCbm9R4pub1HigG1vQeCbW9B4Jub1Him5vUeKAbW9B4Jtb0Hgm5vUeKbm9R4oBtb0Hgm1vQeCbm9R4pub1HigG1vQeCbW9B4Jub1Him5vUeKAbW9B4Jtb0Hgm5vUeKbm9R4oBtb0Hgm1vQeCbm9R4pub1HigG1vQeCbW9B4Jub1Him5vUeKAbW9B4Jtb0Hgm5vUeKbm9R4oBtb0Hgm1vQeCbm9R4pub1HigG1vQeCbW9B4Jub1Him5vUeKAbW9B4Jtb0Hgm5vUeKbm9R4oBtb0Hgm1vQeCbm9R4pub1HigG1vQeCbW9B4Jub1Him5vUeKAbW9B4Jtb0Hgm5vUeKbm9R4oAAB2AL1eAg9hC9QBERAeP9UrjTWnwx1D8oVH2jl2W/wBUrjTWnwx1D8oVH2jkB2YoZHnOB2KZUzvWd71URniIrZV6ks1vqH09Xd6CnqGY3RTVLGOGRkZBOexUhc0VE+825lCyvdcKRtDI5rW1BmaI3Eu2gB2cElxwPbyU0dbTS1U1IyohdVQNa6SFrwXsDs7SW9oBwcZ7cFATovNwztyM9uFDVVtNQsY+qqIYGSPbEx0rw0Oe44a0Z7STyA70BOiIgCKCqraahYx9VUQwMke2JjpXhoc9xw1oz2knkB3qYkNBJIAHMkoD1FDSVlPX00VTSTxVFNKNzJYXh7HjqCORCp7herbaTGLhcKSkMgJYKiZse4DGSMkZxkeISNgK5F5uAOMjOM4UNHW01wpo6qjqIqimlGWSwvD2OHUEcigJ0Xm4EkZGR3L5lmjhDTLIxgcdoLnAZPRAfaIoDW0wrBRmoi87MZlEG8cQsBwXbe3GTjKAnRQVVbTULGPqqiGBkj2xMdK8NDnuOGtGe0k8gO9TBwOcEHHI4QHqIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCnjfuGD2hQKSH1j7lGUmREUKeP9UrjTWnwx1D8oVH2jl2W/1SuNNafDHUPyhUfaOQHZipnes73qpVM71ne9VEZ4uSvKvQTjyjXmubThstXJCZoKipgaYYWRhnFHDrWE7sAgPa3ke0Y59arGrloOy3O5VFxkZXU9ZUhrZ5KG41FLxg0YbvEUjQ4gcsnnjktUuKpI1KaNA61oKeTQmj666UdNHQwCWmD2vpuLUZG+F8by+oZszu3ZduOM+xZZ5EI4KrU1ynt90pZKaltVBBURW5jGwTT/AJYvO1zd7cOJIxtB3HkewbqtNporFbaa226nZT0VMzZFEzOGj5+Z65PMr4p7HQ0l4rrtFCW19dHFFPJvJ3tj3bBjOBje7s6rbtNa39e/UHO5gvt6rkalgqbTU6j8ql41PB53R0D6e3NpjGZCIGsDwxrRzO+RwPtOFhFktTNP6et9tuGn3WrUJv8Abql7uE7bLTyVTXxt4mMZYXFhbnILei6CoNKU9t1XddQU08rJLrDFHU0/LY58eQ2Qd4dtO09xwFX3ix0N+hp4a+Eyx09RFVRgPLcSRuDmHkR2EDl2KUVXXS913hE+ev7mqqbyqW+eKheWr7Gi75b6Cya0rRcNLaS1E/U11e6lqqyta19O0Qt9CTMD9o/Ju7CebuxbZqJ7FQ6Deb1T22kscNOGVFNTP41NG0ctjSGt3DsAAaD7FUSeT3SMsjpJNK2N73kuc51viJJPaSdqvFttVBZ6VtJbqKmo6VpLhDTRNjYCe07WgBZmaFT1uLH13jQTdNMtllF6itJs9JetT2t9Jb3M2Php45WtYXt/Nc4lziO7IVRQ6Ov0uu6unbJC+mtt8mvE0ba3IZHO2Xh/keHkPIPbvPYeS3heLHQ36Gnhr4TLHT1EVVGA8txJG4OYeRHYQOXYrHePJrpu+3ae61lLVivqGsZLLT3Cpg3how3IjkaOQ9it6V1qimP/AFJd6+9XszALVZ7/AGv+DxPSSy1slwNAS2nEHCmp4uW6IAAOJDQ7mfS5+5aY1OyAVDn2iOvZpZ8tWbSJQG03C3U2/hNPpj8pvzkbezHeuxbRaaWx2+GgohKKaHOwSzPmdzJJy95LjzPeVb9R6OsmrWRNvFC2p4LXNjdvc0sDi0kAtI7Sxvgtq0i1v9dfzvMuiaFTu66+xq+93W5WzyoS3masp3VFJEKJ9qo6OqrDHQOcXGaSSKJ3DlJAcGkYwMZ71kfkpp6mw+Se3Sx1sVxl4DpoWzVDGQNyTiMSNZ6Lc97g5wJOegyW7aDsd6uMtxqIauGtmY2OWairp6UzNGcB/Ce3fjJxnOFcqfTtpprJHZI7dTfgiOMRCkfGHxlo7i05zz58+081ifou7f8APP1Nx9d7Z115Gm7LU3uk1/erpHd7PVX58MAuNBLNwqHzd27gtp5wC8yM2O3FzMHf3KPW9bQUHlJuzrhYbRqqOptbauFlfURgUDIfRfG0Pa4ekXB3LDu3keS2t/J1o7/9J2H/APx0P/pUk+gdJ1Uz5p9MWSWZ5y58lBE5zj7SW80lYd3XW0kPHvj26/waalfVXPyKWvS1raa68aoY59LTxlzoaCBzxKWOkPqtiY5refPJAAwrzFef4zatgqpLDJPM7S08VRaK5nC3yNqGB8R3jBGQRnmCtvWqx2uxQvgtVto6CF7t7o6SBsTXOxjJDQMnkviSw0E15F4fATcBSuoxJvd/NFwcW4zjtA54yrVVLbXf6Ne/v3BUuEn3eqb9DRFLofzrR0N9sUFLDVagvNuqGssdIeDQwRSgNLWuYN230nOc5oBPdgLOdMtls/ll1La45TLTV9qprjOcADzhrjEX4HIFzQCcdFd7f5INJWo0/mVJcIG07g6ONl2q9jSDkejxcYz3Ywr1aNKU9q1FfL6ZpJ6+7GNry8ACKONuGxtx3cyT1JVvqXGrH0SS4T9iXXGOvD1l8i/oiLmbCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIApIfWPuUakh9Y+5RlJkRFCnj/VK401p8MdQ/KFR9o5dlv9UrjTWnwx1D8oVH2jkB2YqZ3rO96qVTO9Z3vVRGeIixXV+u6DR9Ra4ana99ZMBLh+PNaf8+pfgHEbSWguOANwyQqsWkTZJlSLH9RaiqLXNbKK20UVdcrm97adk1RwIcMYXuc6QNeRyHINa4knswCR8WHUNfWXSstF4t1NRXOmhjqcUlWamF8T3PaCHuZG4OzG7ILQOzBPPCBJkaLBqvXz36Yr7tRU8IkpLv8Agste/e04rG07ncsYOCSB3HHap7hqm+27VdrtD7Pa5KS5TvZFLFc5DUthawudM6EwBoaPRB/KYy9ozkgKpNx3huJ7vYzJFjOgtU/xw0xSXN/DFQ/c2ZkTXBrHgkYGfZjvKyZQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEVi1fqKTS1iqLlDaa67Tx4bHR0LQXvcezJJAa0dpcTgDmsLb5aqNvkjtevJbXKJ7rG0Udpil3yzzvcWsha7aMkkZJxyAJxyTY6t0cdRdqW/2NootKX3y9m06G0Dq0WmBlLqaYNmp562OLgNMMjxiaQsZnLAMuwOfXCu/kr8sH8p2otS0UNFTU9FaYaV8b4a2GqL3S8XcC+F7mcuGOQOefNa0dWPdrMX1Ce/HzNqKSH1j7lGpIfWPuWGbJkRFCnj/VK401p8MdQ/KFR9o5dlv9UrjTWnwx1D8oVH2jkB2YqZ3rO96qVTO9Z3vVRGeLALtp3UD7m6ps9ssFEJK1tRVTMqpI5bhG1rmhk22HnkEdpcBjvWfoqtck2Qayv2h7pddN0FlpLDpmgpaauhrBBBUyMjZw5WyEMDYBtLwHtJGMbu/sU2ldGX2xu1JGxtstjLzNx2VVFO+eamIijjAAkjAfjY5wLjgF3Ye/Y6K3sGt/8AHJEjFPd17s1ndfJm202Oak046pMU01E+WhfI0skkiqInPqC53PiFjDu5+lgHG7trnNrtMasvNXR6PuF0gr44nito5aUyFwL97HunnY/aMt2tHojJxhZ8iXmI2Fh0XbhadMW6ibR1tG2BhaIK50TpmekfWMTnM9vInl7VfkRRlQREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQFPXwOqqGqgYQHyxOYC7syQQtIWnyBXaLSOiaOfVU1vvWmKCalj8xiinpjJJkOeBNGee07d20EDOMZW90TY1v8A55su1Pd7xyRzqfJfq3SWhND2dtPc9SVdubJTV4tN6kt/Ah4biwRbZqdr9sm3nIHOIyFl3ke0xcrQyI3uzaqpbnSUrYn193v5q46155OcIG1UrGHlnsGM8ituItu0bbb29db9pi4kkt3XW7YFJD6x9yjUkPrH3LmzZMiIoU8f6pXGmtPhjqH5QqPtHLst/qlcaa0+GOoflCo+0cgOzFTO9Z3vVSqZ3rO96qIzxaVvesqi3eUYUk2pa2mY2oEUdE58LYJKZxjMkxlMGxpY4luxz+IQDggvC3UsKq9JX2DUdwu1kvFopYq1sYdBXWp9SY3NzlzXNqI8bs5PLtA5rVDSqTfXWrqVmrGlpGK6f11ef4g2gCirfwlVNkj8+rYpahoaHECfEDZC7t5McYydvMtGCsm0Bf66ZjbLcXVFbLRwgR3SSkqIDVtGBukbJG0Mk6gOcDzIxnaPLl5PZblpe02h1VbWut8he6F1A99vqQQ4bZKZ02XNG7IBkOHNB9iq9E6JdpKSvldPQAVYjaKS1UPmVHFtLjvbDxH4e7f6Tt3MNby5ZNUQ56662Ec4R111tMvREWTQREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAFqnyiaj1LVa901obTN2isk9ypKm4VNzfRtqXsZEWtaxjHejkudzJ7AOS2ssN1t5N6DWtbari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" alt="학생 상세 화면: 지도와 점수 계산" loading="lazy"></div>
    </div>
  </div>
</div></section>

<section id="result"><div class="wrap res">
  <div>
    <h2>가상 명단으로<br>시험해 봤습니다</h2>
    <p class="lead">가상 학생 221명, 4학년 1학기 관리 실습 병원 9곳으로 돌려 보고, 정원만 지키며 무작위로 배정한 경우(100번 평균)와 비교했습니다.</p>
    <div class="big">−21.7분</div>
    <div>학생 1명당 평균 통학시간이 이만큼 줄었습니다.</div>
  </div>
  <div>
    <table class="cmp">
      <thead><tr><th>항목</th><th>무작위 배정</th><th>로켓단 추천</th></tr></thead>
      <tbody>
        <tr><td>평균 통학시간</td><td>56.7분</td><td class="us">35분<small>38% 감소</small></td></tr>
        <tr><td>지망 충족률(1~3지망)</td><td>19.4%</td><td class="us">76%<small>+56.6%p</small></td></tr>
        <tr><td>평균 매칭점수</td><td>41.8점</td><td class="us">75.6점<small>+33.8점</small></td></tr>
        <tr><td>미배정 학생</td><td>-</td><td class="us">0명<small>221명 전원, 정원 준수</small></td></tr>
      </tbody>
    </table>
    <div class="note">통학시간은 지도 기반 예상값입니다. 실제 학기 명단으로 다시 확인할 계획입니다.</div>
  </div>
</div></section>

<section id="try" class="alt"><div class="wrap">
  <h2>비중을 직접 바꿔 보세요</h2>
  <p class="lead">통학을 중요하게 보면 가까운 병원이, 희망을 중요하게 보면 희망 병원이 위로 올라옵니다.</p>
  <div class="demo">
    <div>
      <div class="who"><img class="px" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAB0AAAAwCAYAAADtoXHnAAACzklEQVR42u1YP2gTURj/XbihFLElZJBohmoqOrRDiLEWskkGJ0nRVodwk2MQ6eTQwcGpIGQqnY4MkjTkJnHQbgdRmnBIExFMA0KkUwmpBHHyHOw73728u3t3SbD++eDg7vu+937fv/e99066eGUJDmQiGEleCrIbYOhb3xfa96lZMtYVOOQo4AD2+j1XUFEjQ0HCqih5KEreTcV1DomTU5O1WFHyUNUCev0ewrNhGx8AVLXAC7MUGJQAkvDSoDydsYDS+eQBuhSU5LuQWBIFDFpIE6fTDdoobQjxhFqWSCGlLkes972PR1xeo7SB5NoTtpBMXouU/Xrng0wAqOwUAQB37uasFhkSAUwoBcxcvzekw/JZIys7RSQW5pFYmLfAAZghUQ8vXb3BNY7lk7EE0DKQAv4t1cstJAAw1HygCRPKz3ZoqHl0ppOWt0azTfLK9dTWvozSS5uwtamhtanBTccK/dcGjGbbBghAkj0tX7tl+z4XjyNye9FVhwUm3nsumaPDz3xBKuwso+h5K4Ts6srJV8EWQUfQzGPNd14pj3ydkWwdpDOdpHOBV0+zQ4bRa5LQ+vpDykuBg9mW9gIAsBibsSYYRFPIpWMYvH5m063VdRT1Ls4c7uH83BxqdR373WMc1HY9G75JPVzKpWMo6l1EohdsT1HvIpeO2XQParuIL9/0Dm+trluM/e6xI/C2DjwIvwUAbPeWhgDHtrVlV1eglauOcq1ctfKnlaueXnruMk7ekvegJPsd4ARGPHZKj29QEmJ2GZCwa+UqBtGUTbZ8LT26pywwAfTIoTRyeOmiEg3lWHLq1mkmBkoT6Vp/xLlXBoCp95VfnLOZv/OEL7P74JaWce3BkziYmewWR9O7T1+w9ei+kNzPOpVG/LMi9Ifl/1Xx3wA1xzCv//9InQ9vnE/sJ7c0AR3p9HUkloxm23OgiI7vq+I4up2T4AfMHhxmQ6yU0AAAAABJRU5ErkJggg==" alt=""><p>예시 학생 한 명<br>1지망 고려대학교구로병원, 2지망 서울특별시서남병원</p></div>
      <div class="wl"><span id="dLblP">희망 병원 30%</span><span id="dLblD">통학 70%</span></div>
      <input type="range" id="dSlider" min="0" max="100" step="5" value="70" aria-label="통학과 희망 병원의 비중">
    </div>
    <div id="dRows"></div>
  </div>
</div></section>

<section id="faq"><div class="wrap">
  <h2>자주 묻는 질문</h2>
  __GUIDE_BTN__
  <div class="faq">
    <details><summary>학생 명단은 서버에 저장되나요?</summary><p>저장하지 않습니다. 올린 파일은 메모리에서 처리한 뒤 사라지고, 결과는 엑셀로 내려받아 보관합니다.</p></details>
    <details><summary>명단 파일은 어떤 형식이어야 하나요?</summary><p>엑셀(.xlsx) 또는 CSV입니다. 필수 열은 학번, 이름, 성별, GPA, 출생연도이며, 프로그램 안에서 양식 샘플을 내려받을 수 있습니다.</p></details>
    <details><summary>통학시간 앞의 '약'은 무슨 뜻인가요?</summary><p>지도 위치를 바탕으로 계산한 예상 소요시간이라는 뜻입니다. '추정'은 주소 정보가 부족해 규칙으로 계산한 값입니다.</p></details>
    <details><summary>지도에 표시되는 거주지는 실제 집인가요?</summary><p>아닙니다. 주소가 동 단위까지만 있어서, 동네 기준의 대략적인 위치로 표시합니다.</p></details>
    <details><summary>배정 결과를 그대로 확정하나요?</summary><p>아닙니다. 프로그램은 추천안과 근거를 보여 주고, 검토와 최종 결정은 교수님이 합니다.</p></details>
    <details><summary>미배정 학생이 나오면 어떻게 하나요?</summary><p>미배정·사유 탭에서 이유(정원 부족, 조건 불일치 등)를 확인하고, 정원이나 조건을 조정해 다시 배정하면 됩니다.</p></details>
  </div>
</div></section>

<section id="team" class="alt"><div class="wrap">
  <h2>만든 사람들</h2>
  <p class="lead">안산대학교 간호학과 팀 로켓단입니다.</p>
  <div class="team">
    <div class="mem"><img class="px" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACEAAAAzCAYAAAAZ+mH/AAADdklEQVR42tVYTUhUURT+3jRIhGiIiyGcGMbJxkU2WJkKA2KzbCGWBi2GiAhxISaznIWLljLJW8RjaNPbqSktWhREGANqWoPlIsmfEUdjFiIpLkSE12J8z3vf3Pc7byY7m3d/zr33u+d+55z7LucPtsKCSHKhw7NPdUznqsgqZ2VSt1UA5OKemkrkdg8oUCdgJCtAXFYQswB4aio1dUoCghTZAvK3GHHhDMiZAGGFmJjOVSHg3gIA1NVS3oCtnTwXVo/rSgJCIisfIhFd5fr3y+pxhl7CmYgTkh3GE+7KFWsJiXRJPS8xmIMrxhJMK2zt7FOcYAGUwZmxhtuOmdWkzO0eFAArmXfIHsACYxeAZRDqhZICr6n7tG/AeRBJgacmPqyoUcrRgWGlLPLDpbWE1u7tLGw7bCcFHg1dcZy7eqegT6vdERCsndc3trGjJdGeFHhdzli2hNnJ7IphsAKAX2+f25q8oStu6rpnZAlq8OGnRapzLi5gLi5AT8eJ3EHJ+c4QVfcEAvA9iujqOOaifzY3mO0XOwOafU5ywjYvzPLBCISkN3Cef0LVWwZewQq/TB+HMPUOANDkrT6NjqksomEvMPua0p1ZSJ32nciP7B4AoK/7riVOSEbbiYa9EFNZDPl9VHtCBcDAspyuJWYWUgU7YQFJpIAh/3oewLpfFwA5Z/utsDOXGgeEsorbzBHoWUMum5HtTAYT46JS7+mNAoDE+YOt0gVfh+bARCJGEVMX2ItvEJ/dKDjOJm81tjMZAEDztSsAgPTSigLEsT8wMZVVvnKZtAAJYC0cRnX/YwDAxLiY58QlX7BoECFfFUI+7XumDKC+sQ1rP2eV9vTSSh7E741l5sD++yHHGUkCMJVFX75ZxOrMR0cWlznAajN00UB7BFNjkwCA7gf3NPX0yHu9YhPfjy4jvbTCJCbnD7ZKRsFKXmBqbFIXiFrUYXtiXMSX+TRutzSTLgq3KoRKtdufSbpRk1oBwJKe3ihisUFlcTlYlfORhGPUOSYniHuAkkW1ckipXmqoo5HPUgZDyuLGPoShh6b6jTK0y8B8nJM7jsUGMTIy+m+yaIVbnx+ucgBovtkCAOBHR5lH4yqHBdJf5/Ws8n88pkrFTH50bK7N8F+UlfXUf+FGOvIREAC4snuHavecZRdlpWAbOlzRTwMlyBuU/AV+2j+v68a6jwAAAABJRU5ErkJggg==" alt="강지우 캐릭터"><b>강지우</b><span>팀장</span></div>
    <div class="mem"><img class="px" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABoAAAAvCAYAAAD9/drQAAACqklEQVR42mNU0rBgwAH+M5AHGLEJsuCzhIeFNBu+/IHrxbCMCZcmUi0hpIeFUJAZmdkQtODcqSPYzEDxFSOWOPpPqm+MzGxQLIMGISNRQYcNvHnzAisfi48YiI4jagMmUnwjIiKBkz/0fITuelJ8Q9Cic2taGBgYGBh0pFjgNLIYTBwmhg8QlZAXTWogSoyioDu3poVBLaCGgVndGcOnyOKEfEV0HClrWpIkTlTQERPm+PSpBdRgFunYiiAGBgaGWxvIswzJEoJFEIqCH/suoEieqJnBcKJmBgM+NdiqCYKpjsPJAIUvoaLCoJDgglcNScn7w6MHWMUFnFRwyuGtdnFU5WTFE674wWYR3nbCqUkpKHyzvDlEtx8wgm7Gui0MDAwMDHqy/IhS4PBjhjhbWQaG4wtR1B47fRghBwWXHn9kYGBgYMgI8sGII4KtnThbWYZFhx8zFCkpoIj3oVmCr1nAqKRh8f/Y6cMYLkL2EbLPipTuQSy5pwS3ZN3KtSjqVKxcUPRbmdoysJCTKfdtWcBgwMDAcOkaLCmiBhPDtgKGS5RUfHG2sgx995TIKjFI9lGcrSzDpWsMDK7Fi+FirhiqwuGs3b2x9K3KWSjR3L39Pk65TV9mMRwObaeORaWeirjlGNop85HI04MMTj4JDH9v7sWZMEo9FRmeBwYySK5fT3sfiXXkMPy9uZfBySeBgWFODcQiK1NbjCKIkjiKmVXEINaRgxF0jKR0vt5I2zNcmpnI4Fq8mKFUHYciz/UMf2/upV6qwwdgrSNYPsJpEazMoxbAWR9hi6sLDz4xzCiKYpiRokKSJRlz7mD4iJGYuMqYcwevQ6heMpDSMx983ZYhbdF/Kpj7n6i2993rxwn2KohQwzggQYe923L5NuEuChFqiOq2UDsvAQDmxOC3ExOgygAAAABJRU5ErkJggg==" alt="최영락 캐릭터"><b>최영락</b><span>팀원</span></div>
    <div class="mem"><img class="px" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAB0AAAAwCAYAAADtoXHnAAACzklEQVR42u1YP2gTURj/XbihFLElZJBohmoqOrRDiLEWskkGJ0nRVodwk2MQ6eTQwcGpIGQqnY4MkjTkJnHQbgdRmnBIExFMA0KkUwmpBHHyHOw73728u3t3SbD++eDg7vu+937fv/e99066eGUJDmQiGEleCrIbYOhb3xfa96lZMtYVOOQo4AD2+j1XUFEjQ0HCqih5KEreTcV1DomTU5O1WFHyUNUCev0ewrNhGx8AVLXAC7MUGJQAkvDSoDydsYDS+eQBuhSU5LuQWBIFDFpIE6fTDdoobQjxhFqWSCGlLkes972PR1xeo7SB5NoTtpBMXouU/Xrng0wAqOwUAQB37uasFhkSAUwoBcxcvzekw/JZIys7RSQW5pFYmLfAAZghUQ8vXb3BNY7lk7EE0DKQAv4t1cstJAAw1HygCRPKz3ZoqHl0ppOWt0azTfLK9dTWvozSS5uwtamhtanBTccK/dcGjGbbBghAkj0tX7tl+z4XjyNye9FVhwUm3nsumaPDz3xBKuwso+h5K4Ts6srJV8EWQUfQzGPNd14pj3ydkWwdpDOdpHOBV0+zQ4bRa5LQ+vpDykuBg9mW9gIAsBibsSYYRFPIpWMYvH5m063VdRT1Ls4c7uH83BxqdR373WMc1HY9G75JPVzKpWMo6l1EohdsT1HvIpeO2XQParuIL9/0Dm+trluM/e6xI/C2DjwIvwUAbPeWhgDHtrVlV1eglauOcq1ctfKnlaueXnruMk7ekvegJPsd4ARGPHZKj29QEmJ2GZCwa+UqBtGUTbZ8LT26pywwAfTIoTRyeOmiEg3lWHLq1mkmBkoT6Vp/xLlXBoCp95VfnLOZv/OEL7P74JaWce3BkziYmewWR9O7T1+w9ei+kNzPOpVG/LMi9Ifl/1Xx3wA1xzCv//9InQ9vnE/sJ7c0AR3p9HUkloxm23OgiI7vq+I4up2T4AfMHhxmQ6yU0AAAAABJRU5ErkJggg==" alt="정현민 캐릭터"><b>정현민</b><span>팀원</span></div>
    <div class="mem"><img class="px" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABwAAAAxCAYAAADJP8l8AAADCElEQVR42r1YTWgTQRT+VkPpoYdYAoYQ2xKlRrEFMbRpQUELsSBIKihBEIr0HBQ8KJQQtaAHTznvwVMpCpqjFFFRTE31EKyaWJKSlhCWEiSHHIJo1kO66+44M5ndbHyX3Z03M9987837SaRAMAyGqLAnEk/p4oH5oVhCqsCrrWWC7iNAVLtgxBrTXkyGfijaKXWJhI8zAVY/fOOCk3sBgGTwoWqVlXYYFvAeoCTiQyHhMRTxYUcTWmUuBNhrsWzSV2mZqTsXXbBuUrtgjjI8ExpHcjFO1SWXUtbSEC8srDIiTSocFkq9Aa97AAvJFdO4nIzp76Tube4L/Ia1tkwqJ2MYjS4CADbTSyYQmm4rp3R/aUr5NUy7m1SWpM5W4JNybeokpk8ELOtsATZqTbhGQlhvDKJRa+rs5GSMqbMVFuvpCC7OZzDg6WduxNMJM9TKyUR0FQDQytWAaqvzLtVWey5lL14cdozFzYcvTLrR27OWYpAKqJs1xciLn3+1n+P0iJqIy9z+RmI0UaoWX0ap75RN3+6hETPjvbjk9TRMhtTLRLAm2Ah1cC4SpJRfo67MbmwDykvT2PKTFUyODVPnHz42RRKQ9LDgLTRK2ReB3+dpX4pqDWi1uBlKO+jVKzF79dDv86D/xyP9/b9U/ObgLdP3nacF0/eDy0FnKv7k2DDKld2/5q3sCrmha4akdGJERkDXgEY5fXaOOu7zttNcVVGcBQSAeOoxdTwVn7dm0uzGNnWM5cdLfZ/092c/Q/bCwri5nQtjy4eVarsMHTxwiDvPyEooLM7fz1LHE3IGCTnjbCOs+Ut7Gs2nlAvdd96hviJ+fy8CAPYfnQEA3Iv+W7FPzd1FvdAuwhduBpG4Hmbexp4FvlHevX5OVgs6oKFwopSfcfIM1HookenH7pV3pBHu+Q9SWkaxmoEsAWrVOTh7A0cCQyhu7TAXL7+pmvRKuaDfYlFAqcu/vTp2bVwf8tiJ6EUBVQfuhiralwKAymoVjQEtMEcSzjTvP37tSEFkjjDDXmUZAPgDFuQhjuV0nZQAAAAASUVORK5CYII=" alt="강민정 교수님 캐릭터"><b>강민정 교수님</b><span>지도교수 · 간호관리학</span></div>
  </div>
  <div class="end"><p>명단 없이도 샘플 데이터로 바로 해볼 수 있습니다.</p><a class="btn" href="/app">배정 시스템 열기</a></div>
</div></section>
</main>

<footer><div class="wrap">
  <div style="margin-bottom:6px"><a href="/app">배정 시스템 열기</a>__GUIDE_FOOT__</div>
  팀 로켓단 · AI기반 간호실습 배치 의사결정 프로그램 · 안산대학교 간호학과 캡스톤디자인 (__APP_VERSION__)
</div></footer>

<script>
(function () {
  /* 비중 체험: 앱과 같은 환산(15분=100점, 90분 이상=0점) */
  var H = [
    {n: '고려대학교구로병원', t: 47, p: 100, g: '1지망'},
    {n: '서울특별시서남병원', t: 50, p: 70, g: '2지망'},
    {n: '단원병원', t: 21, p: 10, g: '지망 외'},
    {n: '시화병원', t: 28, p: 10, g: '지망 외'}
  ];
  var sl = document.getElementById('dSlider'), box = document.getElementById('dRows');
  function esc(x) { return String(x).replace(/[&<>"']/g, function (c) { return {'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'}[c]; }); }
  function draw() {
    var w = +sl.value / 100;
    document.getElementById('dLblD').textContent = '통학 ' + sl.value + '%';
    document.getElementById('dLblP').textContent = '희망 병원 ' + (100 - sl.value) + '%';
    var r = H.map(function (h) {
      var d = Math.max(0, Math.min(100, 100 - (h.t - 15) / 75 * 100));
      return {n: h.n, t: h.t, g: h.g, s: d * w + h.p * (1 - w)};
    }).sort(function (a, b) { return b.s - a.s; });
    box.innerHTML = r.map(function (x, i) {
      return '<div class="dr' + (i === 0 ? ' top' : '') + '"><div><span class="nm">' + esc(x.n) + (i === 0 ? '<em>추천</em>' : '') + '</span><div class="sub">' + esc(x.g) + ' · 통학 약 ' + x.t + '분</div></div><div class="sc">' + x.s.toFixed(1) + '점</div><div class="tr"><i style="width:' + Math.max(2, x.s) + '%"></i></div></div>';
    }).join('');
  }
  sl.addEventListener('input', draw); draw();
})();
</script>
</body>
</html>
"""


def _base_url(request: Request) -> str:
    host = request.headers.get("x-forwarded-host") or request.headers.get("host") or ""
    local = host.startswith(("localhost", "127.0.0.1"))
    return f"{'http' if local else 'https'}://{host}"


@app.get("/", response_class=HTMLResponse)
def render_landing(request: Request):
    import html as _html
    guide = os.environ.get("GUIDE_URL", "").strip()
    link = f'<a href="{_html.escape(guide, quote=True)}" target="_blank" rel="noopener">사용설명서</a>' if guide.startswith("http") else ""
    btn = ('<div style="margin:6px 0 4px">' + link.replace("<a ", '<a class="btn" ', 1) + '</div>') if link else ''
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
