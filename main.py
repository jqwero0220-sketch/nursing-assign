"""
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
    --font:'Pretendard Variable','Pretendard',-apple-system,'Apple SD Gothic Neo','Noto Sans KR','Malgun Gothic',system-ui,sans-serif;
    /* 예전 이름(화면 스크립트가 쓰는 변수)을 새 색에 연결 */
    --navy:var(--ink); --teal:var(--blue); --teal-d:var(--blue-d); --amber:var(--gold);
    --bg:var(--liner); --card:#fff; --line:var(--hair); --text:var(--ink); --good:var(--green); --bad:var(--red); --warn:var(--gold-d);
  }
  * { box-sizing:border-box; }
  html { -webkit-text-size-adjust:100%; scrollbar-color:var(--blue) var(--liner); }
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
  .brand b { background:var(--blue); color:#fff; font-size:20px; font-weight:900; letter-spacing:-.04em; padding:7px 15px 8px; border-radius:10px; box-shadow:0 3px 0 var(--blue-d); white-space:nowrap; }
  .brand span { font-size:14.5px; font-weight:650; color:var(--muted); letter-spacing:-.01em; }
  .ver { font-size:12.5px; font-weight:800; color:var(--blue-d); background:#fff; padding:4px 11px; border-radius:99px; white-space:nowrap; }

  /* 단계: 상태는 색이 아니라 표시(완료=체크, 현재=노란 번호) */
  .stepper { list-style:none; display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:12px; padding:0 0 26px; margin:0; }
  .stepper li { display:flex; align-items:center; gap:12px; padding:13px 18px; border-radius:12px; border:2px dashed rgba(12,26,58,.28); color:var(--muted); font-weight:750; font-size:15.5px; letter-spacing:-.02em; cursor:default; transition:background .2s, transform .2s var(--ease); }
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
  .card h2 { margin:0 0 6px; font-size:25px; font-weight:850; letter-spacing:-.04em; color:var(--ink); }
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
  @keyframes spin { to { transform:rotate(360deg); } }
  .actions { display:flex; justify-content:space-between; gap:12px; margin-top:30px; flex-wrap:wrap; }

  /* 배정 방식 */
  .mode { display:block; position:relative; border:2px solid rgba(12,26,58,.18); border-radius:var(--r); padding:20px 56px 20px 22px; cursor:pointer; transition:border-color .15s, background .15s, transform .15s var(--ease); height:100%; background:#fff; }
  .mode:hover { border-color:var(--blue); transform:translateY(-1px); }
  .mode input { position:absolute; opacity:0; pointer-events:none; }
  .mode b { display:block; font-size:17px; font-weight:800; letter-spacing:-.03em; color:var(--ink); margin-bottom:4px; }
  .mode > span { font-size:14px; color:var(--muted); display:block; }
  .mode::after { content:''; position:absolute; right:20px; top:20px; width:26px; height:26px; border-radius:50%; border:2px solid rgba(12,26,58,.28); background:#fff; }
  .mode.on { border-color:var(--blue); background:var(--blue-l); box-shadow:0 0 0 3px rgba(62,91,184,.14); }
  .mode.on::after { border-color:var(--blue); background:var(--blue); }
  .mode.on::before { content:''; position:absolute; right:29px; top:28px; width:9px; height:5px; border-left:3px solid #fff; border-bottom:3px solid #fff; transform:rotate(-45deg); z-index:1; }
  .mode:focus-within { box-shadow:0 0 0 4px var(--gold), 0 0 0 6px var(--ink); }
  .mode b .pill { vertical-align:2px; }

  details.adv { margin-top:24px; border:2px solid rgba(12,26,58,.14); border-radius:var(--r); background:#fff; }
  details.adv > summary { cursor:pointer; font-weight:800; font-size:16px; letter-spacing:-.02em; padding:17px 22px; list-style:none; color:var(--ink); display:flex; justify-content:space-between; align-items:center; gap:14px; border-radius:var(--r); }
  details.adv > summary:hover { background:var(--blue-l); }
  details.adv > summary::-webkit-details-marker, details.more > summary::-webkit-details-marker { display:none; }
  details.adv > summary::after, details.more > summary::after { content:''; flex:none; width:26px; height:26px; border-radius:50%; background:linear-gradient(var(--ink),var(--ink)) center/11px 2px no-repeat, linear-gradient(var(--ink),var(--ink)) center/2px 11px no-repeat, var(--blue-l); transition:transform .3s var(--ease); float:none; }
  details.adv[open] > summary::after, details.more[open] > summary::after { transform:rotate(45deg); }
  details.adv .adv-body { padding:6px 22px 24px; border-top:1px solid var(--hair); }
  .sec { margin-top:18px; }
  .adv-body h4 { margin:28px 0 10px; font-size:16px; font-weight:800; letter-spacing:-.03em; color:var(--ink); }
  .adv-body h4:first-child { margin-top:20px; }
  details.more { margin-top:26px; border:2px dashed rgba(12,26,58,.22); border-radius:12px; }
  details.more > summary { cursor:pointer; padding:13px 18px; font-weight:700; font-size:14.5px; color:var(--ink-2); list-style:none; display:flex; justify-content:space-between; align-items:center; gap:12px; border-radius:12px; }
  details.more .more-body { padding:4px 18px 20px; }
  .wbar { display:flex; justify-content:space-between; font-weight:800; font-size:16px; margin-bottom:2px; letter-spacing:-.02em; }
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
  .drop .big { font-size:19px; font-weight:800; letter-spacing:-.03em; color:var(--ink); margin-bottom:6px; }
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
  .hero-t { margin:0 0 18px; font-size:20px; font-weight:850; letter-spacing:-.03em; color:#fff; }
  .bignums { display:flex; gap:30px; flex-wrap:wrap; }
  .bignum .n { font-size:46px; font-weight:900; line-height:1.05; letter-spacing:-.05em; }
  .bignum .n small { font-size:17px; font-weight:800; margin-left:2px; letter-spacing:-.02em; }
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
  .kpi .v { font-size:32px; font-weight:900; color:var(--ink); letter-spacing:-.045em; margin-top:2px; }
  .kpi .v small { font-size:14px; font-weight:700; color:var(--muted); margin-left:4px; letter-spacing:-.01em; }

  .charts { display:grid; grid-template-columns:repeat(auto-fit,minmax(310px,1fr)); gap:12px; margin-bottom:12px; }
  .chart { background:#fff; border-radius:var(--r); box-shadow:var(--lift-s); padding:24px 26px; }
  .chart h3 { margin:0 0 18px; font-size:17px; font-weight:850; letter-spacing:-.03em; color:var(--ink); }
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
  .donut-hole b { font-size:28px; font-weight:900; letter-spacing:-.05em; color:var(--ink); line-height:1; }
  .donut-hole span { font-size:11.5px; color:var(--muted); margin-top:4px; font-weight:650; }
  .legend { font-size:13.5px; font-weight:600; color:var(--ink-2); } .legend div { display:flex; align-items:center; gap:9px; margin:7px 0; }
  .legend i { width:13px; height:13px; border-radius:4px; display:inline-block; }

  .alert { background:var(--gold-l); border:1.5px solid #F0D06A; color:var(--gold-d); border-radius:12px; padding:14px 18px; margin-bottom:14px; font-size:13.5px; font-weight:600; }
  .alert ul { margin:6px 0 0; padding-left:18px; }
  .tabs { display:flex; gap:2px; border-bottom:2px solid var(--hair); padding:0 18px; flex-wrap:wrap; align-items:center; }
  .tab { border:none; background:none; padding:17px 18px; font-weight:800; font-size:15px; letter-spacing:-.02em; color:var(--muted); border-bottom:4px solid transparent; margin-bottom:-2px; cursor:pointer; transition:color .15s, border-color .15s; }
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
  .score { font-weight:900; color:var(--blue); font-size:15.5px; letter-spacing:-.02em; }
  .empty { padding:40px; text-align:center; color:var(--muted); }

  /* 학생 상세: 인쇄된 배정 라벨 */
  .modal { position:fixed; inset:0; background:rgba(12,26,58,.62); display:none; align-items:center; justify-content:center; z-index:50; padding:16px; }
  .modal.open { display:flex; }
  .sheet { background:#fff; border-radius:18px; width:100%; max-width:580px; max-height:92vh; overflow:auto; padding:0 28px 28px; box-shadow:0 28px 70px rgba(0,0,0,.35); position:relative; }
  .sheet-top { display:flex; align-items:center; gap:12px; padding:24px 0 16px; margin-bottom:16px; border-bottom:3px dashed rgba(12,26,58,.2); }
  .sheet-top .bcs { height:26px; width:132px; color:var(--ink); }
  .sheet h3 { margin:0; font-size:23px; font-weight:850; letter-spacing:-.04em; color:var(--ink); }
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
  .auth-card h2 { color:var(--ink); font-weight:850; letter-spacing:-.04em; }
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
    .sheet { padding:0 18px 22px; }
  }
  @media (prefers-reduced-motion:reduce) { *, *::before, *::after { transition:none !important; animation:none !important; } .btn.busy::before { border-top-color:var(--ink); } }
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
  </ol></div></header>

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
  const pc = [['1지망', '#3E5BB8'], ['2지망', '#8EA2E0'], ['3지망', '#CBD6F3'], ['비지망', '#DDE3EE']];
  const pv = pc.map(([l]) => rs.filter(r => l === '비지망' ? r.preference_status.startsWith('비') : r.preference_status.startsWith(l)).length);
  const tot = pv.reduce((a, b) => a + b, 0); let acc = 0;
  const stops = tot ? pc.map(([, col], i) => { const a = acc / tot * 100; acc += pv[i]; return `${col} ${a}% ${acc / tot * 100}%`; }).join(',') : '#F1F4F9 0 100%';
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
  new N.Polyline({map: m, path: [a, b], strokeColor: '#3E5BB8', strokeWeight: 4, strokeOpacity: 0.9, strokeStyle: 'shortdash'});
  new N.Marker({position: a, map: m, icon: {content: pin('#FFD34D', '거주지'), anchor: new N.Point(0, 0)}});
  new N.Marker({position: b, map: m, icon: {content: pin('#3E5BB8', hname), anchor: new N.Point(0, 0)}});
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
  L.polyline([a, b], {color: '#3E5BB8', weight: 3, dashArray: '6 6'}).addTo(m);
  L.circleMarker(a, {radius: 8, color: '#fff', weight: 2, fillColor: '#FFD34D', fillOpacity: 1}).addTo(m).bindTooltip('거주지', {permanent: true, direction: 'top'});
  L.circleMarker(b, {radius: 8, color: '#fff', weight: 2, fillColor: '#3E5BB8', fillOpacity: 1}).addTo(m).bindTooltip(esc(hname), {permanent: true, direction: 'top'});
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
    <div class="sheet-top" aria-hidden="true">${barcode(r.student_id + r.name)}</div>
    <h3>${esc(mName(r.name))} <span class="sub" style="font-weight:500">${esc(mId(r.student_id))}</span></h3>
    <div class="sub">${esc(r.gender)} · GPA ${esc(r.gpa)} · ${esc(r.birth_year)}년생 · ${esc(r.address || '주소 없음')}</div>
    <div class="box" style="background:var(--blue-l)"><b style="color:var(--blue-d)">${esc(r.hospital)}</b> · 병원 내 ${esc(r.rank)}순위 · 종합 <b>${esc(r.composite_score)}점</b></div>
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
<meta name="description" content="통학 거리와 학생 희망을 함께 계산해 간호학과 실습 병원 배정안을 추천합니다. 최종 결정은 교수자가 하고, 시스템은 근거를 설명합니다.">
<link rel="canonical" href="__BASE__/">
<meta property="og:type" content="website">
<meta property="og:title" content="AI기반 간호실습 배치 의사결정 프로그램 | 로켓단">
<meta property="og:description" content="통학 거리와 학생 희망을 함께 계산해 실습 병원 배정안을 추천합니다.">
<meta property="og:url" content="__BASE__/">
<script type="application/ld+json">{"@context":"https://schema.org","@type":"WebApplication","name":"AI기반 간호실습 배치 의사결정 프로그램 (로켓단)","url":"__BASE__/","applicationCategory":"EducationalApplication","inLanguage":"ko","description":"통학 거리와 학생 희망을 함께 계산해 간호학과 실습 병원 배정안을 추천하는 웹 서비스"}</script>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/variable/pretendardvariable-dynamic-subset.min.css">
<script>document.documentElement.className += ' js';</script>
<style>
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
    --font:'Pretendard Variable','Pretendard',-apple-system,'Apple SD Gothic Neo','Noto Sans KR','Malgun Gothic',sans-serif;
  }
  * { box-sizing:border-box; }
  html { scroll-behavior:smooth; scroll-snap-type:y proximity; -webkit-text-size-adjust:100%; scrollbar-color:var(--blue) var(--liner); }
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
  .logo b { font-size:21px; font-weight:900; letter-spacing:-.04em; color:var(--ink); }
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
  .sec-h h2 { font-size:clamp(30px,3.7vw,50px); line-height:1.16; font-weight:850; letter-spacing:-.04em; }
  .sec-h p { margin-top:14px; font-size:18px; line-height:1.6; color:var(--muted); max-width:36em; }
  .on-blue .sec-h p { color:#E3E9FA; }

  /* 라벨: 이 페이지의 유일한 카드 형태 */
  .lb { background:#fff; color:var(--ink); border-radius:var(--r); padding:28px; box-shadow:var(--lift-s); position:relative; }
  .on-blue .lb { box-shadow:var(--lift); }
  .lb h3 { font-size:20px; font-weight:800; letter-spacing:-.03em; line-height:1.3; }
  .lb p { color:var(--muted); font-size:16px; margin-top:8px; }
  .tag { display:inline-flex; align-items:center; height:26px; padding:0 11px; border-radius:99px; background:var(--liner); color:var(--ink-2); font-size:13px; font-weight:700; white-space:nowrap; letter-spacing:-.01em; }
  .tag.hit { background:var(--green-l); color:var(--green); }
  .tag.gold { background:var(--gold-l); color:var(--gold-d); }
  .bc { display:block; color:var(--ink); }

  /* 1. 첫 화면 */
  .hero-grid { display:grid; grid-template-columns:minmax(0,8fr) minmax(0,4fr); gap:12px; align-items:stretch; }
  .lb-main { padding:48px 52px 44px; display:flex; flex-direction:column; }
  .lb-main h1 { font-size:clamp(40px,5vw,68px); line-height:1.12; font-weight:900; letter-spacing:-.045em; }
  .lb-main .sub { margin-top:22px; font-size:19px; line-height:1.6; color:var(--ink-2); max-width:31em; }
  .lb-main .fact { color:var(--ink); margin-top:14px; font-size:19px; font-weight:800; line-height:1.5; letter-spacing:-.02em; max-width:30em; }
  .lb-main .cta { display:flex; flex-wrap:wrap; gap:12px; margin-top:30px; }
  .lb-main .bc { position:absolute; right:30px; bottom:28px; width:118px; height:26px; opacity:.9; }
  .roll { display:grid; grid-template-rows:repeat(4,1fr); gap:12px; }
  .st { padding:18px 22px 16px; display:flex; flex-direction:column; justify-content:center; gap:6px; }
  .st-top { display:flex; justify-content:space-between; align-items:baseline; gap:10px; }
  .st-h { font-size:17px; font-weight:800; letter-spacing:-.03em; line-height:1.3; }
  .st-sc { font-size:22px; font-weight:900; color:var(--blue); letter-spacing:-.03em; font-variant-numeric:tabular-nums; }
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
  .cmp td:first-child { font-weight:800; width:21%; letter-spacing:-.02em; }
  .cmp td.old { color:var(--muted); width:33%; }
  .cmp td.new { background:var(--blue-l); color:var(--ink); font-weight:650; border-bottom-color:rgba(62,91,184,.14); }
  .cmp tbody tr:last-child td.new { border-radius:0 0 10px 10px; }

  /* 4. 숫자: 무작위 배정과 나란히 */
  .res { padding:12px 34px 14px; }
  .rrow { display:grid; grid-template-columns:minmax(0,1.05fr) minmax(0,2.4fr) minmax(0,1.15fr); gap:34px; align-items:center; padding:24px 0; border-bottom:1px solid var(--hair); }
  .rrow:last-child { border-bottom:none; }
  .rrow h3 { font-size:19px; font-weight:800; letter-spacing:-.03em; }
  .rrow .bars { display:grid; gap:10px; }
  .bar-l { display:grid; grid-template-columns:78px 1fr 84px; align-items:center; gap:12px; font-size:14px; font-weight:700; color:var(--muted); }
  .bar-l .trk { height:16px; border-radius:99px; background:var(--liner); overflow:hidden; }
  .bar-l .trk i { display:block; height:100%; border-radius:99px; background:#B9C4D8; width:var(--w); transition:width 1.1s var(--ease); }
  .bar-l.us { color:var(--ink); }
  .bar-l.us .trk i { background:var(--blue); }
  .bar-l .v { text-align:right; font-weight:800; color:var(--ink); font-size:17px; font-variant-numeric:tabular-nums; letter-spacing:-.02em; }
  .bar-l .v small { font-size:12px; font-weight:700; color:var(--muted); margin-left:1px; }
  .js .res:not(.go) .trk i { width:0; }
  .gain { justify-self:start; display:inline-flex; align-items:center; min-height:34px; padding:0 14px; border-radius:99px; background:var(--green-l); color:var(--green); font-size:15px; font-weight:800; letter-spacing:-.02em; }
  .zero { display:flex; align-items:baseline; gap:6px; }
  .zero b { font-size:46px; font-weight:900; letter-spacing:-.05em; line-height:1; color:var(--blue); font-variant-numeric:tabular-nums; }
  .zero > span { font-size:17px; font-weight:700; color:var(--ink-2); }
  .on-blue .note { margin-top:18px; font-size:14px; color:#E3E9FA; max-width:60em; }

  /* 5. 주요 기능 */
  .svc { display:grid; grid-template-columns:minmax(0,7fr) minmax(0,5fr); gap:12px; align-items:stretch; }
  .svc .lb.l { padding:38px 42px; }
  .svc h3 { font-size:clamp(23px,2.3vw,30px); letter-spacing:-.04em; line-height:1.28; }
  .svc ul { margin:22px 0 28px; padding:0; list-style:none; display:grid; gap:10px; }
  .svc li { position:relative; padding-left:26px; font-size:16.5px; color:var(--ink-2); font-weight:550; }
  .svc li::before { content:''; position:absolute; left:2px; top:.55em; width:12px; height:7px; border-left:3px solid var(--blue); border-bottom:3px solid var(--blue); transform:rotate(-45deg); }
  .receipt { padding:32px 32px 26px; display:flex; flex-direction:column; }
  .receipt h4 { margin:0; font-size:17px; font-weight:800; letter-spacing:-.03em; }
  .receipt .total { font-size:60px; font-weight:900; letter-spacing:-.055em; line-height:1; margin:16px 0 4px; font-variant-numeric:tabular-nums; color:var(--blue); }
  .receipt .total small { font-size:20px; letter-spacing:-.02em; color:var(--ink-2); font-weight:800; margin-left:4px; }
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
  .demo .ctl h3 { font-size:24px; letter-spacing:-.04em; }
  .demo .wl { display:flex; justify-content:space-between; gap:12px; margin-top:26px; font-size:16px; font-weight:800; letter-spacing:-.02em; }
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
  .dr .nm { font-size:18px; font-weight:800; letter-spacing:-.03em; display:flex; align-items:center; gap:8px; flex-wrap:wrap; }
  .dr .flag { display:none; align-items:center; height:22px; padding:0 14px 0 9px; background:var(--gold); color:var(--ink); font-size:12px; font-weight:900; font-style:normal; clip-path:polygon(0 0,100% 0,calc(100% - 7px) 50%,100% 100%,0 100%); }
  .dr.top .flag { display:inline-flex; }
  .dr .sub { font-size:13.5px; color:var(--muted); font-weight:600; }
  .dr .sc { font-size:26px; font-weight:900; letter-spacing:-.04em; color:var(--blue); text-align:right; font-variant-numeric:tabular-nums; align-self:center; grid-row:1 / 3; grid-column:2; }
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
  .stp h3 { font-size:22px; font-weight:800; letter-spacing:-.03em; }
  .stp p { color:var(--muted); font-size:16.5px; margin-top:10px; max-width:22em; }
  .steps-cta { margin-top:30px; }

  /* 8. 자주 묻는 질문 */
  .guide-grid { display:grid; grid-template-columns:minmax(0,4fr) minmax(0,8fr); gap:56px; align-items:start; }
  .guide-l { position:sticky; top:120px; }
  .guide-l .sec-h { margin-bottom:26px; }
  .faq { display:grid; gap:10px; }
  .faq details { background:#fff; border-radius:var(--r); box-shadow:var(--lift-s); interpolate-size:allow-keywords; }
  .faq summary { list-style:none; cursor:pointer; padding:21px 26px; font-size:18px; font-weight:750; letter-spacing:-.03em; display:flex; justify-content:space-between; align-items:center; gap:16px; border-radius:var(--r); }
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
  .mem b { display:block; margin-top:18px; font-size:21px; font-weight:850; letter-spacing:-.04em; }
  .mem span { margin-top:6px; font-size:14.5px; color:var(--muted); font-weight:600; }
  .mem .tag { margin-top:10px; }

  /* 마무리 */
  .final { padding:56px 56px 52px; display:grid; grid-template-columns:minmax(0,8fr) auto; gap:30px; align-items:end; }
  .final h2 { font-size:clamp(32px,4.2vw,58px); font-weight:900; letter-spacing:-.045em; line-height:1.14; }
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
</style>
</head>
<body>
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
      <p class="sub">통학 거리와 희망 병원을 함께 계산해 배정안을 추천하고, 근거까지 보여 줍니다. 최종 결정은 교수자가 합니다.</p>
      <p class="fact">가상 학생 221명 시험에서 평균 통학시간이 56.7분에서 35분으로 줄었습니다.</p>
      <div class="cta"><a class="btn" href="/app">배정 시스템 열기</a><a class="btn ghost" href="#start">비중 바꿔 보기</a></div>
      <svg class="bc" data-bc="로켓단-221" viewBox="0 0 120 26" preserveAspectRatio="none" aria-hidden="true"></svg>
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
    <article class="lb m4"><h3>최종 결정은 교수자</h3><p>시스템은 추천안을 제공하고, 검토와 확정은 교수자가 합니다.</p>
      <div class="viz v-sign" aria-hidden="true"><span>추천안</span><span class="arrow"></span><span class="last">교수자 확정</span></div></article>
  </div>
</div></section>

<section class="sh on-liner" id="compare"><div class="wrap">
  <div class="sec-h"><h2>엑셀 수작업과 무엇이 다른가</h2><p>최종 결정은 그대로 교수자가 합니다. 계산과 근거 정리를 시스템이 돕습니다.</p></div>
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
  <div class="sec-h"><h2>시스템이 계산하고, 교수자가 결정합니다</h2><p>최적화 알고리즘(헝가리안 알고리즘)이 병원 정원 안에서 학생 전체의 점수 합이 가장 높은 조합을 찾습니다.</p></div>
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
      <svg class="bc" data-bc="종합-70.1" viewBox="0 0 120 26" preserveAspectRatio="none" aria-hidden="true"></svg>
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
    <div><h2>결정은 교수자가,<br>근거는 시스템이</h2><p>로켓단은 배정을 대신 결정하지 않습니다. 더 합리적으로 선택할 수 있도록 추천안과 그 이유를 제공합니다.</p></div>
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
