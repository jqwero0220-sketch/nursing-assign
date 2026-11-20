"""
AI기반 간호실습 배치 의사결정 프로그램

버전 기록 (수정할 때마다 APP_VERSION 을 올립니다)
  v14.0 소개 페이지를 design-taste-frontend 기준으로 새로 제작(화면 단위 구성, 실제 프로그램 화면, 무작위 배정 비교, 팀 캐릭터, 질문 펼침). 배정 도구 화면·계산은 그대로
  v13.8 v13.3 디자인 기준으로 다듬기 — v13.4~13.7 화면은 사용하지 않음 (제목 위 꼬리표 제거, 예시 표 점수순, 장점 목록형, 숫자 효과·안내 문구 정리, 선택·초점·버튼 반응)
  v13.7 접근성 점검 반영(안내문·초록 글자 대비, 제목 순서, 낡은 버튼 규칙 정리) — 화면 모양은 그대로
  v13.6 v13.4 화면(글래스모피즘·새 색·이동 동작·주아체)으로 복원
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

APP_VERSION = "v14.0"
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
<html lang="ko" data-type="software" data-radius="soft" data-buttons="tonal">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<meta name="page-type" content="software">
<meta name="site-kind" content="software">
<title>간호학과 실습 병원 배정 프로그램 | 로켓단</title>
<meta name="description" content="간호학과 학생 명단을 올리면 통학시간과 희망 병원을 함께 계산해 병원 정원 안에서 실습 배정안을 추천합니다. 최종 결정은 교수자가 합니다.">
<link rel="canonical" href="__BASE__/">
<meta property="og:type" content="website">
<meta property="og:title" content="간호학과 실습 병원 배정 프로그램 | 로켓단">
<meta property="og:description" content="통학시간과 희망 병원을 함께 계산해 실습 병원 배정안을 추천합니다.">
<meta property="og:url" content="__BASE__/">
<script type="application/ld+json">{"@context":"https://schema.org","@type":"WebApplication","name":"AI기반 간호실습 배치 의사결정 프로그램 (로켓단)","url":"__BASE__/","applicationCategory":"EducationalApplication","inLanguage":"ko","description":"통학시간과 희망 병원을 함께 계산해 간호학과 실습 병원 배정안을 추천하는 웹 서비스"}</script>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Reddit+Sans:wght@400;500;600&display=swap">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard-dynamic-subset.min.css">
<style>
/* deal: palette cool-white-slate, type reddit (+ Pretendard for Hangul), hero centred, nav-centre,
   footer-statement, close-word, rhythm dark-band, buttons tonal, mark initial, radius soft */
:root{
  --bg:#f7f9fa; --surface:#ecf0f2; --ink:#1f3038; --muted:#5d6b72; --accent:#0d7480; --accent-ink:#fff;
  --accent-d:#0a5d67; --accent-l:#dcedef; --dark:#13232a; --dark-2:#1b2f37; --dark-ink:#eef4f6; --dark-muted:#a9bbc2;
  --line:rgba(31,48,56,.13); --text:rgba(31,48,56,.86);
  --font:'Reddit Sans','Pretendard',-apple-system,'Apple SD Gothic Neo','Noto Sans KR',sans-serif;
  --nav:72px; --padt:120px; --padb:72px; --gut:80px; --hgap:48px; --hpadb:64px; --bar:0px;
  --r-lg:20px; --r-md:12px; --r-sm:8px; --r-btn:10px; --btn-h:44px;
  --t-h1:72px; --t-h2:48px; --t-h3:28px; --t-lead:20px; --t-card:20px; --t-body:16px; --t-small:14px; --t-label:13px; --t-stat:38px;
  --m-in:140ms; --m-out:200ms; --m-press:90ms; --m-image:600ms; --m-ease:cubic-bezier(.2,.8,.2,1);
}
@media (max-width:639px){ :root{ --nav:56px; --padt:96px; --padb:48px; --gut:24px; --hgap:40px; --hpadb:24px;
  --t-h1:38px; --t-h2:30px; --t-h3:24px; --t-lead:18px; --t-card:18px; --t-stat:28px; --r-lg:18px; } }
@media (min-width:640px) and (max-width:1023px){ :root{ --gut:48px; --t-h1:56px; --t-h2:40px; } }

*{box-sizing:border-box} html,body{margin:0}
html{scroll-padding-top:0;scroll-behavior:auto;accent-color:var(--accent);scrollbar-color:#b7c5cb transparent}
body{font-family:var(--font);font-size:var(--t-body);line-height:1.625;color:var(--ink);background:var(--bg);word-break:keep-all;overflow-wrap:break-word;-webkit-font-smoothing:antialiased}
::selection{background:var(--accent-l);color:var(--ink)}
:focus-visible{outline:2px solid var(--accent);outline-offset:2px;border-radius:4px;transition:outline-offset var(--m-out) var(--m-ease)}
h1,h2,h3,p{margin:0} h1,h2{text-wrap:balance}
.num,.tab{font-variant-numeric:tabular-nums}
.wrap{width:min(100% - 2 * var(--gut),1280px);margin:0 auto;position:relative;z-index:1}
.skip{position:absolute;width:1px;height:1px;clip-path:inset(50%);overflow:hidden}
.skip:focus{width:auto;height:auto;clip-path:none;background:var(--bg);padding:8px 12px;z-index:20}

/* ---------- buttons (tonal) + motion layer ---------- */
.btn{display:inline-flex;align-items:center;justify-content:center;gap:8px;height:var(--btn-h);padding:0 24px;border-radius:var(--r-btn);font:500 16px/1 var(--font);text-decoration:none;border:0;cursor:pointer;white-space:nowrap;
  transition:transform var(--m-out) var(--m-ease),background-color var(--m-out) var(--m-ease),box-shadow var(--m-out) var(--m-ease)}
.btn--primary{background:var(--accent);color:var(--accent-ink)}
.btn--secondary{background:var(--accent-l);color:var(--accent-d)}
.btn--xl{height:56px;padding:0 32px;font-size:18px}
@media (hover:hover) and (pointer:fine){
  .btn:hover{transform:translateY(-1px);transition-duration:var(--m-in)}
  .btn--primary:hover{background:color-mix(in srgb,var(--accent) 86%,var(--ink))}
  .btn--secondary:hover{background:color-mix(in srgb,var(--accent-l) 86%,var(--accent))}
  .btn--xl:hover{transform:translateY(-2px)}
}
.btn:active{transform:translateY(1px) scale(.985);transition-duration:var(--m-press)}
.btn-row{display:flex;gap:12px;flex-wrap:wrap}
.link{color:var(--accent-d);text-decoration:none;background:linear-gradient(currentColor,currentColor) 0 100%/0 1px no-repeat,linear-gradient(var(--line),var(--line)) 0 100%/100% 1px no-repeat;padding-bottom:2px;transition:background-size var(--m-out) var(--m-ease)}
.link:hover{background-size:100% 1px,100% 1px;transition-duration:var(--m-in)}
@media (prefers-reduced-motion:reduce){*{transition:none!important;animation:none!important}}

/* ---------- nav (nav-centre) ---------- */
header.nav{position:fixed;inset:0 0 auto 0;height:var(--nav);background:var(--bg);border-bottom:1px solid var(--line);z-index:10}
header.nav .wrap{height:100%;display:grid;grid-template-columns:1fr auto 1fr;align-items:center;gap:24px}
.brand{display:flex;align-items:center;gap:10px;text-decoration:none;color:var(--ink);font-size:20px;font-weight:500;letter-spacing:-.02em}
.mark--initial{display:grid;place-items:center;width:30px;height:30px;border-radius:8px;background:var(--accent);color:var(--accent-ink);font-size:15px;font-weight:600}
.ver{font-size:13px;font-weight:500;color:var(--accent-d);background:var(--accent-l);padding:2px 8px;border-radius:99px}
.nav__words{display:flex;gap:32px}
.nav__words a{font-size:15px;font-weight:500;color:var(--muted);text-decoration:none;white-space:nowrap;background:linear-gradient(currentColor,currentColor) 0 calc(100% + 0px)/0 1px no-repeat;padding-bottom:4px;transition:color var(--m-out),background-size var(--m-out) var(--m-ease)}
.nav__words a:hover{color:var(--ink);background-size:100% 1px;transition-duration:var(--m-in)}
.nav__act{justify-self:end}

/* ---------- section box ---------- */
main>section,footer.closing{position:relative;height:calc(var(--h,1) * 100svh);padding:var(--padt) 0 calc(var(--padb) + var(--bar));overflow:hidden}
@media (max-width:639px){ main>section:not(.hero),footer.closing{height:calc(var(--hm,var(--h,1)) * 100lvh)} }
@media (max-width:1279px){ .faq .q5{display:none} .panel .x1{display:none} }
@media (min-width:640px) and (max-width:1023px){ main>section:not(.hero),footer.closing{height:calc(var(--ht,var(--h,1)) * 100svh)} #weights .who{display:none} #weights .frame{padding:16px} #weights .card{padding:24px} .wl{margin-top:20px} .dr{margin-bottom:8px}
  #steps .cards3{grid-auto-rows:1fr}
  .split--faq>.faq{grid-column:span 6!important} .split--faq>.frame{grid-column:7 / span 6!important;padding:16px} .sheet .t-hide{display:none} .faq .q5{display:block} #faq .wrap{height:100%;display:flex;flex-direction:column} #faq .split--faq{flex:1;align-items:stretch} }
.ground-surface{background:var(--surface)}
.ground-inverse{background:var(--dark);color:var(--dark-ink)}
.h2{font-size:var(--t-h2);line-height:1.15;font-weight:500;letter-spacing:-.03em}
.head{display:grid;grid-template-columns:repeat(12,minmax(0,1fr));gap:24px;align-items:start}
.head .h2{grid-column:span 6}
.head p{grid-column:7 / span 6;font-size:18px;line-height:1.6;color:var(--text);padding-top:8px}
.ground-inverse .head p{color:var(--dark-muted)}
.muted{color:var(--muted)}
@media (max-width:1023px){ .head{grid-template-columns:1fr;gap:16px} .head .h2,.head p{grid-column:auto} .head p{padding-top:0} }

/* ---------- UI kit (the app's own look) ---------- */
.ui{--u-bg:#fff;--u-line:#e3e9ec;--u-border:rgba(31,48,56,.14);--u-head:#f3f6f7;background:var(--u-bg);border:1px solid var(--u-border);border-radius:16px;box-shadow:0 18px 48px -12px rgba(19,35,42,.18);overflow:hidden;font-size:14px;color:var(--ink)}
.ui-bar{display:flex;align-items:center;gap:12px;height:48px;padding:0 20px;border-bottom:1px solid var(--u-line);background:var(--u-head);font-size:13px;color:var(--muted)}
.ui-bar b{color:var(--ink);font-weight:500;font-size:14px}
.ui-bar .sp{flex:1}
.ui-tabs{display:flex;gap:4px}
.ui-tabs span{padding:13px 4px;margin-left:12px}
.ui-tabs span.on{color:var(--ink);box-shadow:inset 0 -2px 0 var(--ink);border-radius:0}
.chip{display:inline-block;font-size:12px;font-weight:500;padding:2px 9px;border-radius:99px;white-space:nowrap}
.chip--hit{background:#d7f0ea;color:#0b5c4f}
.chip--two{background:#e2edf8;color:#244e7a}
.chip--miss{background:#eef1f3;color:#4b5a61}
.kpis{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px;padding:20px 20px 0}
.kpi{border:1px solid var(--u-line);border-radius:10px;padding:12px 16px}
.kpi small{display:block;font-size:12px;color:var(--muted)}
.kpi b{display:block;font-size:24px;font-weight:500;letter-spacing:-.02em;margin-top:2px}
.kpi i{font-style:normal;font-size:12px;color:#0b5c4f}
.tbl{width:100%;border-collapse:collapse}
.tbl th{font-size:12px;font-weight:500;color:#4a585e;text-align:left;padding:10px 12px;border-bottom:1px solid var(--u-line);white-space:nowrap}
.tbl td{padding:11px 12px;font-size:14px;border-bottom:1px solid #eef2f4;white-space:nowrap}
.tbl td.r,.tbl th.r{text-align:right}
.tbl .sc{font-weight:500;color:var(--accent-d)}
.fade-b{-webkit-mask-image:linear-gradient(#000 82%,transparent);mask-image:linear-gradient(#000 82%,transparent)}

/* ---------- hero (statement followed by the work) ---------- */
.hero{height:100svh;padding:calc(var(--nav) + var(--hgap)) 0 calc(var(--hpadb) + var(--bar));overflow:hidden;
  background:linear-gradient(180deg,var(--bg) 0,var(--bg) 45%,#e3eff0 100%)}
.hero .wrap,.hero-text{height:100%}
.hero-text{display:flex;flex-direction:column}
.hero__head{display:grid;grid-template-columns:repeat(12,minmax(0,1fr));gap:24px;align-items:stretch}
.hero__head>div:first-child{grid-column:span 6}
.hero__head>div:last-child{grid-column:7 / span 6}
.notice{font-size:14px;color:var(--text);margin-bottom:16px}
.notice b{font-weight:500;color:var(--accent-d)}
.hero h1{font-size:var(--t-h1);line-height:1.08;font-weight:500;letter-spacing:-.035em}
.hero__lead{padding-top:37px;font-size:var(--t-lead);line-height:1.55;color:var(--text)}
.hero .btn-row{margin-top:24px}
.hero__frame{margin-top:48px;flex:1;min-height:0;overflow:hidden;border-radius:16px 16px 0 0;
  -webkit-mask-image:linear-gradient(#000 72%,transparent);mask-image:linear-gradient(#000 72%,transparent)}
.hero__frame .ui{height:100%}
.ui{text-align:left}
@media (max-width:1023px){ .hero__head{grid-template-columns:1fr;gap:20px} .hero__head>div{grid-column:auto!important} }
/* ---------- app frame (section 2) ---------- */
.app{display:grid;grid-template-columns:minmax(0,5fr) minmax(0,7fr);margin-top:48px;flex:1;min-height:0}
.app__pane{position:relative;min-width:0;min-height:0;overflow:hidden;-webkit-mask-image:linear-gradient(#000 calc(100% - 72px),transparent);mask-image:linear-gradient(#000 calc(100% - 72px),transparent)}
.app__pane+.app__pane{border-left:1px solid var(--u-line)}
.app__pane::after{content:"";position:absolute;inset:auto 0 0 0;height:64px;background:linear-gradient(transparent,#fff);pointer-events:none}
.stu{display:grid;grid-template-columns:1fr auto;gap:2px 12px;padding:12px 20px;border-bottom:1px solid #eef2f4}
.stu.on{background:#e6f2f3}
.stu b{font-weight:500}
.stu small{grid-column:1;font-size:12px;color:#4a585e}
.stu .sc{grid-row:1 / 3;grid-column:2;align-self:center;font-weight:500;color:var(--accent-d)}
.open{padding:20px 24px;display:grid;grid-template-columns:1fr 1fr;gap:16px 24px;align-content:start}
.open .ttl{margin:0;font-size:18px;font-weight:500;grid-column:1 / -1}
.open .ttl small{font-size:13px;color:var(--muted);font-weight:400;margin-left:8px}
.box{border:1px solid var(--u-line);border-radius:12px;padding:14px 16px}
.box>small{display:block;font-size:12px;color:var(--muted);margin-bottom:8px}
.bar{display:flex;height:10px;border-radius:5px;overflow:hidden;background:#e9eef0;margin:6px 0 10px}
.bar i{display:block;height:100%}
.row{display:flex;justify-content:space-between;font-size:13px;padding:3px 0}
.row b{font-weight:500}
.map{grid-column:1 / -1;height:180px;border-radius:12px;position:relative;overflow:hidden;
  background:#eef3ef}
.map .road{position:absolute;background:#fff;border:1px solid #dfe6e2}
.map .park{position:absolute;background:#dbe9dc;border-radius:8px}
.map .pin{position:absolute;display:flex;align-items:center;gap:6px;font-size:12px;font-weight:500;background:#fff;border:1px solid var(--u-line);border-radius:99px;padding:3px 10px 3px 4px;box-shadow:0 2px 6px rgba(19,35,42,.12)}
.map .pin i{width:14px;height:14px;border-radius:50%;display:block}
.map svg{position:absolute;inset:0}
.alts .row{padding:6px 0;border-bottom:1px solid #eef2f4}
.alts .row:last-child{border:0}

/* ---------- weights (split) ---------- */
.split{display:grid;grid-template-columns:repeat(12,minmax(0,1fr));gap:24px;margin-top:48px;align-items:stretch}
.split>.card{grid-column:span 6}
.split>.frame{grid-column:7 / span 6}
.card{background:#fff;border:1px solid var(--line);border-radius:var(--r-lg);padding:32px}
.card h3{font-size:var(--t-card);font-weight:500;line-height:1.35}
.card p{margin-top:8px;color:var(--text)}
.who{display:flex;align-items:center;gap:16px;margin-top:24px}
.who img{height:72px;image-rendering:pixelated}
.who p{margin:0;font-size:15px}
.wl{display:flex;justify-content:space-between;font-size:14px;font-weight:500;margin:32px 0 10px}
input[type=range]{width:100%;height:28px;margin:0}
.frame{border-radius:var(--r-lg);padding:24px;background:linear-gradient(160deg,#d9ecee,#eef4f1 60%,#e7eef3)}
.dr{display:grid;grid-template-columns:1fr auto;gap:4px 12px;padding:14px 16px;border-radius:12px;background:#fff;border:1px solid var(--u-line);margin-bottom:10px;transition:border-color .2s,box-shadow .2s}
.dr:last-child{margin-bottom:0}
.dr.top{border-color:var(--accent);box-shadow:0 0 0 1px var(--accent)}
.dr .nm{font-weight:500}
.dr .nm em{font-style:normal;font-size:12px;background:var(--accent);color:#fff;padding:1px 8px;border-radius:99px;margin-left:8px}
.dr .sub{font-size:13px;color:var(--muted)}
.dr .sc{font-weight:500;font-size:18px;color:var(--accent-d);text-align:right;grid-row:1 / 3;grid-column:2;align-self:center}
.dr .tr{grid-column:1 / 3;height:6px;border-radius:3px;background:#e7edef;overflow:hidden}
.dr .tr i{display:block;height:100%;background:var(--accent);transition:width .25s var(--m-ease)}

/* ---------- results (dark band) ---------- */
.res{display:grid;grid-template-columns:repeat(12,minmax(0,1fr));gap:24px;margin-top:48px}
.res .chart{grid-column:span 6;background:var(--dark-2);border:1px solid rgba(238,244,246,.08);border-radius:var(--r-lg);padding:32px}
.res .side{grid-column:7 / span 6;display:flex;flex-direction:column;gap:24px}
.cmp{display:grid;grid-template-columns:150px 1fr 88px;gap:10px 20px;align-items:center;padding:18px 0;border-bottom:1px solid rgba(238,244,246,.1)}
.cmp:first-child{padding-top:0} .cmp:last-child{border:0;padding-bottom:0}
.cmp>b{font-weight:500;grid-row:1 / 3}
.cmp .ln{height:12px;border-radius:6px;background:#3d5560}
.cmp .ln.us{background:#5fc2c9}
.cmp span{font-size:15px;text-align:right;color:var(--dark-muted)}
.cmp span.us{color:var(--dark-ink);font-weight:500}
.legend{display:flex;gap:20px;font-size:13px;color:var(--dark-muted);margin-bottom:24px}
.legend i{display:inline-block;width:10px;height:10px;border-radius:3px;margin-right:6px;vertical-align:0}
.big{background:var(--dark-2);border:1px solid rgba(238,244,246,.08);border-radius:var(--r-lg);padding:28px 32px;flex:1;display:flex;flex-direction:column;justify-content:center}
.big b{display:block;font-size:var(--t-stat);font-weight:500;letter-spacing:-.02em;color:#fff}
.big p{color:var(--dark-muted);margin-top:4px}
.fn{margin-top:24px;font-size:13px;color:var(--dark-muted)}

/* ---------- steps ---------- */
.cards3{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:24px;margin-top:48px;flex:1;min-height:0;grid-template-rows:minmax(0,1fr)}
.cards3 .card{display:flex;flex-direction:column;padding:32px 32px 0;overflow:hidden;min-height:0}
.step{font-size:14px;font-weight:500;color:var(--accent-d)}
.cards3 h3{margin-top:8px}
.panel{margin-top:32px;flex:1;min-height:0;overflow:hidden;-webkit-mask-image:linear-gradient(#000 calc(100% - 72px),transparent);mask-image:linear-gradient(#000 calc(100% - 72px),transparent);border-radius:12px 12px 0 0;background:linear-gradient(160deg,#d9ecee,#eef4f1 70%);padding:20px 20px 0;display:flex;align-items:flex-start}
.panel .ui{max-height:100%;width:100%;border-radius:10px 10px 0 0;border-bottom:0;box-shadow:0 10px 30px -10px rgba(19,35,42,.22)}
.f-row{display:flex;justify-content:space-between;align-items:center;padding:10px 14px;border-bottom:1px solid #eef2f4;font-size:13px}
.f-row>span:first-child{color:#4a585e}
.f-sel{border:1px solid var(--u-line);border-radius:8px;padding:4px 10px;font-weight:500;color:var(--ink)}
.drop{margin:14px;border:1.5px dashed #b9c8ce;border-radius:10px;padding:18px;text-align:center;font-size:13px;color:var(--muted)}
.file{display:flex;align-items:center;gap:10px;margin:0 14px 14px;padding:10px 12px;border:1px solid var(--u-line);border-radius:10px;font-size:13px}
.xl{width:28px;height:32px;border-radius:4px;background:#107c41;color:#fff;font-size:11px;font-weight:600;display:grid;place-items:center}
.file b{font-weight:500}
.file small{display:block;color:var(--muted);font-size:12px}
.ok{margin-left:auto;color:#0b5c4f;font-size:12px;font-weight:500}
.dl{display:flex;justify-content:flex-end;padding:12px 14px}
.dl span{background:var(--accent);color:#fff;border-radius:8px;padding:6px 12px;font-size:12px;font-weight:500}

/* ---------- team ---------- */
#team .wrap,#steps .wrap,#app-view .wrap{height:100%;display:flex;flex-direction:column}
.crew{flex:1;min-height:0;margin-top:48px;border-radius:var(--r-lg);background:linear-gradient(180deg,#dcecee 0,#e9f1f0 70%,#d5e2df 70%,#e2ebe8 100%);display:flex;align-items:flex-end;justify-content:center;padding:32px 32px 48px;overflow:hidden}
.crew img{height:100%;max-height:360px;max-width:100%;object-fit:contain;image-rendering:pixelated;image-rendering:crisp-edges;display:block}
/* ---------- faq (open) ---------- */
.faq{display:flex;flex-direction:column}
.faq div{padding:22px 0;border-top:1px solid var(--line)}
.split--faq{align-items:start} .split--faq>.faq{grid-column:span 6}
.frame--sheet{padding:32px}
.sheet td,.sheet th{padding:10px 10px;font-size:13px}
.faq h3{font-size:18px;font-weight:500;line-height:1.4}
.faq p{margin-top:8px;color:var(--text)}

/* ---------- footer.closing (close-word + footer-statement) ---------- */
footer.closing{--h:1;background:var(--surface)}
footer.closing .wrap{height:100%;display:grid;grid-template-rows:1fr auto}
.close{display:flex;flex-direction:column}
.close__head{display:flex;justify-content:space-between;align-items:flex-start;gap:24px}
.close__head p{margin-top:12px;font-size:18px;color:var(--text)}
.close__name{margin-top:auto;font-size:clamp(96px,17vw,260px);line-height:.9;font-weight:500;letter-spacing:-.05em;color:var(--ink);padding-bottom:40px}
.dir{display:grid;grid-template-columns:repeat(12,minmax(0,1fr));gap:24px;padding-top:32px;border-top:1px solid var(--line);font-size:15px}
.dir>div{grid-column:span 4}
.dir>div:nth-child(3){grid-column:9 / span 4}
.dir b{display:block;font-weight:500;margin-bottom:6px}
.dir span{display:block;color:var(--text)}
.dir a{color:var(--ink)}
.legal{grid-column:1 / -1;font-size:13px;color:var(--muted);padding-top:16px;text-align:left!important}

/* ---------- 1024–1279 / tablet / phone ---------- */
@media (min-width:1024px) and (max-width:1279px){ .stu:nth-child(n+7){display:none} #app-view .map{height:120px} .open{gap:12px 16px;padding:16px 20px} #weights .who{display:none} .wl{margin-top:24px} .dr{padding:10px 14px;margin-bottom:8px} #weights .card{padding:24px} .cmp{padding:12px 0} .res .chart{padding:24px} .faq div:nth-child(4),.faq .q5{display:none} .faq div{padding:16px 0} }
@media (max-width:1279px){ .kpis{grid-template-columns:repeat(2,minmax(0,1fr))} .kpis .kpi:nth-child(n+3){display:none} }
@media (max-width:1023px){
  .hero .tbl .r5,.open .r-x{display:none}
  header.nav .wrap{grid-template-columns:1fr auto} .nav__words{display:none}
  .app{grid-template-columns:1fr;grid-template-rows:252px minmax(0,1fr)} .app__pane+.app__pane{border-left:0;border-top:1px solid var(--u-line)}
  .open{grid-template-columns:1fr} .map{height:150px}
  .split>.card,.split>.frame{grid-column:1 / -1}
  .res .chart,.res .side{grid-column:1 / -1} .res .side{flex-direction:row} .big{flex:1}
  .cards3{grid-template-columns:1fr} .panel{min-height:200px}
 
 
  .dir>div,.dir>div:nth-child(3){grid-column:1 / -1}
}
@media (max-width:639px){
  .ver{display:none}
  .panel .x2{display:none} #steps .panel{flex:none;min-height:0;-webkit-mask-image:none;mask-image:none;border-radius:12px;padding-bottom:20px} #steps .panel .ui{border-radius:10px;border-bottom:1px solid var(--u-border)} #steps .card{padding-bottom:24px}
  #results .head p{display:none} .sheet .t-hide{display:none} .hero .ui-bar>span:nth-of-type(1){display:none} .res .side{order:-1} .frame--sheet{padding:12px!important} .frame--sheet .ui-bar span:last-child{display:none}
  .faq .q5{display:block} .faq div{padding:14px 0} .faq h3{font-size:17px}
  #weights .head p,#weights .who,#weights .card p{display:none} #weights .wl{margin-top:16px} #weights .frame{padding:12px} #weights .dr{padding:10px 12px;margin-bottom:6px}
  #app-view .app{grid-template-rows:170px minmax(0,1fr)} #app-view .map{height:110px} #app-view .alts{display:none}
  .hero h1{line-height:1.15}
  .hero .btn-row{flex-direction:column;width:100%} .hero .btn{width:100%}
  .hero__frame{margin-top:32px}
  .tbl .m-hide{display:none} .hero .tbl td,.hero .tbl th{padding:9px 6px;font-size:13px} .hero .tbl .lg{display:none}
  .kpis{padding:12px 12px 0;gap:8px} .kpi{padding:8px 10px} .kpi b{font-size:18px}
  .card{padding:24px}
  .split{margin-top:32px} .cards3,.crew,.res,.app{margin-top:32px}
  .cmp{grid-template-columns:1fr 64px;gap:6px 12px} .cmp>b{grid-row:auto;grid-column:1 / -1}
  .res .chart{padding:20px} .legend{margin-bottom:8px} .cmp{padding:10px 0} .res .side{flex-direction:row;gap:12px} .big{padding:0;background:none;border:0} .big p{font-size:14px} .fn{margin-top:16px} #results .head p{font-size:16px}
  .crew{padding:16px 16px 28px}
  .close__head{flex-direction:column} .close__head .btn{width:100%}
  .close__name{font-size:88px;padding-bottom:24px}
}
</style>
</head>
<body id="top">
<!-- deal: kind software · palette cool-white-slate · type reddit+Pretendard · hero centred · nav-centre · footer-statement · close-word · rhythm dark-band · buttons tonal · mark initial -->
<header class="nav"><div class="wrap">
  <a class="skip" href="#app-view">본문으로 건너뛰기</a>
  <a class="brand" href="#top"><span class="mark mark--initial" aria-hidden="true">로</span>로켓단<span class="ver">__APP_VERSION__</span></a>
  <nav class="nav__words" aria-label="주요 구역"><a href="#app-view">화면</a><a href="#weights">비중</a><a href="#results">결과</a><a href="#steps">순서</a><a href="#faq">질문</a></nav>
  <a class="btn btn--primary nav__act" href="/app">배정 시작하기</a>
</div></header>

<main>
<section class="hero" style="--h:1;--hm:1"><div class="wrap"><div class="hero-text">
  <div class="hero__head">
    <div><p class="notice">가상 학생 221명 시험에서 평균 통학시간 <b class="num">56.7분 → 35분</b></p><h1>간호학과 실습 병원 배정 프로그램</h1></div>
    <div><p class="hero__lead">학생 명단을 올리면 통학시간과 1~3지망을 함께 계산해, 병원 정원 안에서 배정안을 추천합니다. 최종 결정은 교수자가 합니다.</p>
    <div class="btn-row"><a class="btn btn--primary" href="/app">배정 시작하기</a><a class="btn btn--secondary" href="#weights">비중 바꿔 보기</a></div></div>
  </div>
  <div class="hero__frame" data-visual>
    <div class="ui" role="img" aria-label="배정 결과 화면: 221명 배정, 평균 통학 35분, 지망 충족률 76%">
      <div class="ui-bar"><b>배정 결과</b><span>4학년 1학기 관리 실습 · 9개 병원</span><span class="sp"></span><div class="ui-tabs"><span class="on">전체 학생</span><span>병원별</span><span>미배정·사유</span></div></div>
      <div class="kpis">
        <div class="kpi"><small>배정 학생</small><b class="num">221명</b><i>미배정 0명</i></div>
        <div class="kpi"><small>평균 통학시간</small><b class="num">35분</b><i>무작위 대비 −21.7분</i></div>
        <div class="kpi"><small>지망 충족률</small><b class="num">76%</b><i>무작위 대비 +56.6%p</i></div>
        <div class="kpi"><small>평균 매칭점수</small><b class="num">75.6점</b><i>무작위 대비 +33.8점</i></div>
      </div>
      <div style="padding:12px 20px 0">
      <table class="tbl num"><thead><tr><th>배정 병원</th><th class="m-hide">학번</th><th class="m-hide">이름</th><th>지망</th><th class="r">통학<span class="lg">시간</span></th><th class="r m-hide">환승</th><th class="r"><span class="lg">종합</span>점수</th></tr></thead><tbody>
        <tr><td>순천향대학교부천병원</td><td class="m-hide">202○○○15</td><td class="m-hide">이○은</td><td><span class="chip chip--hit">1지망</span></td><td class="r">약 18분</td><td class="r m-hide">0회</td><td class="r sc">97.2</td></tr>
        <tr><td>고려대학교구로병원</td><td class="m-hide">202○○○07</td><td class="m-hide">김○수</td><td><span class="chip chip--hit">1지망</span></td><td class="r">약 24분</td><td class="r m-hide">1회</td><td class="r sc">91.6</td></tr>
        <tr><td>단원병원</td><td class="m-hide">202○○○31</td><td class="m-hide">박○준</td><td><span class="chip chip--two">2지망</span></td><td class="r">약 29분</td><td class="r m-hide">1회</td><td class="r sc">77.9</td></tr>
        <tr><td>아주대학교병원</td><td class="m-hide">202○○○42</td><td class="m-hide">최○린</td><td><span class="chip chip--miss">지망 외</span></td><td class="r">약 21분</td><td class="r m-hide">0회</td><td class="r sc">67.4</td></tr>
        <tr class="r5"><td>시화병원</td><td class="m-hide">202○○○58</td><td class="m-hide">정○우</td><td><span class="chip chip--two">3지망</span></td><td class="r">약 33분</td><td class="r m-hide">1회</td><td class="r sc">64.1</td></tr>
      </tbody></table></div>
    </div>
  </div>
</div></div></section>

<section id="app-view" style="--h:1"><div class="wrap">
  <div class="head"><h2 class="h2">학생마다 배정 이유를 보여 줍니다</h2><p>학생을 누르면 거주지와 병원 위치, 점수 계산 내역, 다른 병원이었을 때의 점수가 열립니다. 학생에게 설명할 때 이 화면을 그대로 쓰면 됩니다.</p></div>
  <div class="ui app" data-visual role="img" aria-label="학생 목록과 선택한 학생의 배정 근거 화면">
    <div class="app__pane app__pane--list">
      <div class="ui-bar"><b>학생 221명</b><span class="sp"></span><span>점수순</span></div>
      <div class="stu on"><b>김○수 · 202○○○07</b><span class="sc num">91.6</span><small>고려대학교구로병원 · 1지망 · 약 24분</small></div>
      <div class="stu"><b>이○은 · 202○○○15</b><span class="sc num">97.2</span><small>순천향대학교부천병원 · 1지망 · 약 18분</small></div>
      <div class="stu"><b>박○준 · 202○○○31</b><span class="sc num">77.9</span><small>단원병원 · 2지망 · 약 29분</small></div>
      <div class="stu"><b>최○린 · 202○○○42</b><span class="sc num">67.4</span><small>아주대학교병원 · 지망 외 · 약 21분</small></div>
      <div class="stu"><b>정○우 · 202○○○58</b><span class="sc num">64.1</span><small>시화병원 · 3지망 · 약 33분</small></div>
      <div class="stu"><b>한○영 · 202○○○63</b><span class="sc num">88.0</span><small>고려대학교안산병원 · 1지망 · 약 15분</small></div>
      <div class="stu"><b>윤○호 · 202○○○70</b><span class="sc num">72.5</span><small>서울특별시서남병원 · 2지망 · 약 38분</small></div>
      <div class="stu"><b>서○아 · 202○○○81</b><span class="sc num">83.3</span><small>한림대학교성심병원 · 1지망 · 약 27분</small></div>
      <div class="stu"><b>오○민 · 202○○○94</b><span class="sc num">69.8</span><small>단원병원 · 지망 외 · 약 19분</small></div>
    </div>
    <div class="app__pane app__pane--open">
      <div class="ui-bar"><b>김○수</b><span>202○○○07</span><span class="sp"></span><span class="chip chip--hit">1지망 배정</span></div>
      <div class="open">
        <div class="ttl">고려대학교구로병원<small>종합 매칭점수 91.6점</small></div>
        <div class="map" aria-hidden="true">
          <div class="road" style="left:-10px;right:-10px;top:96px;height:14px"></div>
          <div class="road" style="left:300px;top:-10px;bottom:-10px;width:14px"></div>
          <div class="road" style="left:-10px;width:260px;top:30px;height:8px"></div><div class="road" style="right:-10px;width:200px;top:146px;height:8px"></div><div class="park" style="left:150px;top:126px;width:120px;height:40px"></div><div class="park" style="right:60px;top:84px;width:90px;height:50px"></div>
          <svg viewBox="0 0 600 180" preserveAspectRatio="none"><path d="M90 128 C 200 120, 260 60, 470 46" fill="none" stroke="#0d7480" stroke-width="3" stroke-dasharray="7 6"/></svg>
          <div class="pin" style="left:60px;top:116px"><i style="background:#e0863a"></i>거주지(동 단위)</div>
          <div class="pin" style="right:24px;top:30px"><i style="background:#0d7480"></i>고려대학교구로병원</div>
        </div>
        <div class="box"><small>점수 계산</small>
          <div class="bar"><i style="width:61.6%;background:#0d7480"></i><i style="width:30%;background:#e0a13a"></i></div>
          <div class="row"><span>통학 점수 88 × 70%</span><b class="num">61.6</b></div>
          <div class="row"><span>지망 점수 100 × 30%</span><b class="num">30.0</b></div>
          <div class="row r-x"><span>약 24분 · 환승 1회 · 도보 6분</span></div>
        </div>
        <div class="box alts"><small>다른 병원이었다면</small>
          <div class="row"><span>고려대학교안산병원</span><b class="num">84.2</b></div>
          <div class="row"><span>단원병원</span><b class="num">71.0</b></div>
          <div class="row"><span>시화병원</span><b class="num">58.7</b></div>
          <div class="row"><span>아주대학교병원</span><b class="num">52.3</b></div>
        </div>
      </div>
    </div>
  </div>
</div></section>

<section id="weights" class="ground-surface" style="--h:1"><div class="wrap">
  <div class="head"><h2 class="h2">통학과 희망의 비중을 직접 정합니다</h2><p>슬라이더를 움직이면 같은 학생의 추천 병원 순서가 바로 바뀝니다. 실제 프로그램도 이 비중으로 전체 학생을 계산합니다.</p></div>
  <div class="split">
    <div class="card">
      <h3>예시 학생 한 명</h3>
      <p>1지망 고려대학교구로병원, 2지망 서울특별시서남병원을 적었고 집은 안산에 있습니다.</p>
      <div class="who"><img src="data:image/png;base64,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" alt="" width="232" height="384"><p class="muted">통학 15분이면 100점, 90분이 넘으면 0점으로 바꿔 계산합니다.</p></div>
      <div class="wl"><span id="dLblP">희망 병원 30%</span><span id="dLblD">통학 70%</span></div>
      <input type="range" id="dSlider" min="0" max="100" step="5" value="70" aria-label="통학과 희망 병원의 비중">
    </div>
    <div class="frame" id="dRows" data-visual aria-live="polite"></div>
  </div>
</div></section>

<section id="results" class="ground-inverse" style="--h:1"><div class="wrap">
  <div class="head"><h2 class="h2">무작위 배정과 비교한 결과</h2><p>가상 학생 221명과 4학년 1학기 관리 실습 병원 9곳으로, 정원만 지키며 무작위로 배정한 경우와 비교했습니다.<sup>1</sup></p></div>
  <div class="res">
    <div class="chart" data-visual data-ui>
      <div class="legend"><span><i style="background:#3d5560"></i>무작위 배정 100회 평균</span><span><i style="background:#5fc2c9"></i>로켓단 추천</span></div>
      <div class="cmp"><b>평균 통학시간</b><div class="ln" style="width:100%"></div><span class="num">56.7분</span><div class="ln us" style="width:61.7%"></div><span class="us num">35분</span></div>
      <div class="cmp"><b>지망 충족률</b><div class="ln" style="width:25.5%"></div><span class="num">19.4%</span><div class="ln us" style="width:100%"></div><span class="us num">76%</span></div>
      <div class="cmp"><b>평균 매칭점수</b><div class="ln" style="width:55.3%"></div><span class="num">41.8점</span><div class="ln us" style="width:100%"></div><span class="us num">75.6점</span></div>
    </div>
    <div class="side">
      <div class="big"><b class="num">−21.7분</b><p>학생 한 명당 하루 편도 통학이 이만큼 줄었습니다.</p></div>
      <div class="big"><b class="num">0명</b><p>221명 모두 병원 정원 안에서 배정됐습니다.</p></div>
    </div>
  </div>
  <p class="fn"><sup>1</sup> 시뮬레이션 결과이며 통학시간은 지도 기반 예상값입니다. 실제 학기 명단으로 다시 확인할 예정입니다.</p>
</div></section>

<section id="steps" style="--h:1;--ht:2;--hm:2"><div class="wrap">
  <div class="head"><h2 class="h2">세 단계로 배정합니다</h2><p>엑셀 명단 하나면 됩니다. 올린 파일은 서버에 저장하지 않고, 결과는 엑셀로 내려받습니다.</p></div>
  <div class="cards3">
    <div class="card"><span class="step">1단계</span><h3>교과목과 조건을 고릅니다</h3>
      <div class="panel" data-visual><div class="ui">
        <div class="f-row"><span>교과목</span><span class="f-sel">4학년 1학기 관리 실습</span></div>
        <div class="f-row"><span>실습 기간</span><span class="f-sel num">10일</span></div>
        <div class="f-row"><span>통학 : 희망</span><span class="f-sel num">70 : 30</span></div>
        <div class="f-row x2"><span>병원별 정원</span><span class="f-sel">9곳 입력됨</span></div>
        <div class="f-row x2"><span>성별 조건</span><span class="f-sel">병원별 적용</span></div>
        <div class="f-row x1 x2"><span>GPA 하한</span><span class="f-sel num">없음</span></div>
        <div class="f-row x1 x2"><span>과거 실습기관</span><span class="f-sel">같은 병원 제외</span></div>
        <div class="f-row x1 x2"><span>배정 방식</span><span class="f-sel">전체 병원 동시 배정</span></div>
      </div></div></div>
    <div class="card"><span class="step">2단계</span><h3>학생 명단을 올립니다</h3>
      <div class="panel" data-visual><div class="ui">
        <div class="drop">엑셀(.xlsx) 또는 CSV 파일을 끌어다 놓으세요</div>
        <div class="file"><span class="xl">XLS</span><div><b>4학년_명단.xlsx</b><small class="num">221명 · 학번, 이름, 성별, GPA, 주소, 1~3지망</small></div></div>
        <div class="f-row x2"><span>필수 열 확인</span><span class="ok">7개 모두 있음</span></div>
        <div class="f-row x2"><span>주소 → 지도 위치</span><span class="ok">221명 완료</span></div>
        <div class="f-row x1 x2"><span>자격 조건 확인</span><span class="ok">221명 통과</span></div>
        <div class="f-row x1 x2"><span>통학시간 계산</span><span class="ok">1,989건 완료</span></div>
      </div></div></div>
    <div class="card"><span class="step">3단계</span><h3>추천안을 보고 확정합니다</h3>
      <div class="panel" data-visual><div class="ui">
        <div class="dl"><span>엑셀로 내려받기</span></div>
        <table class="tbl num"><thead><tr><th>병원</th><th class="r">배정</th><th class="r">정원</th></tr></thead><tbody>
          <tr><td>고려대구로</td><td class="r">32</td><td class="r">32</td></tr>
          <tr><td>순천향부천</td><td class="r">28</td><td class="r">28</td></tr>
          <tr><td>단원병원</td><td class="r">24</td><td class="r">24</td></tr>
          <tr class="x2"><td>고려대안산</td><td class="r">30</td><td class="r">30</td></tr>
          <tr class="x1 x2"><td>아주대학교병원</td><td class="r">26</td><td class="r">26</td></tr>
          <tr class="x1 x2"><td>시화병원</td><td class="r">22</td><td class="r">22</td></tr>
        </tbody></table>
      </div></div></div>
  </div>
</div></section>

<section id="team" class="ground-surface" style="--h:1"><div class="wrap">
  <div class="head"><h2 class="h2">만든 사람들</h2><p>안산대학교 간호학과 3학년 강지우(팀장), 최영락, 정현민이 만들고 강민정 교수님(간호관리학)이 지도했습니다. 실습을 다니며 겪은 배정 문제에서 출발한 캡스톤디자인 작품입니다.</p></div>
  <div class="crew" data-visual><img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAmQAAAEyCAYAAABULszLAAAQ50lEQVR42u3dX4wV130H8B8uQlaECEXI2ZJ1wMuWrCtDKcGwoNK6BFOnsaI1CX+EakRbZCGrWhGHSG5FXJqiOlIItvYhoogHSiTXsROMIiehduI6JcGAXYTAsTcOrMFguo1WFKxViiICfavyMIM57szOzL2fz+NhPHf+HK6/OldfzrggIiK6enqvF3GeezreK/U6Xx6eVMh5hgYPj/PWAaAebvEIAAAEMgAAgQwAAIEMAEAgAwCgGuO0KWNMWpMdUyZmjg9fHI06tTK1LwEgrJABAAhkAAAIZAAAAhkAAAIZAEC7GO8RRKVtyrLbl3n3VdSemABEpXsoEy3xrwRYIQMAEMgAAAQyAAAEMgAAgQwAgNCybAmp7ciy97IEQhsutOGaNH86Y9gkuYHz0ZH0nJsy36yQAQAIZAAAAhkAAAIZAIBABgBAaFlCaM+F9lyD5sMtVy556Tdw7dbJjW7Dlf39oE1Z7HNLbV/W7fvNChkAQPjJEgBAIAMAQCADABDIAAAILUuImrSlJvqbcUOjV6Ml95KrW5vy4qWLmeNTJk9p9HPLa19Gi7b8irK89/ca8XxeOPyG92iFDADAT5YAAAhkAAACGQAAAhkAQGhZNs3Lw5Myx7vHn086T+fUSaVe5/mR95KOP3W108sNbcqxem557cuwh+mYWL++P+n4PXsGGvFemtLSLXtvyqa0F1PboEXdV9P3BrVCBgAgkAEACGQAAAhkAAACGQAAFRkXbbZHYarT9/U0+vnMPDBY6vnr1n4qe8/KeQv+sCW/CI4d/XGUucdl3eZP3jwpe8/KvBZkXtux7L0s69bKzNvjsm7zpOltPm4sby/LsuehFTIAgPCTJQCAQAYAgEAGACCQAQAQ9rIs0j0d7xVynr86frThzyFK3esz2mzPyqLaiHWT1x5Nvd+m73EZFbUpq5J6PU25r6jZHo5N2YPSewkrZAAA4SdLAAAEMgAAgQwAAIEMACC0LKPV96zsmDKxVvc1fHE0WnEv0brtcVk3IyPZe+FNndpRyflbtT0aDWkvlr1nZVX3BYQVMgAAgQwAAIEMAEAgAwBAIAMACC3LqNWelUW1Gs+PZH9u59Ri9nysqg2a93zynnPT97iEOqhbm5L29NL+3Y24zqV9G8IKGQAAAhkAgEAGAIBABgAgkAEAEFqWFSuqTRmJbcey251ES+9ZCdCUNmVYIQMAQCADABDIAAAQyAAABDIAAELLcuzktRejZq1MbUoAftMfzZ+TOb51S3+trnPrtgEvK6yQAQAIZAAACGQAAAIZAAACGQBAaFlGe+5NuWtnM9olD23sNxkKVPaekvasbA2vPf1Y5vj8NV9uyc9tVS8cfiPsNZmvqNbn0r4NlbyXsEIGABB+sgQAQCADABDIAAAQyAAAQsuyBeS1HYtqF16ZMKWQ86zr31rIefYObDXLP4Bj39qWOT7vc1syx++alv1X6fULV5OOz3tfqZ+bJ+96Uu+XD2bBrKmZ4w8/9vWk44++NVKrz62qrdnV03u9Ce99+NJo5njH5ImZ4xu2Pl3q9ezeuqaQ85R9nf9+/PXM8c6CnnNYIQMAQCADABDIAAAQyAAABDIAACLsZRl13YNSO7JZqnpf5knUcg9KotI25bPP7E06/8pV62p1v3ktyFl9aa3nt/ZvK7UdWdV1Dh0fDitkAAAIZAAAAhkAAAIZAIBABgBAaFnG2LQp89oip998JXP81z/7YanXWbfrCXtcjsl7yfvcsueJPS6j0DblvPUDSe/l8pF/KfW+yr6esve4zGtTzpv9u4Wcp27ty7z38uCiP4gyW5OprcyyrzOskAEAIJABAAhkAAAIZAAAAhkAAKFlGe/bjnxoY3+pnzvzzkWRsidX2Yq6ntS9O1P3LGs3rTpPQptyTN7LsT39LXk9ec85r/VZVJsyT1HnKVteS3HxXV2uM6yQAQAgkAEACGQAAAhkAAACGQAAYS/L+P+0BYm23IOSsXme2rhw80ZHrmT/D3vG/Mzxo6PZ51mbc568PSsjcY/Lqq5zad8PwgoZAAACGQCAQAYAgEAGACCQAQBQmHFNudCunt7rKcfbm69YqS25ocHD48wT86RV5klVe0q2qrw9K1Of/+kPzS9kb8pjJ3+eOb5y1brM8c4YTjr/d/Yszhz/zPpDmeMv7d9tktzA0r4NmePDl7JrnB2TJyad/3x0VPJ9ZYUMACD8ZAkAIJABACCQAQAIZAAAhL0s4wO0sVLbUnmuvHQ8c/zWpXMLuf7DW3YmHd+7bWM04b6qasmZJ+ZJnefJsae/lzk+b82fFXL9r2/fl3T8XZtXRBPuK9XMX76W05pMO09emzISW3h57csFfS9ESvvv2vGRzPFbbpuSfUHTGr62cuFa9nP4xcVSPzbvPYYVMgCA8JMlAAACGQCAQAYAgEAGABBalm2oqDZZno7u7szxGeuXNfq+zBPzhBjz1mHePJnaN6fR9xUlty9T99Asqo2bupfisq2PRpl7XL71lQOlPv9Zj96X9h/ktESXPfxolPmc69YKt0IGABB+sgQAEMgAABDIAAAEMgAAQsuySJfeOVOr65m8tLsR19luzBNuxsiF8/W6oAVTmnGdBXnq9ey1gxWrP5t4poFSW3WpeyMeHdhQyffVbWt7sv/gxNW0E80ZX+p17hvYkvbXon93WCEDAEAgAwAQyAAAEMgAAAQyAABCyzJu3LZ4a/82b/0GZvVtaav7NU/Mk5ux/G/3ZY4f29NvMtxA6t6RTZHayszb+7Kq75/cFuScYiLB5I/NqNX3T1V7U4YVMgCA8JMlAAACGQCAQAYAgEAGANAItWse5LVRmiJ1b7IoqBUYLdpOMk/MkzrPk2ef2Zs5vnLVukLO/8I/rogq2qNVPYfNmzdFMXtWZlt895JSW5O+f3z/hBUyAIDwkyUAAAIZAIBABgCAQAYAEPayLMnOfc8nHT/n9g8X8rl7D57LHF+35Pbs/+CVfy71ORx69WAx15noxLnLScdvXHG/eWKetO08yWsLjk5bkPT8R198olbzZOKFo5njH73jjqTz582TU4d+EE1u959+85Va/X/zyMmz2X8wXO5zfuqZpzPHF86eXqvnM/PORaXOh9S2phUyAIDwkyUAgEAGAIBABgAgkAEAEM1uWTZ9b688ee2nvBbSI10zSr2eHSW35MJepeaJeVKbefLQtM5Sr2dX4jzZ982jhXxuXpuye/GySt5LU1qBRTkzbXnmeOe0qUnnOX9hJPsPrl1rxHNIbcXmtVbXrloTVsgAAMJPlgAACGQAAAIZAAACGQBAaFlG4l5pUfIeeVFRW2pHzu0+0jWUdP4dQ12NbsnlvffFdy8xT2o4T/Z989ulPoe89lxR7z11XjVlnuzKeTwPTTmcdP5dF3sb/X1CsfLalLde3J45fmXK5qTznDn/Cw85rJABAAhkAAAIZAAAAhkAAAIZAEBoWQKFeen5PZnjc3OOP/FG4gdMvr/cG/jepuzrbNH3tWL1ZyOlzZp3fFWKus6881S1ZyU3ltemLMrfPDtY6vkfX9kTVsgAABDIAAAEMgAABDIAAIEMAIBk47t6eq97DFHZ3oWp56E958lcj8w8ifff49L3Cb9p4ezpmeNHTp7NHJ/ReVvS+fP2rMz73P2Dg15KWCEDABDIAAAQyAAABDIAAAQyAIBozF6WH5pxTyEfsGzll0q9gR07NkeT21KFnf+J/8gc3/v5T5jl0fxWXd7elPd+4RuFfO69pd/Z6kLO8uLXHjRJatiaTN378sS5y15iG2r6XpOpUv8VCytkAADhJ0sAAIEMAACBDABAIAMAIOrVsqSe9h48V+rxRbW33n37bS8L4uZah6ktxSi5HZl6/Oi0BaVe5+K7l5g8/J8lf/JAra5nWkdH0vEXhofDChkAQPjJEgAAgQwAQCADAEAgAwCI+rcsp83o8XRqaO6MSUnjtKevfr8ZLdfvjO7KHD+48vHQviyufZnapuxevKwRz21o8PC4KHFvQZqlf2BPra5noH99WCEDAAg/WQIAIJABAAhkAAAIZAAA0dyW5YUzg7W60Ic/N9fbgpv0xU/d0YzrjMe9rCiufZl6nhPnLnvILezIybO1+tyFs6fX6+/XhNcKOc++X80PK2QAAOEnSwAABDIAAIEMAACBDAAgWq9lWTdf/9bxzPGHc46fk7jXG9TB1Hd/lDm+9P71meO//tkPM8d3DHVFE1qf//nAA5njv/PccyZDpLcm4WbktR3/67//p5DzVNXuDCtkAAAIZAAAAhkAAAIZAIBABgBAtHLLMk/34mVR5l5vdTPn9g834jp/f8I7/naFvSzfz21f+etIaY/mtU1j9xYvvYW/Txgb5y+MJB3/kd++vdH3W9QelGGFDAAg/GQJAIBABgAgkAEAIJABAERLtix/eeblQj7g0KsHCznPiXOXo8zWZF4r055xH8ysvrTW2+K7lxTyuTv3Pe/hR8RXv/92ra7nz3c9EiktS2gHf/oPRzLH//VLCyu5nsd2H0o6/rtPrPYSwwoZAIBABgCAQAYAIJABACCQAQA03vihwcPjijhRV0/v9azxqe/+qKBLnVvqg9CmrKfNmzdljm/f/mRL3u/IR/84c/zEP/1F5vi9X/hG5vgXP16zG/vUc5GyZ2VRbv3ps9l/MGm5v1yMmSMnz5Z6/MLZ0wu5zuEzg15WWCEDABDIAAAQyAAABDIAAAQyAICwlyWUKK/Vm9fSpbX91sc/mXT8i1970EOjcPMnnIrsFvCpUufzl/s6KrnfTzzw95njlwYPJJ3n05/vyRx/7C97K7mvgf71YYUMAACBDABAIAMAQCADABDIAACI2rUsZ/VtKeQ8O/c9n3T8iXOXvV3MkzY3b/1AzvfJcvMEGurgvz1XyHlm3rkorJABACCQAQAIZAAACGQAAAIZAADFtSzz9ihMlben4cYV90eZrcyiHD/zXvb1PLK2ra4/9f0WpVXnyc4N3WFvR/PE90nU+l8DOP3mJ/3litbdczmskAEAhJ8sAQAQyAAABDIAAAQyAIDQsoz6tRvKbu0xNjZv3pQ5vn37k4XMn6bPk427T4X2nO+T0G6LOv9rAHkWzp7u5RJWyAAABDIAAAQyAACBDAAAgQwAILQsaSMTxrfmXmN4XwBhhQwAQCADAEAgAwAQyAAAEMgAAELLEiLy25Tz5i/IHD/22tHM8YEnnyxkLzltPqAqR06e9RA8n7BCBgAgkAEAIJABAAhkAAAIZAAAoWVJ2Jsyr01Z1Of+6qp3AtTL2lVrko7vuW9T5nh318cyx08NvdOI5/DUyxdKvf7hM4OZ45cGD4QVMgAABDIAAIEMAACBDABAIAMAILQsAaBEqXvm5u29Sz3fV1ghAwAIP1kCACCQAQAIZAAACGQAAKFlWUNaKtXK2ztywvhqPtc8aZa891JV+8k8MU/qrCl7Vrbq9YcVMgAAgQwAAIEMAEAgAwBAIAMAaC/jmt6+Of3mK424/pl3Lko6vun3ldq+TG1T5rWuzJNm3VfdWpbmiXkylu3dnvs2VXKdw2cGM8cvDR6Idrr+us0rK2QAAAIZAIBABgCAQAYAIJABABD2siRaf+/LVO22hx3QeprTck1rKX73idU1u4O/K6SVHFbIAADCT5YAAAhkAAACGQAAAhkAQGhZNsixkz9vyRfT9PuqWzvSPME8oSo/efWn7ouwQgYAIJABACCQAQAIZAAACGQAADXVmH0Cu3p6r3tdoWVpnpgn5ol54v9TtOC8skIGACCQAQAIZAAACGQAAAIZAAAV+V+lQDsIKq3QnAAAAABJRU5ErkJggg==" width="612" height="306" alt="팀 로켓단 캐릭터: 왼쪽부터 강지우, 최영락, 정현민, 강민정 교수님"></div>
  </div>
</div></section>

<section id="faq" style="--h:1;--hm:2"><div class="wrap">
  <div class="head"><h2 class="h2">자주 묻는 질문</h2><p>처음 쓰는 교수자와 조교가 가장 많이 묻는 내용입니다.</p></div>
  <div class="split split--faq">
    <div class="faq">
      <div><h3>학생 명단은 서버에 저장되나요?</h3><p>저장하지 않습니다. 메모리에서 처리한 뒤 사라지고, 결과는 엑셀로 내려받아 보관합니다.</p></div>
      <div><h3>통학시간은 어떻게 계산하나요?</h3><p>지도 위치로 예상 소요시간을 구하고, TMAP 키를 넣으면 평일 오전 6시 출발 대중교통 경로를 씁니다.</p></div>
      <div><h3>배정 결과를 그대로 확정하나요?</h3><p>아닙니다. 프로그램은 추천안과 근거를 보여 주고, 최종 결정은 교수자가 합니다.</p></div>
      <div><h3>지도에 나오는 거주지는 실제 집인가요?</h3><p>아닙니다. 주소를 동 단위까지만 써서 동네 기준의 대략적인 위치입니다.</p></div>
      <div class="q5"><h3>미배정 학생이 나오면 어떻게 하나요?</h3><p>미배정·사유 탭에서 이유를 확인하고, 정원이나 조건을 바꿔 다시 계산하면 됩니다.</p></div>
    </div>
    <div class="frame frame--sheet" data-visual>
      <div class="ui" role="img" aria-label="명단 양식: 학번, 이름, 성별, GPA, 출생연도, 주소, 1~3지망 열">
        <div class="ui-bar"><span class="xl">XLS</span><b>명단_양식.xlsx</b><span class="sp"></span><span>필수 열 7개</span></div>
        <table class="tbl sheet num"><thead><tr><th>학번</th><th>이름</th><th class="t-hide">성별</th><th class="r t-hide">GPA</th><th class="r t-hide">출생</th><th class="t-hide">주소(동)</th><th>1지망</th></tr></thead><tbody>
          <tr><td>202○○○07</td><td>김○수</td><td class="t-hide">남</td><td class="r t-hide">3.84</td><td class="r t-hide">2004</td><td class="t-hide">안산 고잔동</td><td>고려대구로</td></tr>
          <tr><td>202○○○15</td><td>이○은</td><td class="t-hide">여</td><td class="r t-hide">4.12</td><td class="r t-hide">2005</td><td class="t-hide">부천 중동</td><td>순천향부천</td></tr>
          <tr><td>202○○○31</td><td>박○준</td><td class="t-hide">남</td><td class="r t-hide">3.47</td><td class="r t-hide">2003</td><td class="t-hide">안산 와동</td><td>아주대</td></tr>
          <tr><td>202○○○42</td><td>최○린</td><td class="t-hide">여</td><td class="r t-hide">3.95</td><td class="r t-hide">2005</td><td class="t-hide">수원 매탄동</td><td>아주대</td></tr>
          <tr><td>202○○○58</td><td>정○우</td><td class="t-hide">남</td><td class="r t-hide">3.21</td><td class="r t-hide">2004</td><td class="t-hide">시흥 정왕동</td><td>시화병원</td></tr>
          <tr><td>202○○○63</td><td>한○영</td><td class="t-hide">여</td><td class="r t-hide">3.68</td><td class="r t-hide">2005</td><td class="t-hide">안산 사동</td><td>고려대안산</td></tr>
        </tbody></table>
      </div>
    </div>
  </div>
</div></section>
</main>

<footer class="closing" id="contact"><div class="wrap">
  <div class="close">
    <div class="close__head"><div><h2 class="h2">명단을 올리면 배정안이 나옵니다</h2><p>설치 없이 웹에서 바로 씁니다. 샘플 명단으로 먼저 해 볼 수도 있습니다.</p></div><a class="btn btn--primary" href="/app">배정 시작하기</a></div>
    <div class="close__name" aria-hidden="true">로켓단</div>
  </div>
  <div class="dir">
    <div><b>로켓단</b><span>AI기반 간호실습 배치 의사결정 프로그램</span><span>안산대학교 간호학과 캡스톤디자인</span></div>
    <div><b>팀</b><span>강지우(팀장), 최영락, 정현민</span><span>지도교수 강민정</span></div>
    <div><b>바로가기</b><span><a class="link" href="/app">배정 시스템</a></span>__GUIDE_FOOT__</div>
    <div class="legal">© 2026 로켓단</div>
  </div>
</div></footer>

<script>
(function () {
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
      return '<div class="dr' + (i === 0 ? ' top' : '') + '"><span class="nm">' + esc(x.n) + (i === 0 ? '<em>추천</em>' : '') + '</span><span class="sc num">' + x.s.toFixed(1) + '점</span><span class="sub">' + esc(x.g) + ', 통학 약 ' + x.t + '분</span><div class="tr"><i style="width:' + Math.max(2, x.s) + '%"></i></div></div>';
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
        uvicorn.run(app, host="0.0.0.0", port=8000)
