from fastapi import FastAPI, HTTPException, UploadFile, File, Form
from fastapi.responses import HTMLResponse
from pydantic import BaseModel, Field
from typing import List, Optional, Tuple
import pandas as pd
import numpy as np
import io
import uvicorn

from sklearn.ensemble import RandomForestClassifier
from scipy.optimize import linear_sum_assignment

app = FastAPI(title="간호학과 스마트 실습지 최적 배정 시스템 (v2026.09.10-6.1)")

SECRET_PASSWORD = "ansan king"

MASTER_HOSPITAL_DB = {
    "3학년": {
        "1학기": {
            "성인1": [
                "가톨릭대학교 부천성모병원", "가톨릭대학교 성빈센트병원", "고려대학교 안산병원", 
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
            "성
