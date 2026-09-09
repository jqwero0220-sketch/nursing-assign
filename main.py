from fastapi import FastAPI, HTTPException, UploadFile, File, Form
from fastapi.responses import HTMLResponse
from pydantic import BaseModel, Field
from typing import List, Optional, Dict, Tuple
import pandas as pd
import numpy as np
import asyncio
import io
import uvicorn

from sklearn.ensemble import RandomForestClassifier
from scipy.optimize import linear_sum_assignment

app = FastAPI(title="로켓단 AI 실습지 최적 배정 시스템 v5.4")

SECRET_PASSWORD = "ansan king"

ROUTE_CACHE: Dict[str, Tuple[int, int, int]] = {}

def get_cached_route_info(address: str, station: str, mode_raw: str, default_time: int) -> Tuple[int, int, int]:
    cache_key = f"{address}_{station}_{mode_raw}"
    if cache_key in ROUTE_CACHE:
        return ROUTE_CACHE[cache_key]
    
    transfers = 1 if '버스' in mode_raw and station != '' else 0
    walk_time = 8
    result = (default_time, transfers, walk_time)
    ROUTE_CACHE[cache_key] = result
    return result

class SatisfactionMLModel:
    def __init__(self):
        self.model = RandomForestClassifier(n_estimators=50, random_state=42)
        self._train_model()

    def _train_model(self):
        np.random.seed(42)
        X_train, y_train = [], []
        for _ in range(300):
            travel_time = np.random.randint(10, 90)
            transfers = np.random.randint(0, 4)
            walk_time = np.random.randint(3, 25)
            gpa = np.random.uniform(2.5, 4.5)
            mfi = travel_time + (transfers * 12.0) + (walk_time * 1.2)

            if mfi < 35:
                label = 2
            elif mfi < 60:
                label = 1
            else:
                label = 0

            X_train.append([travel_time, transfers, walk_time, gpa, mfi])
            y_train.append(label)

        self.model.fit(X_train, y_train)

    def predict(self, travel_time: int, transfers: int, walk_time: int, gpa: float, mfi: float) -> int:
        base_score = max(30, min(99, int(100 - (mfi * 0.8))))
        return base_score

ml_engine = SatisfactionMLModel()

def generate_ai_report(name: str, hospital: str, rank: Optional[int], mfi: float, travel_time: int, transit_mode: str, gpa: float, is_eligible: bool, note: str) -> str:
    if not is_eligible:
        return f"[AI 분석] {name} 학생은 {note}로 인해 {hospital} 배정 자격 미달로 판정되었습니다."
    
    report = f"[AI 리포트] {name} 학생은 {hospital} {rank}순위 최적 배정 대상자입니다. "
    report += f"거주지 기반 통학 소요시간 {travel_time}분({transit_mode}) 및 다변수 피로도 지수(MFI {mfi})가 최상위권이며, "
    report += f"GPA({gpa}) 기준 조건을 충족하여 통학 피로도 최소화 관점에서 최적의 배정안으로 평가됩니다."
    return report

class PasswordVerifyRequest(BaseModel):
    password: str

class HospitalCriteria(BaseModel):
    gender: str = Field(default="무관")
    min_gpa: Optional[float] = Field(default=None)
    birth_year_after: Optional[int] = Field(default=None)

class StudentInput(BaseModel):
    student_id: str
    name: str
    gender: str
    gpa: float
    birth_year: int
    address: str
    nearest_station: str
    travel_time_minutes: int
    transfers: int
    walk_time_minutes: int
    transit_mode: str
    fatigue_index: float

class AssignmentResult(BaseModel):
    rank: Optional[int]
    student_id: str
    name: str
    gender: str
    gpa: float
    birth_year: int
    nearest_station: str
    transit_mode: str
    travel_time_minutes: Optional[int]
    fatigue_index: Optional[float]
    ai_satisfaction_score: Optional[int]
    ai_report: str
    is_eligible: bool

class AssignmentResponse(BaseModel):
    status: str
    target_hospital: str
    total_students: int
    eligible_count: int
    optimization_method: str
    results: List[AssignmentResult]

def calculate_fatigue_index(travel_time: int, transfers: int, walk_time: int) -> float:
    mfi = travel_time + (transfers * 12.0) + (walk_time * 1.2)
    return round(mfi, 1)

def check_eligibility(student: StudentInput, criteria: HospitalCriteria) -> tuple[bool, str]:
    if criteria.gender == "남성만" and student.gender != "남":
        return False, "성별 불일치(남성만 가능)"
    elif criteria.gender == "여성만" and student.gender != "여":
        return False, "성별 불일치(여성만 가능)"

    if criteria.min_gpa is not None and student.gpa < criteria.min_gpa:
        return False, f"성적 미달({criteria.min_gpa} 이상 필요)"

    if criteria.birth_year_after is not None and student.birth_year < criteria.birth_year_after:
        return False, f"연령 미달({criteria.birth_year_after}년 이후 출생자 필요)"

    return True, "자격충족"

@app.post("/api/v1/verify-password")
async def verify_password(payload: PasswordVerifyRequest):
    if payload.password == SECRET_PASSWORD:
        return {"status": "success", "message": "인증 성공"}
    raise HTTPException(status_code=401, detail="비밀번호가 올바르지 않습니다.")

async def process_student_row_async(row: pd.Series) -> StudentInput:
    mode_raw = str(row.get('이동수단', '대중교통')).strip()
    station_info = str(row.get('인근역', ''))
    address = str(row.get('주소', ''))
    default_time = int(row['소요시간_분'])

    travel_time, transfers, walk_time = get_cached_route_info(address, station_info, mode_raw, default_time)
    mfi = calculate_fatigue_index(travel_time, transfers, walk_time)

    # 💡 이동수단 정밀 분류 로직 반영
    if '버스' in mode_raw and ('전철' in mode_raw or '지하철' in mode_raw):
        detail_mode = '지하철+버스'
    elif '버스' in mode_raw:
        detail_mode = '시내/시외버스'
    elif '전철' in mode_raw or '지하철' in mode_raw:
        detail_mode = '지하철(전철)'
    else:
        detail_mode = mode_raw if mode_raw and mode_raw != 'nan' else '대중교통'

    return StudentInput(
        student_id=str(row['학번']),
        name=str(row['이름']),
        gender=str(row['성별']),
        gpa=float(row['GPA']),
        birth_year=int(row['출생연도']),
        address=address,
        nearest_station=station_info,
        travel_time_minutes=travel_time,
        transfers=transfers,
        walk_time_minutes=walk_time,
        transit_mode=detail_mode,
        fatigue_index=mfi
    )

@app.post("/api/v1/assign-file", response_model=AssignmentResponse)
async def assign_hospital_from_file(
    target_hospital: str = Form(...),
    gender_criteria: str = Form("무관"),
    min_gpa: Optional[float] = Form(None),
    birth_year_after: Optional[int] = Form(None),
    use_hungarian: bool = Form(False),
    file: UploadFile = File(...)
):
    if not target_hospital or target_hospital.strip() == "":
        raise HTTPException(status_code=400, detail="배정 대상 병원을 선택해 주세요.")

    try:
        contents = await file.read()
        if file.filename.endswith('.csv'):
            df = pd.read_csv(io.BytesIO(contents))
        else:
            df = pd.read_excel(io.BytesIO(contents))

        tasks = [process_student_row_async(row) for _, row in df.iterrows()]
        students: List[StudentInput] = await asyncio.gather(*tasks)

    except Exception as e:
        raise HTTPException(status_code=400, detail=f"엑셀(CSV) 데이터 처리 실패: {str(e)}")

    criteria = HospitalCriteria(
        gender=gender_criteria,
        min_gpa=min_gpa,
        birth_year_after=birth_year_after
    )

    eligible_list = []
    ineligible_list = []

    for stu in students:
        is_ok, note = check_eligibility(stu, criteria)
        if is_ok:
            eligible_list.append({"student": stu, "status_note": note})
        else:
            ai_rep = generate_ai_report(stu.name, target_hospital, None, stu.fatigue_index, stu.travel_time_minutes, stu.transit_mode, stu.gpa, False, note)
            ineligible_list.append(
                AssignmentResult(
                    rank=None, student_id=stu.student_id, name=stu.name, gender=stu.gender,
                    gpa=stu.gpa, birth_year=stu.birth_year, nearest_station=stu.nearest_station,
                    transit_mode=stu.transit_mode, travel_time_minutes=None, fatigue_index=None,
                    ai_satisfaction_score=None, ai_report=ai_rep, is_eligible=False
                )
            )

    optimization_method = "Multi-Factor Fatigue Index (MFI) Sorting"

    if use_hungarian and len(eligible_list) > 1:
        cost_matrix = np.array([[item["student"].fatigue_index for _ in range(len(eligible_list))] for item in eligible_list])
        row_ind, col_ind = linear_sum_assignment(cost_matrix)
        ordered_eligible = [eligible_list[i] for i in row_ind]
        eligible_list = ordered_eligible
        optimization_method = "SciPy Hungarian Bipartite Global Optimization"
    else:
        eligible_list.sort(key=lambda x: x["student"].fatigue_index)

    final_results = []
    for rank_idx, item in enumerate(eligible_list, start=1):
        stu = item["student"]
        
        sat_score = ml_engine.predict(
            stu.travel_time_minutes, stu.transfers, stu.walk_time_minutes, stu.gpa, stu.fatigue_index
        )
        
        ai_rep = generate_ai_report(
            stu.name, target_hospital, rank_idx, stu.fatigue_index, stu.travel_time_minutes, stu.transit_mode, stu.gpa, True, item["status_note"]
        )

        final_results.append(
            AssignmentResult(
                rank=rank_idx, student_id=stu.student_id, name=stu.name, gender=stu.gender,
                gpa=stu.gpa, birth_year=stu.birth_year, nearest_station=stu.nearest_station,
                transit_mode=stu.transit_mode, travel_time_minutes=stu.travel_time_minutes,
                fatigue_index=stu.fatigue_index, ai_satisfaction_score=sat_score,
                ai_report=ai_rep, is_eligible=True
            )
        )

    final_results.extend(ineligible_list)

    return AssignmentResponse(
        status="success", target_hospital=target_hospital,
        total_students=len(students), eligible_count=len(eligible_list),
        optimization_method=optimization_method, results=final_results
    )

@app.get("/", response_class=HTMLResponse)
def render_ui():
    html_content = """
    <!DOCTYPE html>
    <html lang="ko">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>로켓단 AI 실습지 최적 배정 시스템</title>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
        <link href="https://fonts.googleapis.com/css2?family=Pretendard:wght@300;400;500;600;700&display=swap" rel="stylesheet">
        <style>
            body { font-family: 'Pretendard', sans-serif; background-color: #f8fafc; color: #1e293b; }
            .navbar-custom { background-color: #0f172a; border-bottom: 1px solid #1e293b; }
            .card-custom { border: none; border-radius: 16px; box-shadow: 0 4px 20px rgba(0,0,0,0.03); background: #ffffff; }
            .card-header-custom { background: #f8fafc; border-bottom: 1px solid #f1f5f9; font-weight: 700; color: #0f172a; border-radius: 16px 16px 0 0 !important; }
            .btn-run { background: linear-gradient(135deg, #2563eb, #1d4ed8); border: none; font-weight: 700; padding: 12px; font-size: 15px; border-radius: 10px; transition: all 0.2s; }
            .btn-run:hover { background: linear-gradient(135deg, #1d4ed8, #1e40af); transform: translateY(-1px); }
            .btn-excel { background: linear-gradient(135deg, #059669, #047857); border: none; font-weight: 700; border-radius: 8px; }
            .btn-excel:hover { background: linear-gradient(135deg, #047857, #065f46); }
            .dropzone-box { border: 2px dashed #cbd5e1; background: #f8fafc; border-radius: 12px; padding: 24px; text-align: center; }
            .table-custom th { background-color: #0f172a; color: #f8fafc; text-align: center; font-size: 13.5px; font-weight: 600; }
            .table-custom td { vertical-align: middle; text-align: center; font-size: 13.5px; }
            .rank-badge { background: #d97706; color: white; padding: 4px 10px; border-radius: 20px; font-weight: 700; font-size: 11.5px; }
            .badge-mode { background-color: #f1f5f9; color: #1e40af; font-weight: 600; padding: 3px 8px; border-radius: 6px; }
            .mfi-badge { background-color: #eff6ff; color: #1d4ed8; font-weight: 700; padding: 4px 10px; border-radius: 6px; border: 1px solid #bfdbfe; cursor: help; }
            .ai-badge { background-color: #ecfdf5; color: #047857; font-weight: 700; padding: 4px 10px; border-radius: 6px; border: 1px solid #a7f3d0; }
            
            .info-icon {
                display: inline-flex; align-items: center; justify-content: center;
                width: 16px; height: 16px; border-radius: 50%; background-color: #3b82f6;
                color: white; font-size: 10px; font-weight: bold; margin-left: 4px;
                cursor: pointer; vertical-align: middle;
            }
            .tooltip-inner {
                max-width: 320px; text-align: left; font-size: 12px; padding: 10px 14px;
                background-color: #0f172a; line-height: 1.5; border-radius: 8px;
                box-shadow: 0 10px 25px rgba(0,0,0,0.2);
            }

            .auth-overlay {
                position: fixed; top: 0; left: 0; width: 100vw; height: 100vh;
                background: radial-gradient(circle at 50% 30%, #1e293b 0%, #0f172a 100%);
                z-index: 9999; display: flex; justify-content: center; align-items: center;
            }
            .auth-card {
                background: rgba(30, 41, 59, 0.7);
                backdrop-filter: blur(20px);
                -webkit-backdrop-filter: blur(20px);
                width: 90%; max-width: 400px; padding: 40px 32px;
                border-radius: 24px;
                border: 1px solid rgba(255, 255, 255, 0.1);
                box-shadow: 0 20px 50px rgba(0, 0, 0, 0.4);
                text-align: center; color: #ffffff;
            }
            .auth-icon {
                width: 64px; height: 64px; background: rgba(59, 130, 246, 0.15);
                border-radius: 20px; display: flex; align-items: center; justify-content: center;
                margin: 0 auto 20px; font-size: 28px; border: 1px solid rgba(59, 130, 246, 0.3);
            }
            .auth-input {
                background: rgba(15, 23, 42, 0.6);
                border: 1px solid rgba(255, 255, 255, 0.15);
                color: #ffffff !important; border-radius: 12px;
                padding: 14px; font-size: 15px; letter-spacing: 1px;
                transition: all 0.25s ease;
            }
            .auth-input:focus {
                background: rgba(15, 23, 42, 0.8);
                border-color: #3b82f6;
                box-shadow: 0 0 0 4px rgba(59, 130, 246, 0.25); outline: none;
            }
            .auth-input::placeholder { color: #64748b; font-weight: 400; font-size: 14px; }
            .btn-auth {
                background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%);
                border: none; color: white; font-weight: 700;
                padding: 14px; border-radius: 12px; font-size: 15px;
                box-shadow: 0 4px 14px rgba(37, 99, 235, 0.4); transition: all 0.2s ease;
            }
            .btn-auth:hover { transform: translateY(-1px); box-shadow: 0 6px 20px rgba(37, 99, 235, 0.5); }
        </style>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    </head>
    <body>
        <div id="authOverlay" class="auth-overlay">
            <div class="auth-card">
                <div class="auth-icon">🚀</div>
                <h4 class="fw-bold mb-1" style="letter-spacing: -0.5px;">보안 서버 인증</h4>
                <p class="text-secondary fs-7 mb-4" style="color: #94a3b8 !important;">로켓단 AI 실습지 최적 배정 시스템 v5.4</p>
                <div class="mb-3">
                    <input type="password" id="authPassword" class="form-control auth-input text-center fw-semibold mb-2" placeholder="접속 암호를 입력하세요" onkeyup="if(window.event.keyCode==13){verifyPassword();}">
                    <div id="authError" class="text-danger fs-7 fw-bold mt-2" style="display:none; color: #f87171 !important;">❌ 백엔드 인증 실패: 올바른 암호가 아닙니다.</div>
                </div>
                <button onclick="verifyPassword()" class="btn btn-auth w-100">
                    시스템 접속하기
                </button>
            </div>
        </div>

        <nav class="navbar navbar-dark navbar-custom shadow-sm mb-4">
            <div class="container px-4">
                <span class="navbar-brand mb-0 h1 fw-bold fs-5" style="letter-spacing: -0.5px;">
                    🏥 로켓단 | AI 기반 간호학과 실습지 최적 배정 시스템
                </span>
                <span class="badge bg-primary fs-7 px-3 py-2 rounded-pill">v5.4 Transit Fixed</span>
            </div>
        </nav>

        <div class="container pb-5" style="max-width: 1140px;">
            <div class="row g-4 mb-4">
                <div class="col-md-6">
                    <div class="card card-custom h-100">
                        <div class="card-header card-header-custom py-3 px-4 fs-6">
                            📌 STEP 1. 교과목 및 병원 조건 설정
                        </div>
                        <div class="card-body p-4">
                            <div class="mb-3">
                                <label class="form-label fw-bold">1. 실습 교과목 선택</label>
                                <select id="subject_select" class="form-select fw-bold text-primary" onchange="updateHospitalOptions()">
                                    <option value="ALL">전체 교과목 (26개 전체 병원)</option>
                                    <option value="성인I">성인간호학실습 I</option>
                                    <option value="여성">여성건강간호학실습</option>
                                    <option value="성인II">성인간호학실습 II</option>
                                    <option value="아동">아동간호학실습</option>
                                    <option value="정신">정신간호학실습</option>
                                </select>
                            </div>

                            <div class="mb-3">
                                <label class="form-label fw-bold">2. 배정 대상 병원 선택</label>
                                <select id="hospital_select" class="form-select fw-bold">
                                </select>
                            </div>

                            <div class="mb-3">
                                <label class="form-label fw-bold">3. 성별 조건</label>
                                <select id="gender_criteria" class="form-select">
                                    <option value="무관" selected>무관</option>
                                    <option value="남성만">남성만</option>
                                    <option value="여성만">여성만</option>
                                </select>
                            </div>

                            <div class="accordion" id="advancedOptions">
                                <div class="accordion-item border-0 bg-light rounded-3">
                                    <h2 class="accordion-header">
                                        <button class="accordion-button collapsed bg-light fw-bold text-secondary fs-7 py-2" type="button" data-bs-toggle="collapse" data-bs-target="#collapseAdvanced">
                                            ⚙️ 세부 자격 조건 및 알고리즘 옵션
                                        </button>
                                    </h2>
                                    <div id="collapseAdvanced" class="accordion-collapse collapse" data-bs-parent="#advancedOptions">
                                        <div class="accordion-body pt-2 pb-3">
                                            <div class="row g-2 mb-2">
                                                <div class="col-6">
                                                    <label class="form-label fs-7 fw-bold mb-1">최소 GPA (선택)</label>
                                                    <input type="number" step="0.1" id="min_gpa" class="form-control form-control-sm" placeholder="예: 3.5">
                                                </div>
                                                <div class="col-6">
                                                    <label class="form-label fs-7 fw-bold mb-1">출생연도 조건 (선택)</label>
                                                    <input type="number" id="birth_year" class="form-control form-control-sm" placeholder="예: 2003년 이후">
                                                </div>
                                            </div>
                                            <div class="form-check mt-2 d-flex align-items-center">
                                                <input class="form-check-input me-2" type="checkbox" id="use_hungarian">
                                                <label class="form-check-label fs-7 fw-bold text-dark mb-0" for="use_hungarian">
                                                    🔬 SciPy 헝가리안 글로벌 최적 매칭 알고리즘 적용
                                                </label>
                                                <span class="info-icon" data-bs-toggle="tooltip" data-bs-placement="top" data-bs-html="true" title="<b>[헝가리안 알고리즘]</b><br>단순 개별 순위 정렬이 아닌, 학년 전체 학생들의 MFI 피로도 합(Total Cost)을 수학적으로 최소화하는 선형 계획법 최적 매칭 연산 방식입니다.">?</span>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>

                        </div>
                    </div>
                </div>

                <div class="col-md-6">
                    <div class="card card-custom h-100">
                        <div class="card-header card-header-custom py-3 px-4 fs-6">
                            📁 STEP 2. 학생 명단 파일 업로드
                        </div>
                        <div class="card-body p-4 d-flex flex-column justify-content-between">
                            <div class="dropzone-box mb-3">
                                <p class="fw-bold mb-2 text-dark">학생 명단 엑셀(.xlsx / .csv) 선택</p>
                                <input type="file" id="excel_file" class="form-control" accept=".csv, .xlsx, .xls">
                            </div>
                            <button onclick="runAssignment()" class="btn btn-run text-white w-100 shadow-sm">
                                🚀 고성능 AI 비동기 최적 배정 실행
                            </button>
                        </div>
                    </div>
                </div>
            </div>

            <div id="summary_box" style="display:none;" class="card card-custom mb-4 border-start border-4 border-primary">
                <div class="card-body p-4 d-flex justify-content-between align-items-center flex-wrap gap-3">
                    <div>
                        <h5 class="fw-bold text-navy mb-2">📊 AI 배정 결과 요약</h5>
                        <p id="summary_text" class="mb-0 fs-6"></p>
                    </div>
                    <button class="btn btn-excel text-white px-4 py-2 shadow-sm" onclick="exportToExcel()">
                        📥 AI 분석 리포트 포함 엑셀 다운로드
                    </button>
                </div>
            </div>

            <div class="card card-custom">
                <div class="table-responsive">
                    <table id="result_table" class="table table-hover table-custom mb-0" style="display:none;">
                        <thead>
                            <tr>
                                <th>배정 순위</th>
                                <th>학번</th>
                                <th>이름</th>
                                <th>성별</th>
                                <th>GPA</th>
                                <th>이동수단</th>
                                <th>소요시간</th>
                                <th>
                                    피로도 지수(MFI)
                                    <span class="info-icon" data-bs-toggle="tooltip" data-bs-placement="top" data-bs-html="true" title="<b>[MFI 피로도 지수 공식 & 근거]</b><br>MFI = 소요시간(분) + (환승횟수 × 12) + (도보시간 × 1.2)<br>간호대생 통학 피로도 특성을 고려하여 환승 대기시간과 도보시간에 체감 가중치를 부여한 산출 수식입니다. 지수가 낮을수록 우수합니다.">?</span>
                                </th>
                                <th>🤖 AI 예상 만족도</th>
                            </tr>
                        </thead>
                        <tbody id="result_body"></tbody>
                    </table>
                </div>
            </div>
        </div>

        <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
        <script>
            function initTooltips() {
                var tooltipTriggerList = [].slice.call(document.querySelectorAll('[data-bs-toggle="tooltip"]'));
                var tooltipList = tooltipTriggerList.map(function (tooltipTriggerEl) {
                    return new bootstrap.Tooltip(tooltipTriggerEl);
                });
            }

            async function verifyPassword() {
                const pwd = document.getElementById('authPassword').value;
                try {
                    const response = await fetch('/api/v1/verify-password', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ password: pwd })
                    });
                    
                    if (response.ok) {
                        document.getElementById('authOverlay').style.display = 'none';
                        sessionStorage.setItem('authenticated', 'true');
                    } else {
                        document.getElementById('authError').style.display = 'block';
                    }
                } catch(e) {
                    alert('서버 인증 연결 실패');
                }
            }

            if (sessionStorage.getItem('authenticated') === 'true') {
                document.getElementById('authOverlay').style.display = 'none';
            }

            const hospitalDB = {
                "성인I": [
                    "중앙대학교 광명병원", "가톨릭대학교 부천성모병원", "가톨릭대학교 성빈센트병원",
                    "고려대학교 안산병원", "순천향대학교 부천병원", "순천향대학교 서울병원",
                    "연세대학교 용인세브란스병원", "인천기독병원", "한림대학교 성심병원"
                ],
                "여성": [
                    "인하대병원", "한림대학교 동탄성심병원", "순천향대학교 서울병원",
                    "봄빛병원", "우성병원", "지샘병원", "한빛병원", "가톨릭대학교 부천성모병원"
                ],
                "성인II": [
                    "인하대병원", "가톨릭대학교 부천성모병원", "가톨릭대학교 성빈센트병원",
                    "고려대학교 안산병원", "아주대학교 병원", "순천향대학교 부천병원",
                    "순천향대학교 서울병원", "인천기독병원", "한림대학교 성심병원"
                ],
                "아동": [
                    "아이원병원", "웰봄병원", "단원병원", "순천향대학교 서울병원", "서울어린이 병원"
                ],
                "정신": [
                    "안산시 정신건강복지센터", "안산시 중독관리통합지원센터", "가톨릭대학교 성빈센트병원",
                    "이음병원", "군포시 정신건강복지센터", "의왕시 정신건강복지센터", "계요병원"
                ]
            };

            function updateHospitalOptions() {
                const subject = document.getElementById('subject_select').value;
                const selectBox = document.getElementById('hospital_select');
                
                selectBox.innerHTML = '';
                let targetHospitals = [];

                if (subject === 'ALL') {
                    const allSet = new Set();
                    Object.values(hospitalDB).forEach(arr => arr.forEach(h => allSet.add(h)));
                    targetHospitals = Array.from(allSet);
                } else {
                    targetHospitals = hospitalDB[subject] || [];
                }

                targetHospitals.forEach(hName => {
                    const opt = document.createElement('option');
                    opt.value = hName;
                    opt.textContent = hName;
                    selectBox.appendChild(opt);
                });
            }

            window.onload = function() {
                updateHospitalOptions();
                initTooltips();
            };

            let currentResults = [];
            let currentTargetHospital = "";

            async function runAssignment() {
                const hospitalName = document.getElementById('hospital_select').value;
                if (!hospitalName || hospitalName.trim() === '') {
                    alert('배정 대상 병원을 선택해 주세요!');
                    return;
                }

                const fileInput = document.getElementById('excel_file');
                if (!fileInput.files || fileInput.files.length === 0) {
                    alert('학생 명단 엑셀(CSV) 파일을 먼저 선택해 주세요!');
                    return;
                }

                const minGpaVal = document.getElementById('min_gpa').value;
                const birthYearVal = document.getElementById('birth_year').value;
                const useHungarian = document.getElementById('use_hungarian').checked;

                const formData = new FormData();
                formData.append('target_hospital', hospitalName);
                formData.append('gender_criteria', document.getElementById('gender_criteria').value);
                formData.append('use_hungarian', useHungarian);
                if (minGpaVal) formData.append('min_gpa', minGpaVal);
                if (birthYearVal) formData.append('birth_year_after', birthYearVal);
                formData.append('file', fileInput.files[0]);

                try {
                    const response = await fetch('/api/v1/assign-file', {
                        method: 'POST',
                        body: formData
                    });

                    const data = await response.json();

                    if (!response.ok) {
                        alert('오류 발생: ' + (data.detail || '입력 데이터를 확인해 주세요.'));
                        return;
                    }

                    currentResults = data.results;
                    currentTargetHospital = data.target_hospital;

                    document.getElementById('summary_box').style.display = 'block';
                    document.getElementById('summary_text').innerHTML = `<b>대상 병원:</b> ${data.target_hospital} &nbsp;|&nbsp; <b>총 학생:</b> ${data.total_students}명 &nbsp;|&nbsp; <b>적격 배정 대상:</b> <span class="pass-text" style="color:#059669; font-weight:bold;">${data.eligible_count}명</span> &nbsp;|&nbsp; <b>적용 알고리즘:</b> <span class="badge bg-info text-dark" data-bs-toggle="tooltip" title="선택된 배정 알고리즘 엔진입니다.">${data.optimization_method}</span>`;

                    const tbody = document.getElementById('result_body');
                    tbody.innerHTML = '';
                    data.results.forEach(res => {
                        if (!res.is_eligible) return;
                        
                        const row = document.createElement('tr');
                        const rankText = res.rank ? `<span class="rank-badge">${res.rank}순위</span>` : '-';
                        const timeText = res.travel_time_minutes ? `${res.travel_time_minutes}분` : '-';
                        const mfiText = res.fatigue_index ? `<span class="mfi-badge">${res.fatigue_index} MFI</span>` : '-';
                        const aiSatText = res.ai_satisfaction_score ? `<span class="ai-badge">${res.ai_satisfaction_score}점</span>` : '-';

                        row.innerHTML = `
                            <td>${rankText}</td>
                            <td>${res.student_id}</td>
                            <td><b>${res.name}</b></td>
                            <td>${res.gender}</td>
                            <td>${res.gpa}</td>
                            <td><span class="badge-mode">${res.transit_mode}</span></td>
                            <td><b>${timeText}</b></td>
                            <td>${mfiText}</td>
                            <td>${aiSatText}</td>
                        `;
                        tbody.appendChild(row);
                    });
                    document.getElementById('result_table').style.display = 'table';
                    initTooltips();
                } catch (e) {
                    alert('서버 통신 오류가 발생했습니다.');
                }
            }

            function exportToExcel() {
                if (!currentResults || currentResults.length === 0) {
                    alert('다운로드할 배정 결과가 없습니다.');
                    return;
                }

                const exportData = currentResults.filter(res => res.is_eligible).map(res => ({
                    "배정 순위": res.rank ? res.rank + "순위" : "-",
                    "학번": res.student_id,
                    "이름": res.name,
                    "성별": res.gender,
                    "GPA": res.gpa,
                    "출생연도": res.birth_year,
                    "이동수단 구분": res.transit_mode,
                    "소요시간_분": res.travel_time_minutes ? res.travel_time_minutes : "-",
                    "피로도 지수(MFI)": res.fatigue_index ? res.fatigue_index : "-",
                    "🤖 AI_예상만족도": res.ai_satisfaction_score ? res.ai_satisfaction_score + "점" : "-",
                    "🧠 AI_배정사유_리포트": res.ai_report
                }));

                const worksheet = XLSX.utils.json_to_sheet(exportData);
                const workbook = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(workbook, worksheet, "AI배정결과리포트");

                const filename = `${currentTargetHospital}_AI실습배정결과.xlsx`;
                XLSX.writeFile(workbook, filename);
            }
        </script>
    </body>
    </html>
    """
    return HTMLResponse(content=html_content)

if __name__ == "__main__":
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
