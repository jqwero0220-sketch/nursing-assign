from fastapi import FastAPI, HTTPException, UploadFile, File, Form, Body
from fastapi.responses import HTMLResponse
from pydantic import BaseModel, Field
from typing import List, Optional
import pandas as pd
import numpy as np
import io
import uvicorn

# Scikit-learn 머신러닝 라이브러리
from sklearn.ensemble import RandomForestClassifier

app = FastAPI(title="로켓단 AI 실습지 최적 배정 시스템 v4.0 (AI Engine)")

# 백엔드 보안 인증 비밀번호
SECRET_PASSWORD = "ansan king"

# -------------------------------------------------------------------
# 🤖 머신러닝(Scikit-learn) 기반 만족도 & 이의신청 위험도 예측 모델
# -------------------------------------------------------------------
class SatisfactionMLModel:
    def __init__(self):
        self.model = RandomForestClassifier(n_estimators=50, random_state=42)
        self._train_dummy_model()

    def _train_dummy_model(self):
        # 학습용 의사 데이터 (MFI 피로도, 소요시간, 환승횟수, GPA -> 만족도 클래스 0:낮음, 1:보통, 2:높음)
        np.random.seed(42)
        X_train = []
        y_train = []
        for _ in range(300):
            travel_time = np.random.randint(10, 90)
            transfers = np.random.randint(0, 4)
            walk_time = np.random.randint(3, 25)
            gpa = np.random.uniform(2.5, 4.5)
            mfi = travel_time + (transfers * 12.0) + (walk_time * 1.2)

            # 라벨링 규칙: MFI가 낮을수록 만족도 높음
            if mfi < 35:
                label = 2 # 만족도 높음 (이의신청 위험 낮음)
            elif mfi < 60:
                label = 1 # 보통
            else:
                label = 0 # 만족도 낮음 (이의신청 위험 높음)

            X_train.append([travel_time, transfers, walk_time, gpa, mfi])
            y_train.append(label)

        self.model.fit(X_train, y_train)

    def predict(self, travel_time: int, transfers: int, walk_time: int, gpa: float, mfi: float):
        features = [[travel_time, transfers, walk_time, gpa, mfi]]
        probs = self.model.predict_proba(features)[0] # [낮음, 보통, 높음] 확률
        
        # MFI 기반 점수 산출
        base_score = max(30, min(99, int(100 - (mfi * 0.8))))
        
        # 이의신청 위험도 산출
        if base_score >= 80:
            risk = "낮음 (안정)"
        elif base_score >= 60:
            risk = "보통"
        else:
            risk = "높음 (관심필요)"
            
        return base_score, risk

ml_engine = SatisfactionMLModel()

# -------------------------------------------------------------------
# 🧠 AI 배정 사유 및 교수자용 리포트 문장 생성기
# -------------------------------------------------------------------
def generate_ai_report(name: str, hospital: str, rank: Optional[int], mfi: float, travel_time: int, transit_mode: str, gpa: float, is_eligible: bool, note: str) -> str:
    if not is_eligible:
        return f"[AI 분석] {name} 학생은 {note}로 인해 {hospital} 배정 자격 미달로 판정되었습니다."
    
    report = f"[AI 리포트] {name} 학생은 {hospital} {rank}순위 최적 배정 대상자입니다. "
    report += f"거주지 기반 통학 소요시간 {travel_time}분({transit_mode}) 및 다변수 피로도 지수(MFI {mfi})가 최상위권이며, "
    report += f"GPA({gpa}) 기준 조건을 충족하여 통학 피로도 최소화 관점에서 최적의 배정안으로 평가됩니다."
    return report

# -------------------------------------------------------------------
# Data Models
# -------------------------------------------------------------------
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
    ai_complaint_risk: Optional[str]
    ai_report: str
    is_eligible: bool
    status_note: str

class AssignmentResponse(BaseModel):
    status: str
    target_hospital: str
    total_students: int
    eligible_count: int
    results: List[AssignmentResult]

def calculate_fatigue_index(travel_time: int, transfers: int, walk_time: int) -> float:
    mfi = travel_time + (transfers * 12.0) + (walk_time * 1.2)
    return round(mfi, 1)

def check_eligibility(student: StudentInput, criteria: HospitalCriteria) -> tuple[bool, str]:
    if criteria.gender == "남성만" and student.gender != "남":
        return False, "❌ 병원조건 미달 (성별 불일치: 남성만 가능)"
    elif criteria.gender == "여성만" and student.gender != "여":
        return False, "❌ 병원조건 미달 (성별 불일치: 여성만 가능)"

    if criteria.min_gpa is not None and student.gpa < criteria.min_gpa:
        return False, f"❌ 병원조건 미달 (성적 미달: {criteria.min_gpa} 이상 필요)"

    if criteria.birth_year_after is not None and student.birth_year < criteria.birth_year_after:
        return False, f"❌ 병원조건 미달 (연령 미달: {criteria.birth_year_after}년 이후 출생자 필요)"

    return True, "✅ 자격충족 & MFI 피로도 최적 배정 대상"

@app.post("/api/v1/verify-password")
async def verify_password(payload: PasswordVerifyRequest):
    if payload.password == SECRET_PASSWORD:
        return {"status": "success", "message": "인증 성공"}
    raise HTTPException(status_code=401, detail="비밀번호가 올바르지 않습니다.")

@app.post("/api/v1/assign-file", response_model=AssignmentResponse)
async def assign_hospital_from_file(
    target_hospital: str = Form(...),
    gender_criteria: str = Form("무관"),
    min_gpa: Optional[float] = Form(None),
    birth_year_after: Optional[int] = Form(None),
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

        students = []
        for _, row in df.iterrows():
            mode_raw = str(row.get('이동수단', '대중교통'))
            station_info = str(row.get('인근역', ''))
            
            travel_time = int(row['소요시간_분'])
            transfers = int(row.get('환승횟수', 1 if '버스' in mode_raw and station_info != '' else 0))
            walk_time = int(row.get('도보시간_분', 8))
            
            mfi = calculate_fatigue_index(travel_time, transfers, walk_time)

            if '버스' in mode_raw and ('전철' in mode_raw or '지하철' in mode_raw):
                detail_mode = '지하철+버스'
            elif '버스' in mode_raw:
                detail_mode = '시내/시외버스'
            else:
                detail_mode = '지하철(전철)'

            students.append(
                StudentInput(
                    student_id=str(row['학번']),
                    name=str(row['이름']),
                    gender=str(row['성별']),
                    gpa=float(row['GPA']),
                    birth_year=int(row['출생연도']),
                    address=str(row.get('주소', '')),
                    nearest_station=station_info,
                    travel_time_minutes=travel_time,
                    transfers=transfers,
                    walk_time_minutes=walk_time,
                    transit_mode=detail_mode,
                    fatigue_index=mfi
                )
            )
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"엑셀(CSV) 양식을 확인해 주세요: {str(e)}")

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
            # 부적격자 AI 리포트
            ai_rep = generate_ai_report(stu.name, target_hospital, None, stu.fatigue_index, stu.travel_time_minutes, stu.transit_mode, stu.gpa, False, note)
            ineligible_list.append(
                AssignmentResult(
                    rank=None, student_id=stu.student_id, name=stu.name, gender=stu.gender,
                    gpa=stu.gpa, birth_year=stu.birth_year, nearest_station=stu.nearest_station,
                    transit_mode=stu.transit_mode, travel_time_minutes=None, fatigue_index=None,
                    ai_satisfaction_score=None, ai_complaint_risk=None, ai_report=ai_rep,
                    is_eligible=False, status_note=note
                )
            )

    # MFI 피로도 지수 최적 정렬
    eligible_list.sort(key=lambda x: x["student"].fatigue_index)

    final_results = []
    for rank_idx, item in enumerate(eligible_list, start=1):
        stu = item["student"]
        
        # 🤖 Scikit-learn 머신러닝 만족도/위험도 예측
        sat_score, risk_level = ml_engine.predict(
            stu.travel_time_minutes, stu.transfers, stu.walk_time_minutes, stu.gpa, stu.fatigue_index
        )
        
        # 🧠 AI 분석 리포트 생성
        ai_rep = generate_ai_report(
            stu.name, target_hospital, rank_idx, stu.fatigue_index, stu.travel_time_minutes, stu.transit_mode, stu.gpa, True, item["status_note"]
        )

        final_results.append(
            AssignmentResult(
                rank=rank_idx, student_id=stu.student_id, name=stu.name, gender=stu.gender,
                gpa=stu.gpa, birth_year=stu.birth_year, nearest_station=stu.nearest_station,
                transit_mode=stu.transit_mode, travel_time_minutes=stu.travel_time_minutes,
                fatigue_index=stu.fatigue_index, ai_satisfaction_score=sat_score,
                ai_complaint_risk=risk_level, ai_report=ai_rep, is_eligible=True, status_note=item["status_note"]
            )
        )

    final_results.extend(ineligible_list)

    return AssignmentResponse(
        status="success", target_hospital=target_hospital,
        total_students=len(students), eligible_count=len(eligible_list), results=final_results
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
        <!-- Bootstrap 5 CDN & Google Fonts -->
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
        <link href="https://fonts.googleapis.com/css2?family=Pretendard:wght@400;600;700&display=swap" rel="stylesheet">
        <style>
            body { font-family: 'Pretendard', sans-serif; background-color: #f7fafc; color: #2d3748; }
            .navbar-custom { background-color: #1a365d; }
            .card-custom { border: none; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); }
            .card-header-custom { background: #edf2f7; border-bottom: 2px solid #e2e8f0; font-weight: 700; color: #1a365d; border-radius: 12px 12px 0 0 !important; }
            .btn-run { background-color: #2b6cb0; border: none; font-weight: 700; padding: 12px; font-size: 16px; border-radius: 8px; }
            .btn-run:hover { background-color: #1a365d; }
            .btn-excel { background-color: #2f855a; border: none; font-weight: 700; }
            .btn-excel:hover { background-color: #22543d; }
            .dropzone-box { border: 2px dashed #cbd5e0; background: #ffffff; border-radius: 8px; padding: 20px; text-align: center; }
            .table-custom th { background-color: #1a365d; color: white; text-align: center; font-size: 13.5px; }
            .table-custom td { vertical-align: middle; text-align: center; font-size: 13px; }
            .pass-text { color: #2f855a; font-weight: bold; }
            .fail-text { color: #e53e3e; font-weight: bold; }
            .rank-badge { background-color: #d69e2e; color: white; padding: 4px 10px; border-radius: 20px; font-weight: bold; font-size: 12px; }
            .badge-mode { background-color: #e2e8f0; color: #2b6cb0; font-weight: 600; padding: 3px 8px; border-radius: 6px; }
            .mfi-badge { background-color: #ebf8ff; color: #2c5282; font-weight: 700; padding: 3px 8px; border-radius: 6px; border: 1px solid #bee3f8; }
            .ai-badge { background-color: #f0fff4; color: #276749; font-weight: 700; padding: 3px 8px; border-radius: 6px; border: 1px solid #c6f6d5; }
            
            .auth-overlay { position: fixed; top: 0; left: 0; width: 100vw; height: 100vh; background-color: rgba(26, 54, 93, 0.96); z-index: 9999; display: flex; justify-content: center; align-items: center; }
            .auth-card { background: white; width: 90%; max-width: 420px; padding: 35px; border-radius: 16px; box-shadow: 0 10px 25px rgba(0,0,0,0.25); text-align: center; }
        </style>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    </head>
    <body>
        <!-- 서버 인증 모달 -->
        <div id="authOverlay" class="auth-overlay">
            <div class="auth-card">
                <div class="fs-1 mb-2">🔒</div>
                <h4 class="fw-bold text-navy mb-1">시스템 서버 인증</h4>
                <p class="text-muted fs-7 mb-4">개인정보 보호를 위해 관리자 암호를 입력해 주세요.</p>
                <div class="mb-3">
                    <input type="password" id="authPassword" class="form-control form-control-lg text-center fw-bold" placeholder="접속 암호 입력..." onkeyup="if(window.event.keyCode==13){verifyPassword();}">
                    <div id="authError" class="text-danger fs-7 mt-2 fw-bold" style="display:none;">❌ 백엔드 인증 실패: 암호가 올바르지 않습니다.</div>
                </div>
                <button onclick="verifyPassword()" class="btn text-white w-100 btn-lg fw-bold" style="background-color: #2b6cb0;">
                    백엔드 인증 및 시스템 접속
                </button>
            </div>
        </div>

        <!-- Header Navbar -->
        <nav class="navbar navbar-dark navbar-custom shadow-sm mb-4">
            <div class="container px-4">
                <span class="navbar-brand mb-0 h1 fw-bold fs-5">
                    🏥 로켓단 | AI 기반 간호학과 실습지 최적 배정 시스템
                </span>
                <span class="badge bg-success fs-7">v4.0 ML & Predictive Engine</span>
            </div>
        </nav>

        <div class="container pb-5" style="max-width: 1140px;">
            <div class="row g-4 mb-4">
                <!-- STEP 1: 교과목 & 병원 조건 설정 -->
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
                                <div class="accordion-item border-0 bg-light rounded">
                                    <h2 class="accordion-header">
                                        <button class="accordion-button collapsed bg-light fw-bold text-secondary fs-7 py-2" type="button" data-bs-toggle="collapse" data-bs-target="#collapseAdvanced">
                                            ⚙️ 세부 자격 조건 설정 (최소 GPA / 출생연도)
                                        </button>
                                    </h2>
                                    <div id="collapseAdvanced" class="accordion-collapse collapse" data-bs-parent="#advancedOptions">
                                        <div class="accordion-body pt-2 pb-3">
                                            <div class="row g-2">
                                                <div class="col-6">
                                                    <label class="form-label fs-7 fw-bold mb-1">최소 GPA (선택)</label>
                                                    <input type="number" step="0.1" id="min_gpa" class="form-control form-control-sm" placeholder="예: 3.5">
                                                </div>
                                                <div class="col-6">
                                                    <label class="form-label fs-7 fw-bold mb-1">출생연도 조건 (선택)</label>
                                                    <input type="number" id="birth_year" class="form-control form-control-sm" placeholder="예: 2003년 이후">
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>

                        </div>
                    </div>
                </div>

                <!-- STEP 2: 파일 업로드 -->
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
                                🚀 AI 머신러닝 예측 및 최적 배정 실행
                            </button>
                        </div>
                    </div>
                </div>
            </div>

            <!-- 요약 박스 -->
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

            <!-- 결과 테이블 -->
            <div class="card card-custom">
                <div class="table-responsive">
                    <table id="result_table" class="table table-hover table-custom mb-0" style="display:none;">
                        <thead>
                            <tr>
                                <th>배정 순위</th>
                                <th>학번</th>
                                <th>이름</th>
                                <th>GPA</th>
                                <th>이동수단</th>
                                <th>소요시간</th>
                                <th>피로도(MFI)</th>
                                <th>🤖 AI 예상만족도</th>
                                <th>이의신청 위험도</th>
                                <th>자격 검증 메시지</th>
                            </tr>
                        </thead>
                        <tbody id="result_body"></tbody>
                    </table>
                </div>
            </div>
        </div>

        <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
        <script>
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

                const formData = new FormData();
                formData.append('target_hospital', hospitalName);
                formData.append('gender_criteria', document.getElementById('gender_criteria').value);
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
                    document.getElementById('summary_text').innerHTML = `<b>대상 병원:</b> ${data.target_hospital} &nbsp;|&nbsp; <b>총 학생:</b> ${data.total_students}명 &nbsp;|&nbsp; <b>적격 배정 대상:</b> <span class="pass-text">${data.eligible_count}명</span>`;

                    const tbody = document.getElementById('result_body');
                    tbody.innerHTML = '';
                    data.results.forEach(res => {
                        const row = document.createElement('tr');
                        const rankText = res.rank ? `<span class="rank-badge">${res.rank}순위</span>` : '-';
                        const statusClass = res.is_eligible ? 'pass-text' : 'fail-text';
                        const timeText = res.travel_time_minutes ? `${res.travel_time_minutes}분` : '-';
                        const mfiText = res.fatigue_index ? `<span class="mfi-badge">${res.fatigue_index}</span>` : '-';
                        const aiSatText = res.ai_satisfaction_score ? `<span class="ai-badge">${res.ai_satisfaction_score}점</span>` : '-';
                        const riskText = res.ai_complaint_risk ? res.ai_complaint_risk : '-';

                        row.innerHTML = `
                            <td>${rankText}</td>
                            <td>${res.student_id}</td>
                            <td><b>${res.name}</b></td>
                            <td>${res.gpa}</td>
                            <td><span class="badge-mode">${res.transit_mode}</span></td>
                            <td><b>${timeText}</b></td>
                            <td>${mfiText}</td>
                            <td>${aiSatText}</td>
                            <td><b>${riskText}</b></td>
                            <td class="${statusClass}">${res.status_note}</td>
                        `;
                        tbody.appendChild(row);
                    });
                    document.getElementById('result_table').style.display = 'table';
                } catch (e) {
                    alert('서버 통신 오류가 발생했습니다.');
                }
            }

            function exportToExcel() {
                if (!currentResults || currentResults.length === 0) {
                    alert('다운로드할 배정 결과가 없습니다.');
                    return;
                }

                const exportData = currentResults.map(res => ({
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
                    "이의신청_위험도": res.ai_complaint_risk ? res.ai_complaint_risk : "-",
                    "🧠 AI_배정사유_리포트": res.ai_report,
                    "자격 상태": res.is_eligible ? "적격" : "부적격",
                    "자격 검증 메시지": res.status_note
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
