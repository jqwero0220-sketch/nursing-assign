from fastapi import FastAPI, HTTPException, UploadFile, File, Form
from fastapi.responses import HTMLResponse
from pydantic import BaseModel, Field
from typing import List, Optional
import pandas as pd
import io
import uvicorn

app = FastAPI(title="로켓단 AI 실습지 최적 배정 시스템")

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
    transit_mode: str

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
    is_eligible: bool
    status_note: str

class AssignmentResponse(BaseModel):
    status: str
    target_hospital: str
    total_students: int
    eligible_count: int
    results: List[AssignmentResult]

def check_eligibility(student: StudentInput, criteria: HospitalCriteria) -> tuple[bool, str]:
    if criteria.gender == "남성만" and student.gender != "남":
        return False, "❌ 병원조건 미달 (성별 불일치: 남성만 가능)"
    elif criteria.gender == "여성만" and student.gender != "여":
        return False, "❌ 병원조건 미달 (성별 불일치: 여성만 가능)"

    if criteria.min_gpa is not None and student.gpa < criteria.min_gpa:
        return False, f"❌ 병원조건 미달 (성적 미달: {criteria.min_gpa} 이상 필요)"

    if criteria.birth_year_after is not None and student.birth_year < criteria.birth_year_after:
        return False, f"❌ 병원조건 미달 (연령 미달: {criteria.birth_year_after}년 이후 출생자 필요)"

    return True, "✅ 조건충족 & 최단거리 배정 대상"

@app.post("/api/v1/assign-file", response_model=AssignmentResponse)
async def assign_hospital_from_file(
    target_hospital: str = Form(...),
    gender_criteria: str = Form("무관"),
    min_gpa: Optional[float] = Form(None),
    birth_year_after: Optional[int] = Form(None),
    file: UploadFile = File(...)
):
    if not target_hospital or target_hospital.strip() == "":
        raise HTTPException(status_code=400, detail="배정 대상 병원 이름을 입력해 주세요.")

    try:
        contents = await file.read()
        if file.filename.endswith('.csv'):
            df = pd.read_csv(io.BytesIO(contents))
        else:
            df = pd.read_excel(io.BytesIO(contents))

        students = []
        for _, row in df.iterrows():
            students.append(
                StudentInput(
                    student_id=str(row['학번']),
                    name=str(row['이름']),
                    gender=str(row['성별']),
                    gpa=float(row['GPA']),
                    birth_year=int(row['출생연도']),
                    address=str(row.get('주소', '')),
                    nearest_station=str(row.get('인근역', '')),
                    travel_time_minutes=int(row['소요시간_분']),
                    transit_mode=str(row.get('이동수단', '대중교통'))
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
            ineligible_list.append(
                AssignmentResult(
                    rank=None, student_id=stu.student_id, name=stu.name, gender=stu.gender,
                    gpa=stu.gpa, birth_year=stu.birth_year, nearest_station=stu.nearest_station,
                    transit_mode=stu.transit_mode, travel_time_minutes=None, is_eligible=False, status_note=note
                )
            )

    eligible_list.sort(key=lambda x: x["student"].travel_time_minutes)

    final_results = []
    for rank_idx, item in enumerate(eligible_list, start=1):
        stu = item["student"]
        final_results.append(
            AssignmentResult(
                rank=rank_idx, student_id=stu.student_id, name=stu.name, gender=stu.gender,
                gpa=stu.gpa, birth_year=stu.birth_year, nearest_station=stu.nearest_station,
                transit_mode=stu.transit_mode, travel_time_minutes=stu.travel_time_minutes,
                is_eligible=True, status_note=item["status_note"]
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
            .table-custom th { background-color: #1a365d; color: white; text-align: center; font-size: 14px; }
            .table-custom td { vertical-align: middle; text-align: center; font-size: 13.5px; }
            .pass-text { color: #2f855a; font-weight: bold; }
            .fail-text { color: #e53e3e; font-weight: bold; }
            .rank-badge { background-color: #d69e2e; color: white; padding: 4px 10px; border-radius: 20px; font-weight: bold; font-size: 12px; }
        </style>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    </head>
    <body>
        <!-- Header Navbar -->
        <nav class="navbar navbar-dark navbar-custom shadow-sm mb-4">
            <div class="container px-4">
                <span class="navbar-brand mb-0 h1 fw-bold fs-5">
                    🏥 로켓단 | AI 기반 간호학과 실습지 최적 배정 시스템
                </span>
                <span class="badge bg-secondary fs-7">v2.0 Web Dashboard</span>
            </div>
        </nav>

        <div class="container pb-5" style="max-width: 1080px;">
            <div class="row g-4 mb-4">
                <!-- STEP 1: 교과목 & 병원 조건 설정 -->
                <div class="col-md-6">
                    <div class="card card-custom h-100">
                        <div class="card-header card-header-custom py-3 px-4 fs-6">
                            📌 STEP 1. 교과목 및 병원 조건 설정
                        </div>
                        <div class="card-body p-4">
                            <!-- 실습 교과목 선택 추가 -->
                            <div class="mb-3">
                                <label class="form-label fw-bold">1. 실습 교과목 선택</label>
                                <select id="subject_select" class="form-select fw-bold text-primary" onchange="updateHospitalOptions()">
                                    <option value="ALL">전체 교과목 (26개 전체 병원)</option>
                                    <option value="성인I">성인간호학실습 I</option>
                                    <option value="여성">여성건강간호학실습</option>
                                    <option value="성인II">성인간호학실습 II</option>
                                    <option value="아동">아동간호학실습(학기중)</option>
                                    <option value="정신">정신간호학실습</option>
                                </select>
                            </div>

                            <div class="mb-3">
                                <label class="form-label fw-bold">2. 배정 대상 병원 선택</label>
                                <input type="text" id="hospital_name" list="hospital_list" class="form-select" placeholder="교과목 선택 시 해당 병원이 자동 검색됩니다..." value="고려대학교 안산병원">
                                <datalist id="hospital_list">
                                    <!-- JavaScript로 교과목별 자동 생성됨 -->
                                </datalist>
                            </div>

                            <div class="row g-2 mb-3">
                                <div class="col-md-4">
                                    <label class="form-label fw-bold">성별 조건</label>
                                    <select id="gender_criteria" class="form-select">
                                        <option value="무관" selected>무관</option>
                                        <option value="남성만">남성만</option>
                                        <option value="여성만">여성만</option>
                                    </select>
                                </div>
                                <div class="col-md-4">
                                    <label class="form-label fw-bold">최소 GPA</label>
                                    <input type="number" step="0.1" id="min_gpa" class="form-control" value="3.5">
                                </div>
                                <div class="col-md-4">
                                    <label class="form-label fw-bold">출생연도</label>
                                    <input type="number" id="birth_year" class="form-control" value="2003" placeholder="2003년 이후">
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
                                🚀 실시간 최적 배정 실행하기
                            </button>
                        </div>
                    </div>
                </div>
            </div>

            <!-- 요약 박스 -->
            <div id="summary_box" style="display:none;" class="card card-custom mb-4 border-start border-4 border-primary">
                <div class="card-body p-4 d-flex justify-content-between align-items-center flex-wrap gap-3">
                    <div>
                        <h5 class="fw-bold text-navy mb-2">📊 배정 결과 요약</h5>
                        <p id="summary_text" class="mb-0 fs-6"></p>
                    </div>
                    <button class="btn btn-excel text-white px-4 py-2 shadow-sm" onclick="exportToExcel()">
                        📥 결과 엑셀(Excel) 다운로드
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
                                <th>성별</th>
                                <th>GPA</th>
                                <th>이동수단</th>
                                <th>소요시간</th>
                                <th>자격 검증 및 상태 메시지</th>
                            </tr>
                        </thead>
                        <tbody id="result_body"></tbody>
                    </table>
                </div>
            </div>
        </div>

        <script>
            // 2026학년도 교과목별 실습지 데이터베이스 매핑
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
                const datalist = document.getElementById('hospital_list');
                const hospitalInput = document.getElementById('hospital_name');
                
                datalist.innerHTML = '';
                let targetHospitals = [];

                if (subject === 'ALL') {
                    // 전체 26개 중복제거 병원
                    const allSet = new Set();
                    Object.values(hospitalDB).forEach(arr => arr.forEach(h => allSet.add(h)));
                    targetHospitals = Array.from(allSet);
                } else {
                    targetHospitals = hospitalDB[subject] || [];
                }

                targetHospitals.forEach(hName => {
                    const opt = document.createElement('option');
                    opt.value = hName;
                    datalist.appendChild(opt);
                });

                if (targetHospitals.length > 0) {
                    hospitalInput.value = targetHospitals[0];
                }
            }

            // 페이지 로드 시 초기화
            window.onload = function() {
                updateHospitalOptions();
            };

            let currentResults = [];
            let currentTargetHospital = "";

            async function runAssignment() {
                const hospitalName = document.getElementById('hospital_name').value;
                if (!hospitalName || hospitalName.trim() === '') {
                    alert('배정 대상 병원 이름을 선택하거나 입력해 주세요!');
                    return;
                }

                const fileInput = document.getElementById('excel_file');
                if (!fileInput.files || fileInput.files.length === 0) {
                    alert('학생 명단 엑셀(CSV) 파일을 먼저 선택해 주세요!');
                    return;
                }

                const formData = new FormData();
                formData.append('target_hospital', hospitalName);
                formData.append('gender_criteria', document.getElementById('gender_criteria').value);
                formData.append('min_gpa', document.getElementById('min_gpa').value);
                formData.append('birth_year_after', document.getElementById('birth_year').value);
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

                        row.innerHTML = `
                            <td>${rankText}</td>
                            <td>${res.student_id}</td>
                            <td><b>${res.name}</b></td>
                            <td>${res.gender}</td>
                            <td>${res.gpa}</td>
                            <td>${res.transit_mode}</td>
                            <td><b>${timeText}</b></td>
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
                    "이동수단": res.transit_mode,
                    "소요시간_분": res.travel_time_minutes ? res.travel_time_minutes : "-",
                    "자격 상태": res.is_eligible ? "적격" : "부적격",
                    "자격 검증 및 상태 메시지": res.status_note
                }));

                const worksheet = XLSX.utils.json_to_sheet(exportData);
                const workbook = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(workbook, worksheet, "배정결과");

                const filename = `${currentTargetHospital}_실습배정결과.xlsx`;
                XLSX.writeFile(workbook, filename);
            }
        </script>
    </body>
    </html>
    """
    return HTMLResponse(content=html_content)

if __name__ == "__main__":
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
