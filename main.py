from fastapi import FastAPI, HTTPException
from fastapi.responses import HTMLResponse
from pydantic import BaseModel, Field
from typing import List, Optional
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

class AssignmentRequest(BaseModel):
    target_hospital: str
    criteria: HospitalCriteria
    students: List[StudentInput]

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

@app.post("/api/v1/assign", response_model=AssignmentResponse)
def assign_hospital(payload: AssignmentRequest):
    if not payload.students:
        raise HTTPException(status_code=400, detail="학생 목록이 비어 있습니다.")

    eligible_list = []
    ineligible_list = []

    for stu in payload.students:
        is_ok, note = check_eligibility(stu, payload.criteria)
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
        status="success", target_hospital=payload.target_hospital,
        total_students=len(payload.students), eligible_count=len(eligible_list), results=final_results
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
        <style>
            body { font-family: 'Malgun Gothic', sans-serif; background-color: #f4f7f6; margin: 15px; }
            .container { max-width: 1000px; margin: 0 auto; background: white; padding: 20px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); }
            h1 { color: #1a365d; border-bottom: 3px solid #2b6cb0; padding-bottom: 10px; font-size: 22px; }
            .panel { background: #edf2f7; padding: 15px; border-radius: 8px; margin-bottom: 20px; }
            label { font-weight: bold; margin-right: 10px; display: inline-block; margin-top: 5px; }
            select, input { padding: 8px 12px; margin-right: 10px; margin-bottom: 10px; border-radius: 4px; border: 1px solid #cbd5e0; }
            button { background: #3182ce; color: white; border: none; padding: 12px 20px; font-size: 16px; font-weight: bold; border-radius: 6px; cursor: pointer; width: 100%; }
            button:hover { background: #2b6cb0; }
            .table-responsive { overflow-x: auto; }
            table { width: 100%; border-collapse: collapse; margin-top: 20px; min-width: 600px; }
            th, td { border: 1px solid #e2e8f0; padding: 10px; text-align: center; font-size: 14px; }
            th { background: #2b6cb0; color: white; }
            tr:nth-child(even) { background: #f7fafc; }
            .pass { color: #2f855a; font-weight: bold; }
            .fail { color: #e53e3e; font-weight: bold; }
            .rank-badge { background: #d69e2e; color: white; padding: 4px 8px; border-radius: 12px; font-weight: bold; }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>🏥 로켓단 AI 실습지 최적 배정 대시보드</h1>
            <div class="panel">
                <h3>📌 병원 자격 조건 설정</h3>
                <label>배정 대상 병원:</label>
                <input type="text" id="hospital_name" value="가톨릭대학교 부천성모병원"><br>
                <label>성별 조건:</label>
                <select id="gender_criteria">
                    <option value="남성만" selected>남성만</option>
                    <option value="여성만">여성만</option>
                    <option value="무관">무관</option>
                </select>
                <label>최소 GPA:</label>
                <input type="number" step="0.1" id="min_gpa" value="3.5">
                <label>출생연도 조건:</label>
                <input type="number" id="birth_year" value="2003"> 이후 출생자<br><br>
                <button onclick="runAssignment()">🚀 실시간 배정 알고리즘 실행</button>
            </div>
            <div id="summary_box" style="display:none;" class="panel">
                <h3>📊 배정 결과 요약</h3>
                <p id="summary_text"></p>
            </div>
            <div class="table-responsive">
                <table id="result_table" style="display:none;">
                    <thead>
                        <tr>
                            <th>배정 순위</th><th>학번</th><th>이름</th><th>성별</th><th>GPA</th><th>거주지 이동수단</th><th>소요시간</th><th>자격 검증 및 상태 메시지</th>
                        </tr>
                    </thead>
                    <tbody id="result_body"></tbody>
                </table>
            </div>
        </div>
        <script>
            async function runAssignment() {
                const payload = {
                    target_hospital: document.getElementById('hospital_name').value,
                    criteria: {
                        gender: document.getElementById('gender_criteria').value,
                        min_gpa: parseFloat(document.getElementById('min_gpa').value) || null,
                        birth_year_after: parseInt(document.getElementById('birth_year').value) || null
                    },
                    students: [
                        { student_id: "2024001", name: "강O우", gender: "여", gpa: 3.85, birth_year: 2003, address: "안산시 단원구 중앙대로 835", nearest_station: "고잔역", travel_time_minutes: 45, transit_mode: "버스+지하철 45분" },
                        { student_id: "2024002", name: "정O민", gender: "남", gpa: 4.10, birth_year: 2003, address: "광명시 일직로 12", nearest_station: "광명역", travel_time_minutes: 35, transit_mode: "지하철+도보 35분" },
                        { student_id: "2024003", name: "최O락", gender: "남", gpa: 3.65, birth_year: 2004, address: "부천시 원미구 소사로 170", nearest_station: "소사역", travel_time_minutes: 6, transit_mode: "도보 6분" },
                        { student_id: "2024004", name: "김O지", gender: "여", gpa: 3.90, birth_year: 2003, address: "안양시 동안구 관평로 170", nearest_station: "평촌역", travel_time_minutes: 40, transit_mode: "버스+지하철 40분" }
                    ]
                };

                const response = await fetch('/api/v1/assign', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });

                const data = await response.json();
                document.getElementById('summary_box').style.display = 'block';
                document.getElementById('summary_text').innerHTML = `<b>대상 병원:</b> ${data.target_hospital} | <b>총 학생:</b> ${data.total_students}명 | <b>적격 배정 대상:</b> <span class="pass">${data.eligible_count}명</span>`;

                const tbody = document.getElementById('result_body');
                tbody.innerHTML = '';
                data.results.forEach(res => {
                    const row = document.createElement('tr');
                    const rankText = res.rank ? `<span class="rank-badge">${res.rank}순위</span>` : '-';
                    const statusClass = res.is_eligible ? 'pass' : 'fail';
                    const timeText = res.travel_time_minutes ? `${res.travel_time_minutes}분` : '-';

                    row.innerHTML = `
                        <td>${rankText}</td><td>${res.student_id}</td><td><b>${res.name}</b></td>
                        <td>${res.gender}</td><td>${res.gpa}</td><td>${res.transit_mode}</td>
                        <td><b>${timeText}</b></td><td class="${statusClass}">${res.status_note}</td>
                    `;
                    tbody.appendChild(row);
                });
                document.getElementById('result_table').style.display = 'table';
            }
        </script>
    </body>
    </html>
    """
    return HTMLResponse(content=html_content)

if __name__ == "__main__":
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)