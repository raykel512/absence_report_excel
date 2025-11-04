import streamlit as st
import pandas as pd
from datetime import date
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill

st.set_page_config(page_title="자동 결석 신고서 생성기 (Excel)", layout="centered")
st.title("📝 자동 결석 신고서 생성 (Excel 형식)")
st.caption("PDF 원본 양식에 최대한 유사하게 구조화된 Excel 파일을 생성합니다.")

# ----------------------------------------------------
# A. 데이터 입력값 설정 (이전과 동일)
# ----------------------------------------------------

# 예시 학생 명단 (선택 박스용)
STUDENTS = {
    "10101": {"학년": 1, "반": 1, "번호": 1, "이름": "김철수"},
    "10102": {"학년": 1, "반": 1, "번호": 2, "이름": "이영희"},
    "20315": {"학년": 2, "반": 3, "번호": 15, "이름": "박민재"},
}

st.subheader("1. 결석 학생 정보 입력")
student_options = {f"{s['학년']}-{s['반']}-{s['번호']} {s['이름']}": k for k, s in STUDENTS.items()}
selected_key = st.selectbox(
    "학생 선택",
    options=list(student_options.keys()),
    index=None
)

if selected_key:
    student_data = STUDENTS[student_options[selected_key]]
    
    # 총 일수 계산
    def calculate_days(start, end):
        if start > end: return 0
        return (end - start).days + 1
        
    st.subheader("2. 결석 기간 및 사유")
    
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("시작일", date.today())
    with col2:
        end_date = st.date_input("종료일", date.today())
    
    total_days = calculate_days(start_date, end_date)
    st.markdown(f"**👉 총 결석 예상 일수 (단순 계산): {total_days}일**")
        
    reason = st.text_area("결석 사유", "독감으로 인한 자가 격리")
    
    st.subheader("3. 결석 종류 및 첨부 서류 정보")
    absence_type = st.radio(
        "결석 종류 선택",
        options=['질병', '인정', '기타'],
        index=0
    )
    # PDF 양식의 첨부 서류 체크박스 반영 (3일 이상인 경우 첨부, 보건 결석 등)
    col_chk1, col_chk2 = st.columns(2)
    with col_chk1:
        has_diagnosis = st.checkbox("진단서/진료확인서 첨부 (3일 이상인 경우)", value=(total_days >= 3 and absence_type == '질병'))
    with col_chk2:
        has_opinion = st.checkbox("보건결석 학부모 의견서 첨부 (보건 결석인 경우)", value=(absence_type == '인정'))
        
    etc_doc_val = st.text_input("기타 첨부 서류 명칭", "")
    
    # ----------------------------------------------------
    # B. Excel 생성 및 PDF 양식 서식 적용 함수
    # ----------------------------------------------------
    
    def create_excel_report(data, has_diagnosis, has_opinion, etc_doc_val):
        wb = Workbook()
        ws = wb.active
        ws.title = "결석신고서"
        
        # --- 서식 정의 ---
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
        left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)
        title_font = Font(size=14, bold=True)
        header_font = Font(bold=True)
        
        # 열 너비 조정 (PDF 양식 칸 맞추기)
        ws.column_dimensions['A'].width = 15
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 15
        ws.column_dimensions['E'].width = 15
        ws.column_dimensions['F'].width = 15
        
        # --- 1. 결석 신고서 제목 ---
        current_row = 1
        ws.merge_cells(f'A{current_row}:F{current_row}')
        ws[f'A{current_row}'] = "결석신고서"
        ws[f'A{current_row}'].font = title_font
        ws[f'A{current_row}'].alignment = center_align
        ws.row_dimensions[current_row].height = 25
        
        # --- 2. 학생 정보 (A4 용지 칸처럼 병합) ---
        current_row += 2
        
        ws.merge_cells(f'A{current_row}:B{current_row}')
        ws[f'A{current_row}'] = "학생"
        ws[f'A{current_row}'].font = header_font
        ws[f'A{current_row}'].alignment = center_align
        ws[f'A{current_row}'].border = thin_border
        
        ws.merge_cells(f'C{current_row}:F{current_row}')
        ws[f'C{current_row}'] = f"{data['학년']}학년 {data['반']}반 {data['번호']}번"
        ws[f'C{current_row}'].alignment = left_align
        ws[f'C{C{current_row}'].border = thin_border
        
        # --- 3. 기간 ---
        current_row += 1
        ws.merge_cells(f'A{current_row}:B{current_row}')
        ws[f'A{current_row}'] = "기간"
        ws[f'A{current_row}'].font = header_font
        ws[f'A{current_row}'].alignment = center_align
        ws[f'A{current_row}'].border = thin_border
        
        period_str = f"{data['시작일'].strftime('%Y년 %m월 %d일')}부터 ~ {data['종료일'].strftime('%Y년 %m월 %d일')}까지 ({data['총_일수']}일간)"
        ws.merge_cells(f'C{current_row}:F{current_row}')
        ws[f'C{current_row}'] = period_str
        ws[f'C{current_row}'].alignment = left_align
        ws[f'C{current_row}'].border = thin_border
        ws.row_dimensions[current_row].height = 20
        
        # --- 4. 성명 ---
        current_row += 1
        ws.merge_cells(f'A{current_row}:B{current_row}')
        ws[f'A{current_row}'] = "성명"
        ws[f'A{current_row}'].font = header_font
        ws[f'A{current_row}'].alignment = center_align
        ws[f'A{current_row}'].border = thin_border
        
        ws.merge_cells(f'C{current_row}:F{current_row}')
        ws[f'C{current_row}'] = data['이름']
        ws[f'C{current_row}'].alignment = left_align
        ws[f'C{current_row}'].border = thin_border
        
        # --- 5. 사유 ---
        current_row += 1
        ws.merge_cells(f'A{current_row}:B{current_row}')
        ws[f'A{current_row}'] = "사유"
        ws[f'A{current_row}'].font = header_font
        ws[f'A{current_row}'].alignment = center_align
        ws[f'A{current_row}'].border = thin_border
        
        ws.merge_cells(f'C{current_row}:F{current_row}')
        ws[f'C{current_row}'] = data['사유']
        ws[f'C{current_row}'].alignment = left_align
        ws[f'C{C{current_row}'].border = thin_border
        ws.row_dimensions[current_row].height = 60 # 사유 칸 넓게
        
        # --- 6. 붙임 서류 ---
        current_row += 1
        ws.merge_cells(f'A{current_row}:B{current_row}')
        ws[f'A{current_row}'] = "붙임 서류"
        ws[f'A{current_row}'].font = header_font
        ws[f'A{current_row}'].alignment = center_align
        ws[f'A{current_row}'].border = thin_border
        
        doc_list = []
        doc_list.append(f"[{'X' if has_diagnosis else ' '}] 진단서 또는 진료 확인서 (3일 이상인 경우)")
        doc_list.append(f"[{'X' if has_opinion else ' '}] 보건결석 학부모 의견서")
        doc_list.append(f"[{'X' if not (has_diagnosis or has_opinion or etc_doc_val) else ' '}] 없음")
        if etc_doc_val:
            doc_list.append(f"[{'X'}] 기타 ({etc_doc_val})")

        ws.merge_cells(f'C{current_row}:F{current_row}')
        ws[f'C{current_row}'] = '\n'.join(doc_list)
        ws[f'C{current_row}'].alignment = left_align
        ws[f'C{current_row}'].border = thin_border
        ws.row_dimensions[current_row].height = 60
        
        # --- 7. 보호자 연서 (서명) ---
        current_row += 2
        ws.merge_cells(f'A{current_row}:F{current_row}')
        ws[f'A{current_row}'] = f"위와 같이 결석하고자 하였기에 보호자 연서로 신고합니다.  {date.today().strftime('%Y년 %m월 %d일')}"
        ws[f'A{current_row}'].alignment = left_align
        
        current_row += 1
        ws.merge_cells(f'A{current_row}:C{current_row}')
        ws[f'A{current_row}'] = f"학생 성명: {data['이름']}"
        ws.merge_cells(f'D{current_row}:F{current_row}')
        ws[f'D{current_row}'] = "보호자 성명: (서명 또는 인)"
        ws[f'A{current_row}'].alignment = left_align
        ws[f'D{current_row}'].alignment = left_align

        # --- 8. 담임교사 확인서 (새로운 섹션) ---
        current_row += 2
        ws.merge_cells(f'A{current_row}:F{current_row}')
        ws[f'A{current_row}'] = "담임교사 확인서"
        ws[f'A{current_row}'].font = title_font
        ws[f'A{current_row}'].alignment = center_align
        ws.row_dimensions[current_row].height = 25
        
        current_row += 1
        ws.merge_cells(f'A{current_row}:B{current_row}')
        ws[f'A{current_row}'] = "결석 종류"
        ws[f'A{current_row}'].font = header_font
        ws[f'A{current_row}'].alignment = center_align
        ws[f'A{current_row}'].border = thin_border
        
        chk_질병 = 'X' if data['결석_종류'] == '질병' else ' '
        chk_인정 = 'X' if data['결석_종류'] == '인정' else ' '
        chk_기타 = 'X' if data['결석_종류'] == '기타' else ' '
        
        ws.merge_cells(f'C{current_row}:F{current_row}')
        ws[f'C{current_row}'] = f"[{chk_질병}] 질병  [{chk_인정}] 인정  [{chk_기타}] 기타"
        ws[f'C{current_row}'].alignment = left_align
        ws[f'C{current_row}'].border = thin_border
        
        # 확인 방법 (간소화)
        current_row += 1
        ws.merge_cells(f'A{current_row}:B{current_row}')
        ws[f'A{current_row}'] = "확인 방법"
        ws[f'A{current_row}'].font = header_font
        ws[f'A{current_row}'].alignment = center_align
        ws[f'A{current_row}'].border = thin_border
        
        ws.merge_cells(f'C{current_row}:F{current_row}')
        ws[f'C{current_row}'] = "[X] 제출된 증빙서류로 확인"
        ws[f'C{current_row}'].alignment = left_align
        ws[f'C{current_row}'].border = thin_border
        
        # --- 9. 서명 및 결재 라인 ---
        current_row += 2
        ws.merge_cells(f'A{current_row}:F{current_row}')
        ws[f'A{current_row}'] = f"위의 신고 내용이 사실과 같음을 확인합니다.  {date.today().strftime('%Y년 %m월 %d일')}"
        ws[f'A{current_row}'].alignment = left_align

        # 결재 라인 헤더
        current_row += 1
        ws.merge_cells(f'A{current_row}:B{current_row}')
        ws[f'A{current_row}'] = "학급 담임"
        ws[f'A{C{current_row}'].border = thin_border
        ws[f'A{current_row}'].alignment = center_align
        
        ws[f'C{current_row}'] = "출결 담당"
        ws[f'C{current_row}'].border = thin_border
        ws[f'C{current_row}'].alignment = center_align
        
        ws[f'D{current_row}'] = "교무 부장"
        ws[f'D{current_row}'].border = thin_border
        ws[f'D{current_row}'].alignment = center_align
        
        ws.merge_cells(f'E{current_row}:F{current_row}')
        ws[f'E{current_row}'] = "교감"
        ws[f'E{current_row}'].border = thin_border
        ws[f'E{current_row}'].alignment = center_align
        
        # 최종 서명/결재 빈칸
        current_row += 1
        for col in ['A', 'B', 'C', 'D', 'E', 'F']:
            if col not in ['B', 'D', 'F']: # 병합된 셀은 건너뜀
                ws[f'{col}{current_row}'].border = thin_border
            ws.row_dimensions[current_row].height = 30
        
        ws.merge_cells(f'A{current_row}:B{current_row}')
        ws.merge_cells(f'E{current_row}:F{current_row}')
        
        # 학교장 귀하
        current_row += 1
        ws.merge_cells(f'A{current_row}:F{current_row}')
        ws[f'A{current_row}'] = "대동세무고등학교장 귀하"
        ws[f'A{current_row}'].alignment = Alignment(horizontal='right', vertical='center')

        return wb

    # ----------------------------------------------------
    # C. 파일 생성 및 다운로드
    # ----------------------------------------------------
    
    # 최종 대체 데이터 조합
    final_data = {
        "학년": student_data["학년"], "반": student_data["반"], "번호": student_data["번호"],
        "이름": student_data["이름"], "총_일수": total_days,
        "시작일": start_date, "종료일": end_date,
        "사유": reason, "결석_종류": absence_type
    }

    st.markdown("---")
    if st.button("결석 신고서 생성 및 다운로드 (Excel)", use_container_width=True):
        st.subheader("4. 결과 확인")
        
        workbook = create_excel_report(final_data, has_diagnosis, has_opinion, etc_doc_val)
        
        excel_buffer = BytesIO()
        workbook.save(excel_buffer)
        excel_buffer.seek(0)
        
        file_name = f"결석신고서_Excel_{final_data['이름']}_{final_data['시작일'].strftime('%Y%m%d')}.xlsx"
        
        st.download_button(
            label=f"📥 {file_name} 다운로드",
            data=excel_buffer,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        st.success("Excel 신고서 생성이 완료되었습니다! 다운로드 후 인쇄하여 사용하세요.")
        st.balloons()

else:
    st.info("먼저 결석한 학생을 선택해주세요.")
