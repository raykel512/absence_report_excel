import streamlit as st
import pandas as pd
from datetime import date
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill

st.set_page_config(page_title="자동 결석 신고서 생성기 (Excel)", layout="centered")
st.title("📝 자동 결석 신고서 생성 (Excel 형식)")
st.caption("Excel 파일에 깔끔한 보고서 서식을 적용하여 즉시 다운로드합니다.")

# ----------------------------------------------------
# A. 데이터 입력값 설정
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
    etc_doc = st.text_input("기타 첨부 서류 명칭", "")
    
    # ----------------------------------------------------
    # B. Excel 생성 및 서식 적용 함수
    # ----------------------------------------------------
    
    def create_excel_report(data):
        wb = Workbook()
        ws = wb.active
        ws.title = "결석신고서"
        
        # 기본 서식 정의
        thin_border = Border(left=Side(style='thin'), 
                             right=Side(style='thin'), 
                             top=Side(style='thin'), 
                             bottom=Side(style='thin'))
        bold_font = Font(bold=True)
        center_align = Alignment(horizontal='center', vertical='center')
        
        # 1. 문서 제목
        ws.merge_cells('A1:F1')
        ws['A1'] = "학생 결석 신고서 및 담임교사 확인서"
        ws['A1'].font = Font(size=18, bold=True)
        ws['A1'].alignment = center_align
        ws.row_dimensions[1].height = 30
        
        # 2. 신고 내역 (표로 깔끔하게)
        start_row = 3
        
        report_data = [
            ("학생 정보", f"{data['학년']}학년 {data['반']}반 {data['번호']}번 {data['이름']}"),
            ("결석 기간", f"{data['시작일'].strftime('%Y년 %m월 %d일')} ~ {data['종료일'].strftime('%Y년 %m월 %d일')} (총 {data['총_일수']}일간)"),
            ("결석 사유", data['사유']),
            ("신고 일자", date.today().strftime('%Y년 %m월 %d일')),
            ("첨부 서류", f"진단서/진료확인서 등, 기타: {data['기타_서류']}"),
        ]
        
        # 데이터 채우기 및 서식 적용
        for i, (label, value) in enumerate(report_data):
            row = start_row + i
            # 제목 셀 (A열)
            ws[f'A{row}'] = label
            ws.merge_cells(f'A{row}:B{row}')
            ws[f'A{row}'].font = bold_font
            ws[f'A{row}'].alignment = center_align
            ws[f'A{row}'].fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid") # 회색 배경
            
            # 내용 셀 (C~F열)
            ws[f'C{row}'] = value
            ws.merge_cells(f'C{row}:F{row}')
            ws[f'C{row}'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            
            # 테두리 적용
            for col in ['A', 'B', 'C', 'D', 'E', 'F']:
                ws[f'{col}{row}'].border = thin_border
                
        # 3. 담임교사 확인 (아래에 이어서)
        current_row = start_row + len(report_data) + 1
        
        ws.merge_cells(f'A{current_row}:F{current_row}')
        ws[f'A{current_row}'] = "II. 담임교사 확인 및 처리"
        ws[f'A{current_row}'].font = bold_font
        ws[f'A{current_row}'].fill = PatternFill(start_color="CCCCFF", end_color="CCCCFF", fill_type="solid") # 연한 파랑 배경
        ws[f'A{current_row}'].border = thin_border
        
        current_row += 1
        
        # 결석 종류 표시 (체크박스 대신 텍스트로 강조)
        ws.merge_cells(f'A{current_row}:F{current_row}')
        ws[f'A{current_row}'] = f"결석 종류: [{data['결석_종류']}] {data['결석_종류']} 결석 (확인 일자: {date.today().strftime('%Y년 %m월 %d일')})"
        ws[f'A{current_row}'].alignment = Alignment(horizontal='left', vertical='center')
        ws[f'A{current_row}'].border = thin_border
        ws.row_dimensions[current_row].height = 25
        
        current_row += 1
        
        # 서명란 (병합을 많이 사용)
        ws.merge_cells(f'A{current_row}:B{current_row}')
        ws[f'A{current_row}'] = "확인자 (담임)"
        ws[f'A{current_row}'].font = bold_font
        ws[f'A{current_row}'].alignment = center_align
        ws[f'A{current_row}'].border = thin_border
        
        ws.merge_cells(f'C{current_row}:F{current_row}')
        ws[f'C{current_row}'] = "(서명 또는 인)"
        ws[f'C{current_row}'].alignment = Alignment(horizontal='right', vertical='bottom')
        ws[f'C{current_row}'].border = thin_border
        ws.row_dimensions[current_row].height = 40 # 서명 공간 확보
        
        # 열 너비 조정
        ws.column_dimensions['A'].width = 15
        ws.column_dimensions['C'].width = 20
        
        return wb

    # ----------------------------------------------------
    # C. 파일 생성 및 다운로드
    # ----------------------------------------------------
    
    # 최종 대체 데이터 조합
    final_data = {
        "학년": student_data["학년"], "반": student_data["반"], "번호": student_data["번호"],
        "이름": student_data["이름"], "총_일수": total_days,
        "시작일": start_date, "종료일": end_date,
        "사유": reason, "결석_종류": absence_type, "기타_서류": etc_doc
    }

    st.markdown("---")
    if st.button("결석 신고서 생성 및 다운로드 (Excel)", use_container_width=True):
        st.subheader("4. 결과 확인")
        
        # Excel 문서 생성
        workbook = create_excel_report(final_data)
        
        # BytesIO를 사용하여 메모리에 문서를 저장하고 Streamlit 다운로드에 사용
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
        st.success("Excel 신고서 생성이 완료되었습니다!")
        st.balloons()

else:
    st.info("먼저 결석한 학생을 선택해주세요.")