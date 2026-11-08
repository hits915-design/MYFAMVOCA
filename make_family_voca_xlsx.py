"""우리가족 VOCA 엑셀 파일(family_voca.xlsx) 생성기
기존 단어 데이터를 바탕으로 4개 시트(아빠, 엄마, 은봉이, 지봉이)를 갖는 표준 엑셀 파일을 생성합니다.
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

DATA = {
    "아빠": [
        ["negotiate", "협상하다", "nuh-GOH-shee-ayt", "We need to negotiate the contract terms.", "Meeting", "", "Lv3"],
        ["deadline", "마감 기한", "DED-line", "Let’s confirm the deadline before the meeting.", "Project", "", "Lv2"],
        ["agenda", "안건", "uh-JEN-duh", "I’ll send the agenda by email today.", "Meeting", "", "Lv2"],
        ["revenue", "매출", "REV-uh-noo", "Revenue grew ten percent this quarter.", "Report", "", "Lv3"],
        ["stakeholder", "이해관계자", "STAYK-hohl-der", "Please update every stakeholder on the progress.", "Project", "", "Lv4"],
        ["budget", "예산", "BUJ-it", "The budget is tight, so we must prioritize.", "Finance", "", "Lv2"],
        ["approve", "승인하다", "uh-PROOV", "Could you approve this request by Friday?", "Email", "", "Lv2"],
        ["schedule", "일정을 잡다", "SKEJ-ool", "Can we schedule a call for Monday?", "Email", "", "Lv1"],
        ["feedback", "피드백", "FEED-bak", "Thank you for your valuable feedback.", "Email", "", "Lv1"],
        ["deliver", "납품하다", "dih-LIV-er", "We will deliver the report on time.", "Project", "", "Lv2"],
        ["forecast", "전망, 예측", "FOR-kast", "The sales forecast looks promising.", "Report", "", "Lv3"],
        ["attach", "첨부하다", "uh-TACH", "I will attach the revised document.", "Email", "", "Lv2"],
        ["postpone", "연기하다", "pohst-POHN", "We had to postpone the product launch.", "Meeting", "", "Lv3"],
        ["compromise", "타협", "KOM-pruh-myz", "Let’s find a compromise that works for both sides.", "Negotiation", "", "Lv4"]
    ],
    "엄마": [
        ["luggage", "짐, 수하물", "LUG-ij", "Where can I pick up my luggage?", "Airport", "", "Lv2"],
        ["boarding", "탑승", "BOR-ding", "The boarding gate is on the left.", "Airport", "", "Lv2"],
        ["passport", "여권", "PAS-port", "Please show your passport at the counter.", "Airport", "", "Lv1"],
        ["reservation", "예약", "rez-er-VAY-shun", "I have a reservation under Kim.", "Hotel", "", "Lv2"],
        ["available", "이용 가능한", "uh-VAY-luh-bul", "Is this table available for two?", "Restaurant", "", "Lv2"],
        ["recommend", "추천하다", "rek-uh-MEND", "What do you recommend for dinner?", "Restaurant", "", "Lv2"],
        ["delicious", "맛있는", "dih-LISH-us", "This soup is absolutely delicious.", "Restaurant", "", "Lv1"],
        ["receipt", "영수증", "rih-SEET", "Could I get a receipt, please?", "Shopping", "", "Lv1"],
        ["discount", "할인", "DIS-kount", "Is there a discount for this item?", "Shopping", "", "Lv1"],
        ["refund", "환불", "REE-fund", "I’d like a refund for this shirt.", "Shopping", "", "Lv2"],
        ["souvenir", "기념품", "soo-vuh-NEER", "I’m looking for a souvenir for my friend.", "Travel", "", "Lv3"],
        ["directions", "길 안내", "duh-REK-shunz", "Could you give me directions to the station?", "Travel", "", "Lv2"],
        ["fresh", "신선한", "FRESH", "These vegetables look so fresh.", "Daily", "", "Lv1"],
        ["neighbor", "이웃", "NAY-ber", "My neighbor brought us some kimchi.", "Daily", "", "Lv1"]
    ],
    "은봉이": [
        ["habitat", "서식지", "HAB-ih-tat", "Polar bears live in a cold habitat.", "Science", "", "Lv3"],
        ["enormous", "거대한", "ih-NOR-mus", "An enormous whale swam under the boat.", "Story", "", "Lv3"],
        ["ancient", "고대의", "AYN-shunt", "The children found an ancient map.", "Story", "", "Lv3"],
        ["discover", "발견하다", "dis-KUV-er", "Scientists discover new planets every year.", "Science", "", "Lv2"],
        ["courage", "용기", "KUR-ij", "She had the courage to speak first.", "Story", "", "Lv3"],
        ["invent", "발명하다", "in-VENT", "Edison tried to invent a better light bulb.", "Science", "", "Lv2"],
        ["explore", "탐험하다", "ik-SPLOR", "We explore the forest every summer.", "Story", "", "Lv2"],
        ["gravity", "중력", "GRAV-ih-tee", "Gravity pulls the apple to the ground.", "Science", "", "Lv3"],
        ["predict", "예측하다", "pri-DIKT", "Can you predict what happens next?", "Reading", "", "Lv3"],
        ["nervous", "긴장한", "NUR-vus", "He felt nervous before the big show.", "Story", "", "Lv2"],
        ["treasure", "보물", "TREZH-er", "The pirates buried their treasure on the island.", "Story", "", "Lv2"],
        ["whisper", "속삭이다", "WIS-per", "Please whisper in the library.", "Story", "", "Lv2"],
        ["mystery", "수수께끼", "MIS-tuh-ree", "The missing key was a real mystery.", "Reading", "", "Lv3"],
        ["gather", "모으다", "GATH-er", "Squirrels gather nuts in autumn.", "Science", "", "Lv2"]
    ],
    "지봉이": [
        ["apple", "사과", "AP-ul", "A is for apple. I like apple.", "Fruit", "🍎", "Lv1"],
        ["banana", "바나나", "buh-NA-nuh", "The monkey likes banana.", "Fruit", "🍌", "Lv1"],
        ["dog", "강아지", "DOG", "D is for dog. The dog is big.", "Animal", "🐶", "Lv1"],
        ["cat", "고양이", "KAT", "C is for cat. The cat is cute.", "Animal", "🐱", "Lv1"],
        ["red", "빨강", "RED", "The apple is red.", "Color", "🔴", "Lv1"],
        ["blue", "파랑", "BLOO", "The sky is blue.", "Color", "🔵", "Lv1"],
        ["sun", "해", "SUN", "The sun is hot.", "Nature", "☀️", "Lv1"],
        ["moon", "달", "MOON", "I see the moon.", "Nature", "🌙", "Lv1"],
        ["fish", "물고기", "FISH", "The fish can swim.", "Animal", "🐟", "Lv1"],
        ["milk", "우유", "MILK", "I drink milk.", "Food", "🥛", "Lv1"],
        ["bird", "새", "BURD", "The bird can fly.", "Animal", "🐦", "Lv1"],
        ["car", "자동차", "KAR", "The car is fast.", "Thing", "🚗", "Lv1"],
        ["star", "별", "STAR", "Look at the star.", "Nature", "⭐", "Lv1"],
        ["bear", "곰", "BAIR", "B is for bear. The bear is brown.", "Animal", "🐻", "Lv1"],
        ["egg", "달걀", "EG", "I eat an egg.", "Food", "🥚", "Lv1"],
        ["hand", "손", "HAND", "Clap your hand.", "Body", "✋", "Lv1"]
    ]
}

HEADERS = ["번호", "단어", "뜻", "발음", "예문", "카테고리", "그림", "난이도"]

def make_excel(filename="family_voca.xlsx"):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # 기본 시트 제거

    header_font = Font(name="맑은 고딕", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="3B82F6", end_color="3B82F6", fill_type="solid")
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")
    thin_border = Border(
        left=Side(style='thin', color='DDDDDD'),
        right=Side(style='thin', color='DDDDDD'),
        top=Side(style='thin', color='DDDDDD'),
        bottom=Side(style='thin', color='DDDDDD')
    )

    for sheet_name, rows in DATA.items():
        ws = wb.create_sheet(title=sheet_name)
        
        # 1행 제목
        ws.cell(row=1, column=1, value=f"[{sheet_name}] 단어장")
        ws.cell(row=1, column=1).font = Font(name="맑은 고딕", size=14, bold=True)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(HEADERS))
        
        # 2행 가이드
        ws.cell(row=2, column=1, value="※ 4행부터 단어 데이터입니다. 단어/뜻/예문은 필수이며 예문 안에 단어가 반드시 포함되어야 합니다. 난이도는 Lv1~Lv5 중 하나입니다.")
        ws.cell(row=2, column=1).font = Font(name="맑은 고딕", size=9, color="666666", italic=True)
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(HEADERS))

        # 3행 헤더
        for col_idx, h in enumerate(HEADERS, start=1):
            cell = ws.cell(row=3, column=col_idx, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center_align
            cell.border = thin_border

        # 4행~ 데이터
        for row_idx, item in enumerate(rows, start=4):
            # 번호
            num_cell = ws.cell(row=row_idx, column=1, value=row_idx-3)
            num_cell.alignment = center_align
            num_cell.border = thin_border
            
            for col_idx, val in enumerate(item, start=2):
                c = ws.cell(row=row_idx, column=col_idx, value=val)
                c.border = thin_border
                c.font = Font(name="맑은 고딕", size=10)
                if col_idx in (2, 7, 8):  # 단어, 그림, 난이도
                    c.alignment = center_align
                else:
                    c.alignment = left_align

        # 열 너비 자동 조정
        col_widths = {1: 8, 2: 18, 3: 20, 4: 20, 5: 45, 6: 15, 7: 10, 8: 10}
        for col_idx, width in col_widths.items():
            ws.column_dimensions[get_column_letter(col_idx)].width = width

    wb.save(filename)
    print(f"'{filename}' 파일 생성 완료!")

if __name__ == "__main__":
    make_excel()
