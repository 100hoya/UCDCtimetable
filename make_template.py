# make_template.py
import pandas as pd

# 무대 시트: 이름, 길이(초), 참가자(쉼표로 구분), 고정순서(빈칸 가능)
stages = pd.DataFrame({
    "이름": ["우아하게", "백백백", "힙2"],
    "길이(초)": [225, 210, 200],
    "참가자": [
        "김소희, 강유빈, 박민호",
        "김성환, 이윤호",
        "한소민, 최서원, 권민준"
    ],
    "고정순서": ["", 2, ""],  # 예: '2'면 2번째 슬롯에 고정 배치
})

# 옵션 시트: 전역 설정(필요 시 확장)
options = pd.DataFrame({
    "옵션명": ["최소휴식슬롯", "후보안개수", "쉬는시간(초)"],
    "값": [1, 5, 60],
})

with pd.ExcelWriter("타임테이블_템플릿.xlsx", engine="openpyxl") as w:
    stages.to_excel(w, sheet_name="무대", index=False)
    options.to_excel(w, sheet_name="옵션", index=False)

print("✅ '타임테이블_템플릿.xlsx' 생성 완료")
