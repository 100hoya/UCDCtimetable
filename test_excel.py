import pandas as pd

print("✅ pandas:", pd.__version__)

df = pd.DataFrame({
    "무대": ["우아하게", "백백백"],
    "참가자": ["김소희, 강유빈", "김성환, 이윤호"]
})

print(df)
df.to_excel("예시.xlsx", index=False)
print("✅ 엑셀 파일 저장 완료!")
