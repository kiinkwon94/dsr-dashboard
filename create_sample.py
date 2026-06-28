"""샘플 데이터 생성 스크립트 - 한 번만 실행하면 됩니다."""
import pandas as pd
import random
from pathlib import Path

random.seed(42)

teams = {"A팀": ["김철수", "이영희"], "B팀": ["박민준", "최지원", "정수현"], "C팀": ["강동원", "윤소희"]}
rows = []

for day in range(1, 29):
    date = pd.Timestamp(f"2026-06-{day:02d}")
    if date > pd.Timestamp.today():
        break
    for team, members in teams.items():
        for person in members:
            target = random.randint(800, 1200)
            actual = int(target * random.uniform(0.6, 1.3))
            rows.append({"날짜": date, "팀": team, "담당자": person, "실적": actual, "목표": target})

df = pd.DataFrame(rows)
Path("data").mkdir(exist_ok=True)
out = "data/sales_2026_06.xlsx"
df.to_excel(out, index=False)
print(f"샘플 파일 생성 완료: {out}  ({len(df)}행)")
