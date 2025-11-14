# aicode.py - پردازش SEO هوشمند (آفلاین + Gemini اختیاری)
import pandas as pd
import os
import re
import json
import time
from collections import Counter, defaultdict
from difflib import SequenceMatcher
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils.dataframe import dataframe_to_rows

def run_seo(filepath, GEMINI_API_KEY=""):
    try:
        # -------------------------------------------------
        # 1. خواندن فایل اکسل
        # -------------------------------------------------
        df = pd.read_excel(filepath, usecols=[0,1])
        df.columns = ['عبارت', 'حجم جستجو']
        df = df.dropna()
        df['حجم جستجو'] = pd.to_numeric(df['حجم جستجو'], errors='coerce')
        df = df.dropna()
        
        # -------------------------------------------------
        # 2. تمیز کردن عبارات
        # -------------------------------------------------
        def تمیز(متن):
            متن = str(متن)
            متن = re.sub(r'[\u200c\u200d\u200e\u200f]', '', متن)
            متن = متن.replace('-', ' ').replace('_', ' ').replace('/', ' ')
            متن = re.sub(r'\s+', ' ', متن).strip().lower()
            return متن if متن else "نامشخص"
        
        df['تمیز'] = df['عبارت'].apply(تمیز)
        df = df[df['تمیز'] != "نامشخص"].copy()
        
        # -------------------------------------------------
        # 3. ادغام هوشمند عبارات مشابه
        # -------------------------------------------------
        def مشابه(س1, س2):
            return SequenceMatcher(None, س1, س2).ratio() > 0.85

        parent = list(range(len(df)))
        rank = [0] * len(df)

        def find(x):
            if parent[x] != x:
                parent[x] = find(parent[x])
            return parent[x]

        def union(x, y):
            px = find(x)
            py = find(y)
            if px != py:
                if rank[px] < rank[py]:
                    parent[px] = py
                elif rank[px] > rank[py]:
                    parent[py] = px
                else:
                    parent[py] = px
                    rank[px] += 1

        for i in range(len(df)):
            for j in range(i + 1, len(df)):
                if مشابه(df['تمیز'].iloc[i], df['تمیز'].iloc[j]):
                    union(i, j)

        groups = defaultdict(list)
        for i in range(len(df)):
            root = find(i)
            groups[root].append(i)

        ادغام_شده = 0
        new_rows = []

        for root, group in groups.items():
            if len(group) > 1:
                ادغام_شده += len(group) - 1
            total_volume = sum(df['حجم جستجو'].iloc[k] for k in group)
            main_idx = max(group, key=lambda k: df['حجم جستجو'].iloc[k])
            main_original = df['عبارت'].iloc[main_idx]
            main_clean = df['تمیز'].iloc[main_idx]
            new_rows.append({
                'عبارت': main_original,
                'حجم جستجو': total_volume,
                'تمیز': main_clean
            })

        df = pd.DataFrame(new_rows)
        
        # -------------------------------------------------
        # 4. Gemini یا حالت آفلاین
        # -------------------------------------------------
        استفاده_از_gemini = False
        نتایج = []
        if GEMINI_API_KEY and GEMINI_API_KEY.startswith("AIzaSy"):
            try:
                import google.generativeai as genai
                genai.configure(api_key=GEMINI_API_KEY)
                model = genai.GenerativeModel('gemini-2.0-flash')
                model.generate_content("تست")
                استفاده_از_gemini = True
            except:
                استفاده_از_gemini = False

        if استفاده_از_gemini:
            BATCH_SIZE = 5
            for i in range(0, len(df), BATCH_SIZE):
                دسته = df['تمیز'].iloc[i:i+BATCH_SIZE].tolist()
                پرامپت = f"""
تحلیل کن و فقط JSON بده:

{chr(10).join([f'{j+1}. "{ع}"' for j, ع in enumerate(دسته)])}

فرمت:
[
  {{"intent": "Informational|Commercial|Transactional|Navigational", "title": "عنوان H1 (حداکثر 60 کاراکتر)"}}
]
"""
                try:
                    پاسخ = model.generate_content(پرامپت)
                    نتیجه = پاسخ.text.strip()
                    if "```json" in نتیجه:
                        نتیجه = نتیجه.split("```json")[1].split("```")[0].strip()
                    elif "```" in نتیجه:
                        نتیجه = نتیجه.split("```")[1].strip()
                    تحلیل = json.loads(نتیجه)
                    if len(تحلیل) != len(دسته):
                        تحلیل = [{"intent": "Commercial", "title": ع} for ع in دسته]
                    نتایج.extend(تحلیل)
                except:
                    نتایج.extend([{"intent": "Commercial", "title": ع} for ع in دسته])
                time.sleep(6)
