#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys, pdfplumber
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8')

pdfs = [
    'NOVELATION_確定申告_202305.pdf',
    '法人税及び地方法人税申告書_20240529株式会社ＮＯＶＥＬＡＴＩＯＮ.pdf',
    'NOVELATION_確定申告_202505.pdf',
]
for pdf_name in pdfs:
    print(f'\n=== {pdf_name} ===')
    with pdfplumber.open(pdf_name) as pdf:
        print(f'総ページ数: {len(pdf.pages)}')
        for p in pdf.pages:
            text = p.extract_text() or ''
            flags = []
            if '貸借対照表' in text: flags.append('BS')
            if '損益計算書' in text: flags.append('PL')
            if '法人事業概況' in text: flags.append('jigyoukyo')
            if flags:
                print(f'  page {p.page_number}: {flags} (chars={len(text)})')
