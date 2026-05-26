#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import pdfplumber
import sys

sys.stdout.reconfigure(encoding='utf-8')

pdf_path = 'KAISEI第5期2023年7月期決算報告書.pdf'

with pdfplumber.open(pdf_path) as pdf:
    for i, page in enumerate(pdf.pages):
        print(f'\n=== PAGE {i+1} (width={page.width:.0f}) ===')
        text = page.extract_text() or ''
        print('TEXT:', repr(text[:400]))
        words = page.extract_words(keep_blank_chars=True, x_tolerance=4, y_tolerance=4)
        print(f'words: {len(words)}')
        for w in words[:20]:
            print(f'  x0={w["x0"]:.0f} top={w["top"]:.0f} text={repr(w["text"])}')
