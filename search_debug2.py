# -*- coding: utf-8 -*-
# Search for avg_coniferous calculation

with open('molodniki_extended.py', 'r', encoding='utf-8') as f:
    content = f.read()

lines = content.split('\n')

# Search for coniferous_all_ages or avg_coniferous
for i, line in enumerate(lines):
    if 'coniferous_all_ages' in line or 'avg_coniferous_overall_age' in line:
        print(f"Line {i+1}: {line[:120]}")
