# -*- coding: utf-8 -*-
# Find exact lines for coniferous_all_ages

with open('molodniki_extended.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if 'coniferous_all_ages' in line:
        print(f"Line {i+1}: {line.rstrip()}")
