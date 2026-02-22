# -*- coding: utf-8 -*-
# Search for text_parts usage around line 3800-3850

with open('molodniki_extended.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

# Print lines around 3790-3830
for i in range(3785, 3835):
    print(f"{i+1}: {lines[i][:120]}")
