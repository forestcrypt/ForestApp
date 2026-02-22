# -*- coding: utf-8 -*-
# Check lines around 3797-3802

with open('molodniki_extended.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(3796, 3805):
    print(f"{i+1}: {lines[i].rstrip()}")
