# -*- coding: utf-8 -*-
# Show lines around 3795-3810

with open('molodniki_extended.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(3794, 3820):
    print(f"{i+1}: {lines[i].rstrip()}")
