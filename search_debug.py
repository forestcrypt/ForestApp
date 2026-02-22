# -*- coding: utf-8 -*-
import re
import sys

# Read the file
with open('molodniki_extended.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Search for lines containing "возраст"
lines = content.split('\n')
matching_lines = []
for i, line in enumerate(lines):
    if 'возраст' in line.lower():
        matching_lines.append(f"Line {i+1}: {line[:100]}")

print(f"Found {len(matching_lines)} lines with 'возраст':")
for line in matching_lines[:30]:
    print(line)
