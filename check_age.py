# -*- coding: utf-8 -*-
# Check for coniferous_all_ages in the file

with open('molodniki_extended.py', 'r', encoding='utf-8') as f:
    content = f.read()

print("coniferous_all_ages in content:", 'coniferous_all_ages' in content)
print("avg_coniferous_overall_age in content:", 'avg_coniferous_overall_age' in content)

# Find lines with 'coniferous' and 'age' close to each other
import re
matches = re.findall(r'.{0,100}coniferous.{0,100}age.{0,100}', content)
print(f"\nFound {len(matches)} matches for 'coniferous' near 'age'")
for i, m in enumerate(matches[:5]):
    print(f"\nMatch {i+1}:")
    print(m[:200])
