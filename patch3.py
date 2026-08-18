import os

with open('app.js', 'r', encoding='utf-8') as f:
    code = f.read()
    f.seek(0)
    lines = f.readlines()

start_line = -1
end_line = -1
for i, line in enumerate(lines):
    if '// 逐輪（第1節、第2節…）找共同可用時段' in line:
        start_line = i
    if start_line != -1 and '班級: [${roundLsn.map(l=>l.classCode).join' in line:
        end_line = i + 4
        break

if start_line != -1 and end_line != -1:
    original_text = ''.join(lines[start_line:end_line])
    
    with open('patch2.py', 'r', encoding='utf-8') as f:
        patch_code = f.read()
    
    new_code = patch_code.split('new_code = r\'\'\'')[1].split('\'\'\'')[0]
    
    code = code.replace(original_text, new_code)
    with open('app.js', 'w', encoding='utf-8') as f:
        f.write(code)
    print('SUCCESSFULLY PATCHED APP.JS')
else:
    print('Could not find the block boundaries')
