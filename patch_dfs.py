with open('app.js', 'r', encoding='utf-8') as f:
    code = f.read()

code = code.replace('for (let attempt = 0; attempt < 20; attempt++) {', 'for (let attempt = 0; attempt < 2000; attempt++) {')

# Find exactly where candidates are filtered
original_filter_1 = 'candidates = candidates.filter(c => !blacklist[`${r}_${c.day}_${c.per}`]);'
new_filter_1 = '''          const pathKey = chosenSlots.map(s => s.day + '-' + s.per).join('|');
          candidates = candidates.filter(c => !blacklist[`${pathKey}_${r}_${c.day}_${c.per}`]);'''

code = code.replace(original_filter_1, new_filter_1)

original_filter_2 = 'evictCandidates = evictCandidates.filter(c => !blacklist[`${r}_${c.day}_${c.per}`]);'
new_filter_2 = '''            const pathKey2 = chosenSlots.map(s => s.day + '-' + s.per).join('|');
            evictCandidates = evictCandidates.filter(c => !blacklist[`${pathKey2}_${r}_${c.day}_${c.per}`]);'''

code = code.replace(original_filter_2, new_filter_2)

original_blacklist = 'blacklist[`${prevR}_${prevChoice.day}_${prevChoice.per}`] = true;'
new_blacklist = '''            const prevPath = chosenSlots.slice(0, prevR).map(s => s.day + '-' + s.per).join('|');
            blacklist[`${prevPath}_${prevR}_${prevChoice.day}_${prevChoice.per}`] = true;'''

code = code.replace(original_blacklist, new_blacklist)

with open('app.js', 'w', encoding='utf-8') as f:
    f.write(code)
print('Applied DFS Patch!')
