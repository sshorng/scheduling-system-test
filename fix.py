import re

with open('app.js', 'r', encoding='utf-8') as f:
    code = f.read()

# We need to find the block starting with '// 導入回溯機制（Conflict-Directed Backjumping）' and ending right before 'function canPlaceClassSubjectOnDay'
pattern = re.compile(r'(\s*// 導入回溯機制（Conflict-Directed Backjumping）.*?)(?=\s*function canPlaceClassSubjectOnDay)', re.DOTALL)
match = pattern.search(code)

if match:
    original_broken_text = match.group(1)
    
    new_code = r'''
      // 導入回溯機制（Conflict-Directed Backjumping）
      let bindGroupSuccess = false;
      const blacklist = {}; // 記錄 {round_day_period: true} 避免重蹈覆轍

      for (let attempt = 0; attempt < 20; attempt++) {
        const backupSchedule = [...localSchedule];
        const backupRequeued = [...requeuedLessons];
        const backupScheduledFlag = [...scheduledFlag];
        const backupSuccessCount = successCount;
        
        let attemptFailedAtRound = -1;
        const chosenSlots = [];

        // 逐輪（第1節、第2節…）找共同可用時段
        for (let r = 0; r < maxPerCls; r++) {
          updateProgress(`綁班群組配課中（${totalGroups} 群組，第 ${groupIndex + 1} 組，第 ${r + 1} 輪，嘗試 ${attempt + 1}）…`);
          await yieldToUI();
          const roundIdx = [];
          clsCodes.forEach(cc => {
            if (r < byCls[cc].length) roundIdx.push(byCls[cc][r]);
          });
          if (roundIdx.length < 2) break;
          const roundLsn = roundIdx.map(i => pendingLessons[i]);
          const repeatedTeachers = getRepeatedTeacherCodes(roundLsn);
          if (repeatedTeachers.length) { console.warn('綁班組內教師衝堂，取消此組合：'+repeatedTeachers.join('、')); continue; }
          
          let candidates = [];
          const roundLookup = buildScheduleLookup(localSchedule);
          let slotChecks = 0;
          for (let day = 1; day <= 5; day++) {
            for (let per = autoStartPeriod; per <= autoEndPeriod; per++) {
              slotChecks++;
              if (slotChecks % BIND_GROUP_SLOT_YIELD_INTERVAL === 0) await yieldToUI();
              if (roundLsn.every(l => isSlotValid(l.classCode, l.subjectCode, autoTeacherInput(l), day, per, localSchedule, roundLookup))) {
                const allBoundTeachersValid = optTeacherConsec || roundLsn.every(l => {
                  if (!l.teacherCode) return true;
                  const tObj = idx.teacherByCode[l.teacherCode];
                  const maxC = tObj ? parseInt(tObj['最大連堂節數'] || '2', 10) : 2;
                  return countConsecutiveInLocal(localSchedule, l.teacherCode, day, per, roundLookup) <= maxC;
                });
                if (allBoundTeachersValid) {
                  let score = 0;
                  roundLsn.forEach(l => { score += evaluateSlotScore(l, day, per); });
                  candidates.push({ day, per, score });
                }
              }
            }
          }
          
          // 過濾黑名單
          candidates = candidates.filter(c => !blacklist[`${r}_${c.day}_${c.per}`]);

          if (candidates.length > 0) {
            candidates.sort((a, b) => b.score - a.score);
            const best = candidates[0];
            chosenSlots.push(best);
            roundLsn.forEach((l, gi) => {
              const id = 'AUTO_' + Date.now() + '_' + Math.random().toString(36).substr(2, 4) + '_' + r + '_' + gi;
              localSchedule.push({ '課表ID': id, '班級代碼': String(l.classCode), '星期': best.day, '節次': best.per, '科目代碼': String(l.subjectCode), '教師姓名': autoTeacherValue(l), '課堂屬性': l.isVirtual ? '抽離' : '一般', '是否鎖定': 'FALSE', '是否預排': 'FALSE', '__isBindGroup': true });
              successCount++;
              scheduledFlag[roundIdx[gi]] = true;
            });
            console.log(`    ✅ 第 ${r+1} 輪 → 星期${best.day} 第${best.per}節 (${clsCodes.length} 班同格)`);
          } else {
            console.log(`    ⚡ 第 ${r+1} 輪 → 無自然共同時段，啟動讓位搜尋…`);
            let evictCandidates = [];
            for (let day = 1; day <= 5; day++) {
              for (let per = autoStartPeriod; per <= autoEndPeriod; per++) {
                const teacherHardBlock = roundLsn.some(l => {
                  const tcCodes = getAutoTeacherCodes(autoTeacherInput(l));
                  return tcCodes.some(tc => {
                    const canonical = canonicalAutoTeacherCode(tc);
                    if (!canonical) return false;
                    if (idx.blockSet.has(canonical + '|' + day + '|' + per)) return true;
                    return localSchedule.some(s =>
                      s !== undefined &&
                      resolveAutoTeacherCodes(s).includes(canonical) &&
                      parseInt(s['星期'], 10) === day &&
                      parseInt(s['節次'], 10) === per &&
                      isFrozenAutoEntry(s)
                    );
                  });
                });
                if (teacherHardBlock) continue;

                const mustRuleBlock = roundLsn.some(l => {
                  const applicableRules = getApplicableRules(l.subjectCode, l.classCode);
                  const mustRules = applicableRules.filter(entry => entry.rule['規則類型'] === '必排');
                  return mustRules.length > 0 && !mustRules.some(entry => entry.slots.some(slot => slot.day === day && slot.period === per));
                });
                if (mustRuleBlock) continue;

                const toEvict = [];
                let canEvict = true;
                for (const l of roundLsn) {
                  const existing = localSchedule.find(s =>
                    String(s['班級代碼']) === String(l.classCode) &&
                    parseInt(s['星期'], 10) === day &&
                    parseInt(s['節次'], 10) === per
                  );
                  if (existing) {
                    if (isFrozenAutoEntry(existing) || existing.__isBindGroup) { canEvict = false; break; }
                    toEvict.push(existing);
                  }
                }
                if (!canEvict) continue;

                const evictedIndices = toEvict.map(ev => localSchedule.indexOf(ev));
                evictedIndices.slice().reverse().forEach(i => { if (i >= 0) localSchedule.splice(i, 1); });
                const evictLookup = buildScheduleLookup(localSchedule);
                const allValid = roundLsn.every(l => isSlotValid(l.classCode, l.subjectCode, autoTeacherInput(l), day, per, localSchedule, evictLookup));
                toEvict.forEach(ev => localSchedule.push(ev));

                if (allValid) {
                  const score = roundLsn.reduce((sum, l) => sum + evaluateSlotScore(l, day, per), 0);
                  evictCandidates.push({ day, per, score, toEvict: [...toEvict] });
                }
              }
            }

            evictCandidates = evictCandidates.filter(c => !blacklist[`${r}_${c.day}_${c.per}`]);

            if (evictCandidates.length > 0) {
              evictCandidates.sort((a, b) => b.score - a.score);
              const best = evictCandidates[0];
              chosenSlots.push(best);
              best.toEvict.forEach(ev => {
                const i = localSchedule.indexOf(ev);
                if (i >= 0) localSchedule.splice(i, 1);
                requeuedLessons.push(makeAutoLessonFromScheduleEntry(ev, 'bind_evict'));
                console.log(`    ↩️  讓位：${ev['班級代碼']} ${ev['科目代碼']} 星期${ev['星期']}第${ev['節次']}節 → 放回重排`);
              });
              const evictLookup2 = buildScheduleLookup(localSchedule);
              roundLsn.forEach((l, gi) => {
                const id = 'AUTO_' + Date.now() + '_' + Math.random().toString(36).substr(2, 4) + '_EV' + r + '_' + gi;
                localSchedule.push({ '課表ID': id, '班級代碼': String(l.classCode), '星期': best.day, '節次': best.per, '科目代碼': String(l.subjectCode), '教師姓名': autoTeacherValue(l), '課堂屬性': l.isVirtual ? '抽離' : '一般', '是否鎖定': 'FALSE', '是否預排': 'FALSE', '__isBindGroup': true });
                successCount++;
                scheduledFlag[roundIdx[gi]] = true;
              });
              console.log(`    ✅ 第 ${r+1} 輪（讓位）→ 星期${best.day} 第${best.per}節 (${clsCodes.length} 班同格, 讓位 ${best.toEvict.length} 節)`);
            } else {
              attemptFailedAtRound = r;
              let cntTeacher = 0, cntMust = 0, cntFrozen = 0, cntValid = 0;
              for (let day = 1; day <= 5; day++) {
                for (let per = autoStartPeriod; per <= autoEndPeriod; per++) {
                  const tblk = roundLsn.some(l => {
                    const tcCodes = getAutoTeacherCodes(autoTeacherInput(l));
                    return tcCodes.some(tc => {
                      const canonical = canonicalAutoTeacherCode(tc);
                      if (!canonical) return false;
                      if (idx.blockSet.has(canonical + '|' + day + '|' + per)) return true;
                      return localSchedule.some(s => s !== undefined && resolveAutoTeacherCodes(s).includes(canonical) && parseInt(s['星期'],10)===day && parseInt(s['節次'],10)===per && isFrozenAutoEntry(s));
                    });
                  });
                  if (tblk) { cntTeacher++; continue; }
                  const mblk = roundLsn.some(l => {
                    const ar = getApplicableRules(l.subjectCode, l.classCode);
                    const mr = ar.filter(e => e.rule['規則類型']==='必排');
                    return mr.length > 0 && !mr.some(e => e.slots.some(s => s.day===day && s.period===per));
                  });
                  if (mblk) { cntMust++; continue; }
                  const fblk = roundLsn.some(l => {
                    const ex = localSchedule.find(s => String(s['班級代碼'])===String(l.classCode) && parseInt(s['星期'],10)===day && parseInt(s['節次'],10)===per);
                    return ex && isFrozenAutoEntry(ex);
                  });
                  if (fblk) { cntFrozen++; continue; }
                  cntValid++;
                }
              }
              console.warn(`    ❌ 第 ${r+1} 輪 → 讓位搜尋失敗：教師硬衝突=${cntTeacher} 格, 必排規則外=${cntMust} 格, 凍結格=${cntFrozen} 格, 可嘗試格=${cntValid} 格`);
              console.warn(`       班級: [${roundLsn.map(l=>l.classCode).join(', ')}], 科目: ${roundLsn[0]?.subjectCode}`);
              break;
            }
          }
        } // end round loop

        if (attemptFailedAtRound === -1) {
          bindGroupSuccess = true;
          break; // 排課成功
        } else {
          localSchedule.length = 0; localSchedule.push(...backupSchedule);
          requeuedLessons.length = 0; requeuedLessons.push(...backupRequeued);
          for(let i=0; i<scheduledFlag.length; i++) scheduledFlag[i] = backupScheduledFlag[i];
          successCount = backupSuccessCount;

          if (attemptFailedAtRound > 0) {
            const prevR = attemptFailedAtRound - 1;
            const prevChoice = chosenSlots[prevR];
            blacklist[`${prevR}_${prevChoice.day}_${prevChoice.per}`] = true;
            console.log(`    ⚠️ 第 ${attemptFailedAtRound+1} 輪排入失敗。觸發回溯機制：將上一輪(第${prevR+1}輪)的 星期${prevChoice.day}第${prevChoice.per}節 加入黑名單，重新排課 (嘗試第 ${attempt+2} 次)`);
          } else {
            console.log(`    ❌ 第 1 輪就失敗，無法排入，放棄此綁班組合。`);
            break;
          }
        }
      }
'''
    code = code.replace(original_broken_text, new_code)
    with open('app.js', 'w', encoding='utf-8') as f:
        f.write(code)
    print("Fixed!")
else:
    print("Could not find broken block")
