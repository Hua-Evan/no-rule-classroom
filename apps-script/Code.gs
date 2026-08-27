/* 將 Google 試算表網址中的 ID 貼入；不要把任何帳密或金鑰放到學生端。 */
const SHEET_ID = '請填入試算表 ID';
const SHEET_NAME = 'students';
const ANALYSIS_SHEET_NAME = '學生作答分析';
const HEADERS = [
  'student_id','class','seat','avatar_id',
  'prologue_answer','level1_cleared','phone_initial','phone_reason_avoid_self_distraction','phone_reason_avoid_affecting_others','phone_reason_maintain_learning_order','phone_reason_reduce_teacher_management','phone_reason_students_cannot_self_manage','phone_reason_no_rules_needed','phone_event_answer','phone_after_event','phone_changed',
  'level2_cleared','lunch_initial','lunch_branch_answer','lunch_quiet_meaning','lunch_impact_definition','lunch_limit_reasonable','lunch_clarity','lunch_exception_answer','lunch_after_event','lunch_changed',
  'level3_cleared','seat_vote_real','seat_preference','seat_selected_position','seat_selected_type','seat_preference_met','seat_response','level3_scanner_q1','level3_scanner_q2','level3_scanner_q3','level3_scanner_q4','seat_rule_match','majority_rule_answer','seat_final_answer',
  'level4_cleared','paper_cup_action','responsibility_question','recycle_action','reminder_style','power_action','unassigned_task_answer','responsibility_final_answer',
  'final_boss_cleared','final_boss_round1_exec','final_boss_round1_clear','final_boss_round1_fair','final_boss_round1_except','final_boss_round1_community','final_boss_scan_purpose','final_boss_scan_method','final_boss_round2_exec','final_boss_round2_clear','final_boss_round2_fair','final_boss_round2_except','final_boss_round2_community','final_boss_round3_answer','final_boss_round4a_exec','final_boss_round4a_clear','final_boss_round4a_fair','final_boss_round4a_except','final_boss_round4a_community','final_boss_round4b_exec','final_boss_round4b_clear','final_boss_round4b_fair','final_boss_round4b_except','final_boss_round4b_community','final_rule_fragment_individual_needs','final_rule_fragment_actual_impact','final_rule_fragment_rule_purpose','final_rule_fragment_reasonable_exception','final_rule_fragment_majority_support','final_rule_fragment_same_rule_for_all','final_rule_fragment_execution_cost','final_rule_core_1','final_rule_core_2','final_rule_last',
  'ending_type','ending_top1','ending_top2','ending_top3','profile_freedom','profile_learning','profile_order','profile_empathy','profile_community','profile_executability',
  'status_freedom','status_learning','status_order','status_empathy','status_community',
  'started_at','updated_at'
];

function doPost(e) {
  try {
    const request = JSON.parse(e.postData && e.postData.contents || '{}');
    if (request.action !== 'syncStudent') throw new Error('Unsupported action');
    const payload = request.payload || {};
    if (!isUuid(payload.student_id)) throw new Error('Invalid student_id');
    const data = payload.data || {};
    if (data.student_id !== payload.student_id) throw new Error('student_id mismatch');
    const row = pickWhitelisted(data);
    row.student_id = payload.student_id;
    if (!row.updated_at) row.updated_at = new Date().toISOString();
    const result = upsertStudent(row);
    // 分析表僅是教師閱讀用的衍生資料；即使轉換失敗，也不能影響學生端同步。
    if (!result.ignored) {
      try {
        upsertAnalysis(toRecord(result.values));
      } catch (analysisError) {
        console.warn('學生作答分析同步失敗：' + analysisError);
      }
    }
    return json({ok:true,student_id:row.student_id,updated_at:row.updated_at,ignored:result.ignored});
  } catch (error) {
    return json({ok:false,error:String(error)});
  }
}

function doGet() { return json({ok:true}); }
function json(value) { return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON); }
function isUuid(value) { return typeof value === 'string' && /^[-a-zA-Z0-9]{16,80}$/.test(value); }
function pickWhitelisted(data) { return HEADERS.reduce((row,key) => { if (Object.prototype.hasOwnProperty.call(data,key)) row[key] = data[key]; return row; }, {}); }
function sheet() {
  const target = SpreadsheetApp.openById(SHEET_ID);
  let page = target.getSheetByName(SHEET_NAME);
  if (!page) { page = target.insertSheet(SHEET_NAME); page.appendRow(HEADERS); }
  const current = page.getRange(1,1,1,Math.max(page.getLastColumn(),1)).getValues()[0];
  const missing = HEADERS.filter(header => !current.includes(header));
  if (missing.length) page.getRange(1,current.length + 1,1,missing.length).setValues([missing]);
  return page;
}
function upsertStudent(row) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
  const page = sheet();
  const column = HEADERS.indexOf('student_id') + 1;
  const lastRow = page.getLastRow();
  const ids = lastRow > 1 ? page.getRange(2,column,lastRow - 1,1).getValues().flat() : [];
  const index = ids.findIndex(id => String(id) === row.student_id);
  if (index >= 0) {
    const previousUpdatedAt = page.getRange(index + 2,HEADERS.indexOf('updated_at') + 1).getValue();
    if (previousUpdatedAt && new Date(previousUpdatedAt).getTime() > new Date(row.updated_at).getTime()) return {ignored:true};
  }
  const values = HEADERS.map(header => row[header] === undefined || row[header] === null ? '' : row[header]);
  if (index >= 0) page.getRange(index + 2,1,1,HEADERS.length).setValues([values]);
  else page.appendRow(values);
  return {ignored:false,values:values};
  } finally {
    lock.releaseLock();
  }
}

const ANALYSIS_COLUMNS = [
  ['class','班級'], ['seat','座號'], ['avatar_id','遊戲角色'], ['ending_type','最終規範角色'],
  ['level1_cleared','LEVEL 1 完成'], ['level2_cleared','LEVEL 2 完成'], ['level3_cleared','LEVEL 3 完成'], ['level4_cleared','LEVEL 4 完成'], ['final_boss_cleared','FINAL BOSS 完成'], ['started_at','開始時間'], ['updated_at','最後同步時間'], ['student_id','學生識別碼'],
  ['prologue_answer','序章｜你覺得沒有班級規範的班級會怎麼樣？'],
  ['phone_initial','L1｜你認為哪一種手機管理方式最好？【初始立場】'],
  ['phone_reason_avoid_self_distraction','L1｜手機共同規範理由｜避免自己分心'], ['phone_reason_avoid_affecting_others','L1｜手機共同規範理由｜避免影響其他同學'], ['phone_reason_maintain_learning_order','L1｜手機共同規範理由｜維持共同學習秩序'], ['phone_reason_reduce_teacher_management','L1｜手機共同規範理由｜降低老師提醒與管理時間'], ['phone_reason_students_cannot_self_manage','L1｜手機共同規範理由｜學生無法自我管理'], ['phone_reason_no_rules_needed','L1｜手機共同規範理由｜根本不需要規範'],
  ['phone_event_answer','L1｜「每節上課集中、下課歸還」這個制度，最大的問題可能是什麼？'], ['phone_after_event','L1｜經過手機事件後，你認為哪一種手機管理方式最好？【最終立場】'], ['phone_changed','L1｜手機規範前後立場是否改變？'],
  ['lunch_initial','L2｜你認為哪種午休規範比較合適？【初始立場】'], ['lunch_branch_answer','L2｜依你的午休初始選擇，面對情境後的判斷'], ['lunch_quiet_meaning','L2｜「保持安靜」到底是不能講話，還是不能製造干擾？'], ['lunch_impact_definition','L2｜請寫下你認為「較不影響他人」的程度或做法【學生自由回答】'], ['lunch_limit_reasonable','L2｜你認為「所有人都必須趴下」的規範是否限制太多？'], ['lunch_clarity','L2｜一條規範非常明確，所以它就一定是好規範嗎？'], ['lunch_exception_answer','L2｜面對特殊或必要情況，午休規範應該怎麼處理？'], ['lunch_after_event','L2｜經過午休事件後，你會怎麼設計午休規範？【最終立場】'], ['lunch_changed','L2｜午休規範前後立場是否改變？'],
  ['seat_vote_real','L3｜你是否支持「依段考總成績排名，從第一名開始依序自由選擇座位」？'], ['seat_preference','L3｜選座位時，你最在意什麼？'], ['seat_selected_position','L3｜最後選擇的座位'], ['seat_selected_type','L3｜最後選擇的座位類型'], ['seat_preference_met','L3｜最後座位是否符合原本偏好？'], ['seat_response','L3｜座位制度出現問題後，你認為應該怎麼處理？'], ['level3_scanner_q1','L3 規則掃描｜這條規則是否清楚告訴大家「誰先選、誰後選」？'], ['level3_scanner_q2','L3 規則掃描｜每位同學是否都使用「段考成績排名」這個相同標準？'], ['level3_scanner_q3','L3 規則掃描｜這條規則是否經過全班表決？'], ['level3_scanner_q4','L3 規則掃描｜「段考成績」和「安排適合的座位」之間是否一定有直接關係？'], ['seat_rule_match','L3｜是否完成「規則目的 × 標準」配對'], ['majority_rule_answer','L3｜一項規則只要「多數人支持＋每個人適用同一標準」就一定公平合理嗎？'], ['seat_final_answer','L3｜如果 35 人中有 34 人支持、只有 1 人反對，這項規則就一定合理嗎？'],
  ['paper_cup_action','L4｜看到不是自己的飲料杯時，你會怎麼做？'], ['responsibility_question','L4｜如果每個人都完成自己的分工，班級就一定會維持得很好嗎？'], ['recycle_action','L4｜看到回收物被丟進一般垃圾時，你會怎麼做？'], ['reminder_style','L4｜面對回收分類錯誤，你會採取什麼提醒語氣／後續做法？'], ['power_action','L4｜發現不是自己負責的電燈或電扇還開著，你會怎麼處理？'], ['unassigned_task_answer','L4｜共同環境出現沒有明確分配給任何人的事情，比較適合怎麼處理？'], ['responsibility_final_answer','L4｜如果每個人都完成自己被分配的工作，是否就足以維持共同環境？'],
  ['final_boss_round1_exec','FINAL R1｜選擇技能：可執行性'], ['final_boss_round1_clear','FINAL R1｜選擇技能：明確性'], ['final_boss_round1_fair','FINAL R1｜選擇技能：規範合理性'], ['final_boss_round1_except','FINAL R1｜選擇技能：合理例外'], ['final_boss_round1_community','FINAL R1｜選擇技能：共同體責任'],
  ['final_boss_scan_purpose','FINAL｜「作業全準時可優先選座位」：你判斷這條規則的目的為何？'], ['final_boss_scan_method','FINAL｜「作業全準時可優先選座位」：你判斷這條規則使用的手段為何？'],
  ['final_boss_round2_exec','FINAL R2｜選擇技能：可執行性'], ['final_boss_round2_clear','FINAL R2｜選擇技能：明確性'], ['final_boss_round2_fair','FINAL R2｜選擇技能：規範合理性'], ['final_boss_round2_except','FINAL R2｜選擇技能：合理例外'], ['final_boss_round2_community','FINAL R2｜選擇技能：共同體責任'],
  ['final_boss_round3_answer','FINAL｜「生日當天可免除所有班級工作」，你會怎麼修改這條規則？'],
  ['final_boss_round4a_exec','FINAL R4A｜選擇技能：可執行性'], ['final_boss_round4a_clear','FINAL R4A｜選擇技能：明確性'], ['final_boss_round4a_fair','FINAL R4A｜選擇技能：規範合理性'], ['final_boss_round4a_except','FINAL R4A｜選擇技能：合理例外'], ['final_boss_round4a_community','FINAL R4A｜選擇技能：共同體責任'],
  ['final_boss_round4b_exec','FINAL R4B｜選擇技能：可執行性'], ['final_boss_round4b_clear','FINAL R4B｜選擇技能：明確性'], ['final_boss_round4b_fair','FINAL R4B｜選擇技能：規範合理性'], ['final_boss_round4b_except','FINAL R4B｜選擇技能：合理例外'], ['final_boss_round4b_community','FINAL R4B｜選擇技能：共同體責任'],
  ['final_rule_fragment_individual_needs','FINAL RULE｜曾分析：個別需求'], ['final_rule_fragment_actual_impact','FINAL RULE｜曾分析：實際影響'], ['final_rule_fragment_rule_purpose','FINAL RULE｜曾分析：規則目的'], ['final_rule_fragment_reasonable_exception','FINAL RULE｜曾分析：合理例外'], ['final_rule_fragment_majority_support','FINAL RULE｜曾分析：多數支持'], ['final_rule_fragment_same_rule_for_all','FINAL RULE｜曾分析：人人相同'], ['final_rule_fragment_execution_cost','FINAL RULE｜曾分析：執行成本'], ['final_rule_core_1','FINAL RULE｜最終核心漏洞 1'], ['final_rule_core_2','FINAL RULE｜最終核心漏洞 2'], ['final_rule_last','FINAL RULE｜最後一次分析的碎片'],
  ['ending_top1','規範雷達第 1 名'], ['ending_top2','規範雷達第 2 名'], ['ending_top3','規範雷達第 3 名'],
  ['profile_freedom','規範傾向｜個人自由'], ['profile_learning','規範傾向｜學習效率'], ['profile_order','規範傾向｜班級秩序'], ['profile_empathy','規範傾向｜同學感受'], ['profile_community','規範傾向｜共同責任'], ['profile_executability','規範傾向｜可執行性'],
  ['status_freedom','CLASS STATUS｜自由度'], ['status_learning','CLASS STATUS｜學習效率'], ['status_order','CLASS STATUS｜班級秩序'], ['status_empathy','CLASS STATUS｜同學感受'], ['status_community','CLASS STATUS｜共同責任']
];

const LABELS = {
  seat_preference: {friend:'跟熟識朋友坐附近',focus:'坐在較能專心上課的位置',far:'想坐離老師遠一點',comfort:'選最喜歡、最舒服的位置',none:'沒有特別偏好'},
  seat_selected_type: {front:'前排座位',back_window:'靠窗後排',back_friend:'朋友附近',window:'靠窗座位',aisle:'走道座位',middle:'中間座位'},
  recycle_action: {speak:'開口提醒',self:'不說話，自己處理',ignore:'當作沒看到'},
  lunch_initial: {'午休時間所有人都必須趴下休息':'全班趴下休息','不一定要睡，但必須保持安靜':'不睡也須安靜','可以做自己的事，只要不要影響別人':'可做事、不影響他人','午休完全自由':'午休完全自由'},
  lunch_after_event: {'所有人都趴下':'全班趴下休息','可做自己的事，但清楚列出安靜與不干擾的做法':'可做事，但明訂安靜與不干擾','午休完全自由':'午休完全自由','由老師隨時決定':'老師依情況決定'},
  final_boss_round3_answer: {'取消生日福利':'取消生日福利','安排代理':'生日可免部分工作，但須事先安排代理','其他人自己處理':'生日同學全免、其他人自行處理','完全不調整':'生日仍完成全部工作'},
  ending: {freedom:'自由守護者',learning:'學習領航員',order:'秩序守門人',empathy:'共感協調師',community:'共好行動者',executability:'制度工程師'},
  fragment: {'👥 個別需求':'個別需求','👀 實際影響':'實際影響','🎯 規則目的':'規則目的','🧩 合理例外':'合理例外','🗳️ 多數支持':'多數支持','📜 人人相同':'人人相同','🔧 執行成本':'執行成本'}
};

function toRecord(values) { return HEADERS.reduce((record,key,index) => (record[key] = values[index], record), {}); }
function isBlank(value) { return value === '' || value === null || value === undefined; }
function isTrue(value) { return value === true || value === 1 || value === '1' || value === 'TRUE' || value === 'true'; }
function completed(value) { return isTrue(value) ? '已完成' : '未完成'; }
function changed(value) { return isBlank(value) ? '' : isTrue(value) ? '有改變' : '未改變'; }
function checked(value) { return isTrue(value) ? '✓' : ''; }
function displayPosition(value) { return isBlank(value) ? '' : '第 ' + String(Number(value) + 1).padStart(2,'0') + ' 號座位'; }
function displayPreferenceMet(value) { return ({true:'符合偏好',partial:'部分符合偏好',false:'不符合偏好'})[String(value)] || ''; }
function displayLabel(value, labels) { return isBlank(value) ? '' : (labels[value] || value); }
function displayValue(key, value) {
  if (key === 'seat_selected_position') return displayPosition(value);
  if (key === 'seat_preference_met') return displayPreferenceMet(value);
  if (key === 'seat_preference' || key === 'seat_selected_type' || key === 'recycle_action' || key === 'lunch_initial' || key === 'lunch_after_event' || key === 'final_boss_round3_answer') return displayLabel(value,LABELS[key]);
  if (key === 'ending_type' || /^ending_top[123]$/.test(key)) return displayLabel(value,LABELS.ending);
  if (/^final_rule_(core_[12]|last)$/.test(key)) return displayLabel(value,LABELS.fragment);
  if (/^final_boss_round[124][ab]?_(exec|clear|fair|except|community)$/.test(key) || /^final_rule_fragment_/.test(key) || /^phone_reason_/.test(key)) return checked(value);
  if (/_cleared$/.test(key) || key === 'seat_rule_match') return completed(value);
  if (key === 'phone_changed' || key === 'lunch_changed') return changed(value);
  return isBlank(value) ? '' : value;
}

function analysisSheet() {
  const target = SpreadsheetApp.openById(SHEET_ID);
  let page = target.getSheetByName(ANALYSIS_SHEET_NAME);
  const created = !page;
  if (created) page = target.insertSheet(ANALYSIS_SHEET_NAME);
  const headers = ANALYSIS_COLUMNS.map(column => column[1]);
  const current = page.getRange(1,1,1,headers.length).getValues()[0];
  const headerChanged = created || headers.some((header,index) => current[index] !== header);

  // 舊版分析表的欄位順序不同；直接由 students 原始資料重建，避免只換 header 造成資料錯位。
  if (headerChanged && !created) {
    const records = studentRecords();
    const oldRows = page.getLastRow() - 1;
    if (oldRows > 0) page.getRange(2,1,oldRows,headers.length).clearContent();
    if (records.length) {
      page.getRange(2,1,records.length,headers.length)
        .setValues(records.map(record => analysisValues(record)));
    }
  }

  if (headerChanged) page.getRange(1,1,1,headers.length).setValues([headers]);
  if (created || headerChanged) formatAnalysisSheet(page,headers.length);
  return page;
}

function studentRecords() {
  const source = sheet();
  const lastRow = source.getLastRow();
  if (lastRow <= 1) return [];
  return source.getRange(2,1,lastRow - 1,HEADERS.length).getValues().map(toRecord);
}

function formatAnalysisSheet(page,width) {
  page.setFrozenRows(1);
  page.setRowHeight(1,72);
  page.getRange(1,1,1,width).setWrap(true);
  page.setColumnWidths(1,width,150);

  const basicWidths = {class:100,seat:70,avatar_id:100,ending_type:120,student_id:190};
  Object.keys(basicWidths).forEach(key => {
    const column = ANALYSIS_COLUMNS.findIndex(item => item[0] === key) + 1;
    page.setColumnWidth(column,basicWidths[key]);
  });

  const textColumn = ANALYSIS_COLUMNS.findIndex(column => column[0] === 'lunch_impact_definition') + 1;
  page.setColumnWidth(textColumn,320);
  page.getRange(2,textColumn,Math.max(page.getMaxRows() - 1,1),1).setWrap(true);

  const filter = page.getFilter();
  if (filter) filter.remove();
  page.getRange(1,1,Math.max(page.getLastRow(),2),width).createFilter();
}

function analysisValues(record) {
  return ANALYSIS_COLUMNS.map(column => displayValue(column[0],record[column[0]]));
}

function upsertAnalysis(record) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const page = analysisSheet();
    const idColumn = ANALYSIS_COLUMNS.findIndex(column => column[0] === 'student_id') + 1;
    const lastRow = page.getLastRow();
    const ids = lastRow > 1 ? page.getRange(2,idColumn,lastRow - 1,1).getValues().flat() : [];
    const index = ids.findIndex(id => String(id) === String(record.student_id));
    const values = analysisValues(record);
    if (index >= 0) page.getRange(index + 2,1,1,values.length).setValues([values]);
    else page.appendRow(values);
  } finally {
    lock.releaseLock();
  }
}
