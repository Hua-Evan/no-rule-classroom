/* 將 Google 試算表網址中的 ID 貼入；不要把任何帳密或金鑰放到學生端。 */
const SHEET_ID = '請填入試算表 ID';
const SHEET_NAME = 'students';
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
  return {ignored:false};
  } finally {
    lock.releaseLock();
  }
}
