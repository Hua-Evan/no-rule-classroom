# 《無規則教室：班級生存挑戰》

手機可玩的 2D 劇情 RPG：高中班級經營 × 公民社會規範。完整遊戲資料集中在 `js/game-data.js`；Ending 權重集中在 `js/ending-scoring.js`。

## 直接使用

- `index.html`：學生入口；`student.html`：完整遊戲。
- `teacher.html`：教師後台（登入、進度、匿名統計、匿名留言、正式班規）。
- `projection.html`：五幕通關結算（教師最後才公布）。
- 未填 API URL 時，會使用瀏覽器 localStorage 做離線示範，可測試中途續玩。

## 班級設定

在 `js/config.js` 修改 `CLASSES` 與 `STUDENT_COUNT`，班級名稱與人數並未寫死於遊戲流程。

## Google Sheets / Apps Script

1. 建立 Google Sheet，從網址複製試算表 ID。
2. 打開「擴充功能 → Apps Script」，貼上 `apps-script/Code.gs`。
3. 設定 `SHEET_ID` 與不公開的 `TEACHER_TOKEN`。
4. 部署為 Web App，將 Web App URL 填入 `js/config.js` 的 `API_URL`。
5. 教師後台輸入相同 Token；學生不需輸入 Token。

首次寫入會建立 `students`、`settings`、`formal_rules`。一般教師統計不呈現座號與個人答案，投影頁只呈現匿名彙整與教師選取的規範。

## 維護

- 劇情／題目：編輯 `js/game-data.js`，保留各題 `id` 和 `field`。
- Ending 權重：編輯 `js/ending-scoring.js`。
- 正式班規：從教師後台儲存再公布，勿硬寫在學生頁。
- 美術：所有專案插圖放在 `assets/`；角色規格見 `assets/character-bible.md`。新增插圖後，在遊戲資料的 `scene` 填入檔名。

## GitHub Pages

推送到 GitHub 後，在 **Settings → Pages** 選擇分支與根目錄。入口為 `index.html`，教師頁為 `/teacher.html`，投影頁為 `/projection.html`。
