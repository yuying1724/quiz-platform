# 測驗平台 quiz-platform

通用題庫測驗系統。題庫與作答紀錄存在 Google Sheets，後端是 Google Apps Script 的 JSON API，前端是獨立的 PWA（可加到手機主畫面）。目前科目：LOMA 301。

> 題庫內容（LOMA 教材）只放在私人的 Google Sheet，不放進這個 repo。

## 結構

```
apps-script/   Apps Script 後端（Code.gs、appsscript.json、.clasp.json）
web/           PWA 前端（index.html、manifest.json、sw.js、icons/）
```

## 資料

Google Sheet「測驗系統_題庫」，兩張工作表：

- `題庫`：題號, 科目, 來源, 章節, 學習目標, 題目, 選項1–5, 答案, 教材頁碼, 啟用
- `作答紀錄`：時間, 場次, 科目, 章節, 題號, 來源, 我的答案, 正確答案, 是否正確, 模式

題號規則：`C{章}-{序}` 練習題、`M-{序}` 模擬試題。新增題目直接在 Sheet 加列，不用改程式。

## 後端部署（Apps Script）

1. 試算表 → 擴充功能 → Apps Script，貼上 `apps-script/Code.gs`（或用 `clasp push`）
2. 執行一次 `setup()`（建立工作表）與 `generateToken()`（產生 API_TOKEN，到執行記錄複製）
3. 部署 → 新增部署作業 → 網頁應用程式 → 執行身分「我」、存取「**任何人**」→ 複製網址

改程式後：部署 → 管理部署作業 → 編輯 → 版本選「新版本」→ 部署。

### clasp（選用）

```
npm i -g @google/clasp
clasp login
cd apps-script
clasp push
```

## 前端部署（GitHub Pages）

Settings → Pages → Source 選 branch `main`、資料夾 `/web`（或 `/docs`，視設定）。
網址會是 `https://<帳號>.github.io/quiz-platform/`。

第一次打開前端會要求填入 API 網址與 Token，存在該裝置的 localStorage。

## API

| 方法 | 參數 | 說明 |
|---|---|---|
| GET | `action=meta` | 科目、章節、題數、正確率 |
| GET | `action=questions&subject=&chapters=1,2&source=all&onlyWrong=0` | 出題 |
| GET | `action=history&subject=` | 最近 20 場成績 |
| POST | `{action:'answer', rec}` | 存一題 |
| POST | `{action:'answers', recs:[]}` | 存多題（測驗模式交卷） |

所有請求都要帶 `token`。POST 用 `text/plain` 避免 CORS 預檢。
