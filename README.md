# Claude Team 帳號登記網站

智慧自動化工程系 · 四智一丙 — 學生填寫 **姓名、學號、學校 E-mail（@gm.student.ncut.edu.tw）**，資料自動寫入 Google 試算表。

- 網站：https://chioujye-huang.github.io/claude-team-signup/
- 登記截止：2026-09-28 23:59（台灣時間）

```
學生瀏覽器 ──► GitHub Pages（index.html）──fetch──► Google Apps Script 網頁應用程式 ──► Google 試算表「登記名單」
```

## 檔案

| 檔案 | 說明 |
|---|---|
| `index.html` | 前端網頁（GitHub Pages） |
| `Code.gs` | 後台程式備份（實際執行於試算表綁定的 Apps Script） |

## 功能

- 登記：姓名、學號、學校信箱（網域固定，輸入學號自動帶入帳號）
- 同一學號重複送出 → 更新資料；同一信箱不可被不同學號使用
- 查詢登記狀態（姓名、信箱遮罩顯示）、取消登記（需學號＋信箱相符）
- 登記人數統計、截止時間、防機器人欄位、多人同時送出的寫入鎖
- 試算表選單「Claude Team → 產生邀請 Email 清單」

## 維護

- 修改設定（截止時間、名額）：前端改 `index.html` 的 `CONFIG`，後台改 Apps Script 的設定區。
- 修改 Apps Script 後需 **部署 → 管理部署作業 → 編輯 → 版本選「新版本」→ 部署**，網址不變。
- 注意：查詢參數不可命名為 `sid`（Apps Script 保留字），目前使用 `stuid`。
- 試算表只有擁有者看得到；網頁只回傳人數與遮罩後的資料。
