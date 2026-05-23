# 產險數位特助專案 Git 倉儲建立與網頁優化實作計畫

本計畫旨在協助您將現有的產險數位特助系統程式碼進行專案化整理，建立本機 Git 倉儲以利後續上傳至 GitHub，並針對前端 `Index.html` 網頁的顯示介面提出具體的優化調整方案。

## 使用者審查項目

> [!IMPORTANT]
> 1. **專案目錄結構**：計畫將現有程式碼檔案重命名為標準的 `Code.gs` 與 `Index.html`，並儲存於本機工作區 `C:\Users\ASUS\.gemini\antigravity\scratch\insurance-digital-assistant` 目錄中。
> 2. **GitHub 連接**：本計畫會引導您在本機完成 `git init` 與首版提交。後續將程式碼推送到 GitHub 遠端倉儲需要您提供您的 GitHub 帳號權限或設定。
> 3. **前端介面優化**：本計畫預計對 `Index.html` 進行 UI / UX 與視覺美化（如導入現代色彩、優化資訊卡片層級、強化搜尋列視覺引導等），程式邏輯保持不變。

## 預定修改內容

---

### 1. 專案目錄與 Git 倉儲建立

#### [NEW] [Code.gs](file:///C:/Users/ASUS/.gemini/antigravity/scratch/insurance-digital-assistant/Code.gs)
- 複製自 [產險數位特助系統 - 1150523A.txt](file:///D:/AI%E4%BA%BA%E5%B7%A5%E6%99%BA%E6%85%A7/%E8%A4%87%E8%A3%BD%E7%94%A2%E9%9A%AA%E8%B3%87%E6%96%99/%E7%94%A2%E9%9A%AA%E6%95%B8%E4%BD%8D%E7%89%B9%E5%8A%A9%E7%B3%BB%E7%B5%B1%20-%201150523A.txt) 內容，作為後端邏輯的主要檔案。

#### [NEW] [Index.html](file:///C:/Users/ASUS/.gemini/antigravity/scratch/insurance-digital-assistant/Index.html)
- 複製自 [產險數位特助index1150523A.txt](file:///D:/AI%E4%BA%BA%E5%B7%A5%E6%99%BA%E6%85%A7/%E8%A4%87%E8%A3%BD%E7%94%A2%E9%9A%AA%E8%B3%87%E6%96%99/%E7%94%A2%E6%95%B8%E4%BD%8D%E7%89%B9%E5%8A%A9index1150523A.txt) 內容，並套用以下顯示優化調整。

#### [NEW] [README.md](file:///C:/Users/ASUS/.gemini/antigravity/scratch/insurance-digital-assistant/README.md)
- 新增專案說明文件，說明資料庫結構（以 Google Sheets 為資料庫）、環境設定與部署方式。

#### [NEW] [.gitignore](file:///C:/Users/ASUS/.gemini/antigravity/scratch/insurance-digital-assistant/.gitignore)
- 排除不需要被 Git 追蹤的暫存檔案與系統設定檔。

---

### 2. 前端 `Index.html` 顯示優化調整方案

預計針對前端介面進行以下調整以提升視覺美感與使用者體驗（ UX ）：

#### A. 視覺美化與設計系統
- **調色盤調整**：將原本較為飽和的顏色調整為更具質感的和諧色彩，例如使用 `#0083B0`（產險藍）與 `#FF8A65`（旅平橙），並增加微光澤的背景處理。
- **現代字型與間距**：引進 Google Fonts 的 `Inter` 字型，優化卡片間的縱向間距與內邊距，提供更具呼吸感的排版。
- **卡片陰影與漸層**：卡片邊框加上細微的漸層與柔和陰影（ `box-shadow` ），當滑鼠懸停時加上微幅上浮的動畫效果。

#### B. 資訊層級與排版優化
- **到期提醒卡片**：將緊急到期（ 5 天內）與預警（ 45 至 51 天）的卡片標題及背景調整為高雅的漸層警示色（如淺玫瑰紅漸層、淺琥珀橘漸層），加強視覺對比。
- **搜尋欄優化**：搜尋框改為圓角並加入聚焦時的微動畫框線，並將下方的「混合輸入提示」改為更具質感的標籤樣式。
- **資料欄位網格**：將卡片內部的資訊網格調整為清晰的左右對齊結構，避免欄位資訊過密或重疊。

---

## 驗證計畫

### 自動化與本機驗證
1. 檢查檔案是否已成功建立於目標目錄，且程式碼完整性無損。
2. 於本機執行 `git status` 與 `git log`，確保 Git 倉儲已正確初始化並完成首次提交。
3. 在瀏覽器中本機載入 `Index.html`，驗證 UI 渲染效果、排版響應式設計（ Mobile 與 Desktop ）、以及 CSS 樣式之正確性。

### 手動驗證與部署
1. 指引使用者如何在本機將專案與遠端 GitHub 倉儲關聯並完成推代碼流程。
2. 驗證 Google Apps Script 的專案與 Google Sheets 資料庫的讀寫連結是否運作正常。
