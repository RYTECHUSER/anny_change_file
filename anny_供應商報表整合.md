# anny_供應商報表整合 — 操作手冊與對話紀錄

**建立日期：** 2026-05-01  
**GitHub Repo：** https://github.com/RYTECHUSER/anny_change_file  
**負責人：** RYTECHUSER（claude_hans1@hans-tech.asia）

---

## 一、專案說明

每個月從 Hanstrong 流量系統下載各線路的原始 Excel 數據，提取 **95th percentile（95 值）** 後，自動比對填入月報表，再上傳至 GitHub 備存。

整個流程完全用 **Node.js 腳本** 處理，不需要 Python。

---

## 二、檔案說明

### 固定檔案（每月共用）

| 檔案名稱 | 路徑 | 說明 |
|---------|------|------|
| `1. 每月(供應商)頻寬統計(ZL.VOCOM.CMI)_HANS版本.xlsx` | `my-agent/200_Reference/writing-samples/files/` | 月報表原始模板，含 Zenlayer / 遠傳(CMI) / VOCOM 三段，每月固定更新這份 |

### 每月產生的檔案（命名規則加年月前綴）

| 檔案名稱範例 | 說明 |
|------------|------|
| `2026-04_Hanstrong_Traffic_hans_95.xlsx` | 當月從系統下載的原始流量數據（17 個 Sheet + 總表） |
| `2026-04_new_每月(供應商)頻寬統計(ZL.VOCOM.CMI)_HANS版本.xlsx` | 當月產出的月報表（已填入 95 值） |

---

## 三、每月作業流程（SOP）

### Step 1：準備原始流量數據
- 從系統下載當月各線路的 Excel 數據
- 將 17 個線路 Sheet 整合成一個檔案
- 檔名格式：`YYYY-MM_Hanstrong_Traffic_hans_95.xlsx`
- 確認檔案內有「**總表**」Sheet（第一個），欄位為：名字、流入95值、流出95值

### Step 2：告訴 Claude 執行
直接說：
> 「我的 `YYYY-MM_Hanstrong_Traffic_hans_95.xlsx` 已經做好了總表，幫我比對填入月報表，產出 `YYYY-MM_new_每月(供應商)頻寬統計(ZL.VOCOM.CMI)_HANS版本.xlsx`」

Claude 會自動：
1. 讀取總表中的 17 條線路名字與最大 95 值
2. 比對月報表 D 欄名字（自動處理括號差異）
3. 填入對應月份欄位（自動偵測）
4. 保留原始格式（顏色、字體、框線）
5. 產出新檔案

### Step 3：確認後推上 GitHub
確認檔案正確後，說「幫我推上 GitHub」即可。

---

## 四、腳本說明

### 腳本位置
`C:\Users\kench\AppData\Local\Temp\merge_files_v2.js`

> ⚠️ Temp 目錄可能被清除，若腳本不見了，告訴 Claude 重新建立即可。

### 執行方式
```bash
cd C:\Users\kench\AppData\Local\Temp
npm install xlsx xlsx-js-style exceljs   # 首次或套件遺失時執行
node merge_files_v2.js
```

### 腳本邏輯
1. 讀取 `YYYY-MM_Hanstrong_Traffic_hans_95.xlsx` 的**總表** Sheet
2. 建立名字 → 最大95值（Mbps）的對照表（名稱標準化：全小寫、去除括號）
3. 用 ExcelJS 讀取月報表原檔（完整保留所有樣式）
4. 搜尋「N月」欄位（自動偵測，不寫死）
5. 逐行比對 D 欄名字，填入最大 95 值，數字格式設 `0.00`
6. 輸出新檔案

### 核心對照關係（10 條線路）

| 月報表 D 欄名稱 | 總表對應名稱 |
|--------------|------------|
| Zenlayer_香港回國總帶寬 | Zenlayer_香港回國總帶寬 |
| HANSCDN_10M_HK_Zl_CN_line1 | HANSCDN_10M_HK_Zl_CN_line1 |
| HANSCDN_10M_HK_Zl_CN_line2 | HANSCDN_10M_HK_Zl_CN_line2 |
| Zenlayer_香港國際總帶寬 | Zenlayer_香港國際總帶寬 |
| HANSCDN_HK_Zl_global_100M | HANSCDN_HK_Zl_global_100M |
| Zenlayer_日本回國帶寬 | Zenlayer_日本回國帶寬 |
| HANSCDN_20M_JP_Zl_CN | HANSCDN_20M_JP_Zl_CN |
| Zenlayer_日本國際帶寬 | Zenlayer_日本國際帶寬 |
| Zenlayer_CMI(待更正) | Zenlayer_CMI（括號自動忽略）|
| HANSCDN_10M_HK_VCOM | HANSCDN_10M_HK_VCOM |

---

## 五、npm 套件說明

| 套件 | 用途 |
|------|------|
| `xlsx` | 讀取 Excel 基本結構與總表資料 |
| `xlsx-js-style` | 寫入含樣式的 Excel（用於 4. new_Hanstrong 檔） |
| `exceljs` | 讀寫月報表時完整保留原始格式 |

---

## 六、GitHub 操作

### Repo 資訊
- **URL：** https://github.com/RYTECHUSER/anny_change_file
- **Branch：** main
- **Git 初始化位置：** `my-agent/200_Reference/writing-samples/files/`（已有 `.git`）

### 推送指令
```bash
cd "D:\公司共用ONEDRIVE\OneDrive - 恆創網路服務有限公司\Ken Data\claude_use\my-agent\200_Reference\writing-samples\files"
git add "YYYY-MM_new_每月(供應商)頻寬統計(ZL.VOCOM.CMI)_HANS版本.xlsx"
git commit -m "Add YYYY-MM monthly bandwidth summary"
git push
```

### 認證說明
- 使用 **Git Credential Manager**（OAuth App）
- 憑證存於 Windows Credential Manager
- 若憑證失效，重新 `git push` 時會跳出瀏覽器登入視窗

---

## 七、歷史執行紀錄

### 2026-04（4月）
| 項目 | 值 |
|------|---|
| 來源檔 | `2026-04_Hanstrong_Traffic_hans_95.xlsx` |
| 產出檔 | `2026-04_new_每月(供應商)頻寬統計(ZL.VOCOM.CMI)_HANS版本.xlsx` |
| 填入欄位 | 4月（第 10 欄）|
| 更新線路數 | 10 條 |

#### 4月填入數值
| 線路 | 填入值（Mbps）|
|------|-------------|
| Zenlayer_香港回國總帶寬 | 259.48 |
| HANSCDN_10M_HK_Zl_CN_line1 | 12.14 |
| HANSCDN_10M_HK_Zl_CN_line2 | 5.57 |
| Zenlayer_香港國際總帶寬 | 201.41 |
| HANSCDN_HK_Zl_global_100M | 0.1069 |
| Zenlayer_日本回國帶寬 | 171.49 |
| HANSCDN_20M_JP_Zl_CN | 7.80 |
| Zenlayer_日本國際帶寬 | 512.46 |
| Zenlayer_CMI(待更正) | 63.97 |
| HANSCDN_10M_HK_VCOM | 8.74 |

---

## 八、注意事項

1. **月份欄位自動偵測**：腳本會自動搜尋「N月」欄位，不需手動修改腳本，換月直接執行即可
2. **名稱差異處理**：比對時自動全小寫、去除括號，`Zenlayer_CMI(待更正)` 可正確對應 `Zenlayer_CMI`
3. **數字格式**：填入值統一顯示兩位小數（`0.00`），避免整數截斷問題
4. **Zenlayer_CMI 名稱**：月報表標注「待更正」，未來正式名稱確認後記得更新月報表 D 欄
5. **套件遺失**：Temp 目錄可能被 Windows 清除，遺失時重新 `npm install` 即可，套件清單見第五節

---

## 九、下次對話接續方式

把這個 MD 檔內容貼給 Claude，說：
> 「我要執行供應商報表整合，這是上次的紀錄」

Claude 看到這份文件後即可直接接續執行，無需重新說明流程。
