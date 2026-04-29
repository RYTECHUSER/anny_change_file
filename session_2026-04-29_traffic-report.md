# Session 紀錄：Hanstrong Traffic 95th Percentile 報表整合
**日期：** 2026-04-29  
**操作者：** RYTECHUSER（claude_hans1@hans-tech.asia）

---

## 背景說明

本次作業是從原始流量數據 Excel，提取各線路的 95th percentile 值，整合進月報表，並推送至 GitHub。

---

## 檔案清單

| 檔案 | 路徑 | 說明 |
|------|------|------|
| 原始流量數據 | `my-agent/200_Reference/writing-samples/files/3.Hanstrong_Traffic_hans_95_202604.xlsx` | 17 個 Sheet，每個 Sheet 是一條線路的原始流量數據 |
| 95 值彙整報表 | `my-agent/200_Reference/writing-samples/files/4. new_Hanstrong_Traffic_hans_95.xlsx` | 總表（第一個 Sheet）+ 原始 17 個 Sheet |
| 月報表原檔 | `my-agent/200_Reference/writing-samples/files/1. 每月(供應商)頻寬統計(ZL.VOCOM.CMI)_HANS版本.xlsx` | 原始月報，含 Zenlayer / 遠傳(CMI) / VOCOM 三段 |
| 月報表新版 | `my-agent/200_Reference/writing-samples/files/5. new_每月(供應商)頻寬統計(ZL.VOCOM.CMI)_HANS版本.xlsx` | 已填入 2026年4月 95th percentile 數值 |

---

## 完成工作

### 1. 生成 `4. new_Hanstrong_Traffic_hans_95.xlsx`
- 讀取 `3.Hanstrong_Traffic_hans_95_202604.xlsx` 的 17 個 Sheet
- 從每個 Sheet 抓取 流入95值 / 流出95值（兩種格式自動偵測）
- 新增「總表」Sheet 置於第一位，欄位：名字、流入95值、流出95值
- 流入 vs 流出比較，**數值較大的格子填紅色背景、黑色字**
- 保留原始 17 個 Sheet

#### 兩種 Sheet 格式說明
| 格式 | 特徵 | 95值位置 |
|------|------|---------|
| 格式A（恆創 5 條線）| 3 欄：日期、流入、流出 | Row0[5], Row1[5] |
| 格式B（其餘 12 條線）| 7 欄：日期+4個跳過+流入+流出 | Row0[8], Row1[8] |

### 2. 生成 `5. new_每月(供應商)頻寬統計(ZL.VOCOM.CMI)_HANS版本.xlsx`
- 比對「總表」名字 vs 月報表 D 欄名字（名稱標準化處理，自動忽略括號）
- 取流入/流出 95 值中**較大者**換算成 Mbps
- 填入對應線路的 **2026年4月** 欄位（欄位 J，index=9）
- 使用 ExcelJS 讀寫，完整保留原始格式（顏色、字體、框線）
- 數字格式設為 `0.00`（保留兩位小數）

#### 填入對應關係（10 條線路）
| 月報表名稱 | 總表對應名稱 | 填入值（Mbps）|
|-----------|------------|-------------|
| Zenlayer_香港回國總帶寬 | Zenlayer_香港回國總帶寬 | 106.34 |
| HANSCDN_10M_HK_Zl_CN_line1 | HANSCDN_10M_HK_Zl_CN_line1 | 0.5219 |
| HANSCDN_10M_HK_Zl_CN_line2 | HANSCDN_10M_HK_Zl_CN_line2 | 4.86 |
| Zenlayer_香港國際總帶寬 | Zenlayer_香港國際總帶寬 | 672.72 |
| HANSCDN_HK_Zl_global_100M | HANSCDN_HK_Zl_global_100M | 0.1146 |
| Zenlayer_日本回國帶寬 | Zenlayer_日本回國帶寬 | 177.96 |
| HANSCDN_20M_JP_Zl_CN | HANSCDN_20M_JP_Zl_CN | 8.21 |
| Zenlayer_日本國際帶寬 | Zenlayer_日本國際帶寬 | 511.04 |
| Zenlayer_CMI(待更正) | Zenlayer_CMI | 67.42 |
| HANSCDN_10M_HK_VCOM | HANSCDN_10M_HK_VCOM | 0.366 |

### 3. 推送至 GitHub
- Repo：https://github.com/RYTECHUSER/anny_change_file
- 已推送：`4. new_Hanstrong_Traffic_hans_95.xlsx`、`5. new_每月(供應商)頻寬統計(ZL.VOCOM.CMI)_HANS版本.xlsx`
- Git 初始化位置：`my-agent/200_Reference/writing-samples/files/`（已有 `.git`）

---

## 腳本存放位置

所有腳本存於 `C:\Users\kench\AppData\Local\Temp\`，Node.js 執行，npm 套件裝在同目錄。

| 腳本檔名 | 功能 |
|---------|------|
| `create_excel_v3.js` | 生成 `4. new_Hanstrong_Traffic_hans_95.xlsx` |
| `merge_files_v2.js` | 生成 `5. new_每月(供應商)頻寬統計...xlsx` |

#### 執行方式
```bash
cd C:\Users\kench\AppData\Local\Temp
node create_excel_v3.js   # 重新產出 file 4
node merge_files_v2.js    # 重新產出 file 5
```

#### 依賴套件
```
xlsx          # 讀取 Excel 基本結構
xlsx-js-style # 寫入含樣式的 Excel
exceljs       # 保留原始格式讀寫
```

---

## 環境資訊

- **OS：** Windows 11（shell 用 Git Bash / PowerShell）
- **Node.js：** v24.15.0
- **Python：** 未安裝（Microsoft Store stub，無法使用）
- **Git credential：** Windows Credential Manager，Git Credential Manager（OAuth App）
- **GitHub 帳號：** RYTECHUSER

---

## 下次繼續可做的事

- [ ] 每月更新：下次換新的 `3.Hanstrong_Traffic_hans_95_YYYYMM.xlsx` 後，重跑腳本填入對應月份欄位
- [ ] 月份欄位自動偵測：目前寫死找「4月」，下次可改成自動抓當月
- [ ] Zenlayer_CMI 名稱待更正後，更新月報表中的顯示名稱
- [ ] 考慮將腳本整理成固定工作流 Skill

---

## 注意事項

- GitHub token 曾在本次 session 中曝露，已建議 revoke。請確認是否已在 `https://github.com/settings/apps/authorizations` 撤銷 Git Credential Manager 的授權
- 月報表中 HANSCDN_10M_HK_Zl_CN_line1 / line2 的名稱在截圖中看到用「ZI」，但實際 Excel 內是「Zl」，比對時已透過 toLowerCase 處理，無影響
