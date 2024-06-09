## 專案簡介
本專案旨在自動化爬取台灣主要求職網站（518、1111、小雞上工及yes123）的求職資料。該程式會根據不同的網站格式，自動生成HTTP請求，下載並解析應徵者的資料，最終輸出為DataFrame格式，便於進一步處理和分析。

## 功能介紹
- 518 人力爬取:
`creat_request_vacancies_518(page)`: 發送職缺ID的POST請求。
`creat_request_resume_518(base_resume)`: 發送應徵者資料的GET請求。
`parse_518_page()`: 主要爬取程式，處理多頁職缺和應徵資料。
- 1111 人力爬取:
`creat_request_vacancies_1111()`: 發送職缺ID的GET請求。
`creat_request_resume_1111(base_resume)`: 發送應徵者資料的GET請求。
- 小雞上工 人力爬取:
`download_resume_chicken()`: 發送應徵者資料的POST請求並下載資料。
- yes123 人力爬取:
`create_request_123(page)`: 創建職缺資料的POST請求。

## 安裝指南
### 前提條件
- Python 3.12。
- 安裝所需的Python套件：
```
pip install requests openpyxl lxml pandas pylightxl
```

### 步驟
1. 克隆此專案：`git clone https://github.com/yourusername/your-repo.git`
2. 執行爬取程式：`python Taiwan Job Bank Resume Scraper.py`

## 使用說明
- 確保已安裝所需的Python庫:urllib, json, requests, openpyxl, lxml, pandas, re, pylightxl。

- 在程式中設定正確的Cookie和其他必要的標頭資訊。這些資訊用於模擬登入狀態並獲取授權訪問履歷資料。

- 程式將依次從518人力銀行、1111人力銀行、yes123人力銀行爬取履歷資料,並下載小雞上工的履歷Excel檔。

- 爬取完成後,程式會將所有履歷資料整合到一個DataFrame中,並根據姓名和應徵日期欄位去除重複項。

- 最終的整合結果將寫入到指定的Excel檔案中(預設為履歷格式_測試.xlsx)。每次執行程式都會將新的資料添加到現有的Excel檔案中。
