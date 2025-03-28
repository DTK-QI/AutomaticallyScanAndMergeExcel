# Excel 檔案自動掃描與合併工具

這個工具可以自動掃描指定資料夾中的所有 Excel 檔案，並將其合併成單一檔案。為了避免檔案過大，工具會自動將結果分割成較小的檔案。

## 功能特點

- 自動掃描指定資料夾及其子資料夾中的所有 Excel 檔案（.xlsx）
- 將所有 Excel 檔案合併成一個統一的資料集
- 輸出 TSV 格式的合併檔案
- 自動分割大型檔案（預設每個檔案最多 1,000,000 列）
- 錯誤處理機制，確保程式穩定運行

## 系統需求

- Python 3.6 或更高版本
- pandas 函式庫
- openpyxl 函式庫（用於處理 .xlsx 檔案）

## 安裝方法

1. 克隆此專案：
```bash
git clone https://github.com/yourusername/AutomaticallyScanAndMergeExcel.git
```

2. 安裝必要的套件：
```bash
pip install pandas openpyxl
```

## 使用方法

1. 執行 main.py 檔案：
```bash
python main.py
```

2. 當程式提示時，輸入要掃描的資料夾路徑

3. 程式會在指定的資料夾中建立一個名為 `combined_output` 的子資料夾，並在其中生成：
   - combined_data.txt（TSV 格式的合併檔案）
   - combined_part1.xlsx, combined_part2.xlsx, ... （分割後的 Excel 檔案）

## 注意事項

- 請確保有足夠的硬碟空間存放合併後的檔案
- 處理大量資料時可能需要較長時間
- 建議先備份原始資料

## 授權方式

此專案採用 MIT 授權條款 - 詳見 [LICENSE](LICENSE) 檔案