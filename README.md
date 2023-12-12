# SEO Program 使用須知 (For Windows User)
# 在使用前python檔案前先閱讀以下內容

以下為執行python的前置工作，如果電腦已經執行過一次，則無需要再次執行 (除非系統有更新):

1. 安裝python，選擇最新版本 - https://www.python.org/downloads/windows/
2. 安裝路徑及安裝所選項目: 參考此影片 1:13 - 1:55 - https://youtu.be/yivyNCtVVDk?t=72
3. 開啟 Command Propmt - 電腦搜尋欄輸入 CMD
4. 安裝所需python packages: 在搜尋欄分別輸入以下指示，在任何路徑都可，輸入後按enter，將會自動下載

pip install serpapi

pip install google-search-results

pip install pandas

pip install tqdm

pip install openpyxl

5. 將python program (每個file內的全部檔案)從Github下載後，在桌面新增folder，將全部下載資料放入folder內，需要為同一路徑
6. 在Command Propmt進入python program所放置的路徑內，輸入指令: cd {path}

例子: cd C:\Users\user\Desktop\folder
7. 執行python program (假設已放入必須資料在python program內)，輸入指令: python {python program name}，python program name 需要連同file type name (.py)

例子: python python123.py
8. 結果會出現在同一folder內
