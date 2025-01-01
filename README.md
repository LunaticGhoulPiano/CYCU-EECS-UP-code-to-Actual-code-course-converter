# 給中原大學電機資訊學士班系助理使用的：UP虛擬碼與真實課程代碼轉換器

## 用途
- 讀取全校開課總表與各課程修課名單
- 產生```總表.xlsx```，合併所有課程並新增```系所課程代碼```、```必選修```、```開課學系```、```開課班級```欄位

## 檔案架構
```
.CYCU-EECS-UP-code-to-Actual-code-course-converter
├──LICENSE <--- 你可以不用管它
├──README.md <--- 說明文件
├──requirements.txt <--- 程式需要使用的libraries
├──main.py <--- 程式本體
├──StarClass_List.xls <--- 這是全校開課總表，你要自行下載且一定要是這個檔案名稱，必須是```.xls```不可為```.xlsx```
├──UP106G_修課名單.xls <--- 範例的要合併課程修課名單，必須為"虛擬碼" + "_修課名單.xls"，也必須是```.xls```不可為```.xlsx```
├──... <--- 其它的課程名單
└──總表.xlsx <--- 自動產生的合併後總表
```

## 使用說明
1. 安裝Python：到[Microsoft Store - Python](https://apps.microsoft.com/detail/9pjpw5ldxlz5?hl=zh-TW&gl=TW)下載安裝
2. 安裝Git（選擇性）：到[Git官網](https://git-scm.com/downloads)下載安裝
3. 測試是否有安裝：
在Windows搜尋欄打"終端機"或是"cmd"
輸入：
    - ```python --version```或是```python -V```：若安裝正確會輸出```Python 3.10```（或你安裝的python版本）
    - ```git --version```或是```git -v```：若安裝卻會輸出```git version 2.47.1.windows.2```（或你安裝的python版本）
4. 下載此程式：
    - 如果不會寫程式就直接在[本專案網頁](https://github.com/LunaticGhoulPiano/CYCU-EECS-UP-code-to-Actual-code-course-converter)點擊```<> Code```右邊的小箭頭-```Local```-```Download ZIP```，然後把解壓縮後的資料夾放到你想放的地方
    - 如果你會用板控就直接git clone或是用Github Desktop都行
5. 下載相關檔案：
    - 由於我是學生沒有系助理的權限，沒辦法幫你寫自動化一鍵下載，因此請自行用你的手動老方法一門一門課慢慢下載修課名單到你剛剛解壓縮的資料夾內（例如假設```UP106G_修課名單.xls```如果沒有人修就不用下載）
    - 然後一樣在學校的系統下載開課總表到資料夾內```StarClass_List.xls```，我沒有權限
6. 執行：
    1. 打開終端機(cmd或是powershell)並切換到目標資料夾：
        - 在資料夾上按右鍵選擇```在終端中開啟(T)```
        - 或是使用指令：
            1. 在Windows搜尋欄打```終端機```或```cmd```或```powershell```並開啟
            2. 切換到你剛剛解壓縮的那個資料夾：
                - ```dir```：列出當前路徑底下的所有檔案
                - ```cd ..```：往上一層路徑走，假設終端機顯示```C:\Users\TA>```，輸入後就會是```C:\Users>```
                - ```cd 路徑```：切換到當前目錄下的某個路徑，假設終端機顯示```C:\Users\TA>```，我知道我的資料夾```test_folder```在```Downloads```中，而```Downloads```又在```TA```中，所以我可以：
                    - ```cd Downloads```之後再```cd test_folder```，兩次切換到目標資料夾
                    - 或是 ```cd Downloads\test_folder```，直接切換到目標路徑
    2. 安裝程式所需要的Libraries：
        - 在終端機輸入```pip install -r requirements.txt```
    3. 執行程式：
        - 請確認程式用到的```StarClass_List.xls```與所有```{虛擬碼}_修課名單.xls```都已關閉
        - 在終端機輸入```python main.py```，等它跑完就可以了
        - 如果之前已經有跑過，程式會先問你是否要取代就的總表，要的話輸入```Y```，否則可以直接按```Enter```
        - 最後的輸出就是```總表.xlsx```，你確認過後就可以丟給課註組了