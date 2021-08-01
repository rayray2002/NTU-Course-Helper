# NTU-Course-Helper 台大選課小幫手
從 [台大課程網](https://nol.ntu.edu.tw/nol/coursesearch/index.php) 爬課程資料，根據輸入空堂表及系所代號輸出可以選的通識、國文、外文、體育、系所課程。

# 環境設定
## 執行檔
於 [Release](https://github.com/rayray2002/NTU-Course-Helper/releases) 中下載

## Python環境
### Library install
批量安裝：
> pip install -r py-require.txt

單個安裝：
> pip install pandas
> 
> pip install beautifulsoup4
> 
> pip install requests
> 
> pip install styleframe

# 使用方法
## 輸入
輸入格式：
```
名字
通識領域
一有課節數(,分隔)
二
三
四
五
六
日
系所代碼1
系所代碼2
...
(一個空行)
```
範例：
```
Ray
12345678
一3,4
二7
三2,3,4
四7,8
五3,4,7,8
六
日
9010
9020
9210
9220
9430
2020

Jennifer
123458
一2,3,4
二
三5,10,A,B,C,D
四
五10,A,B,C,D
六
日
2020
9010
9020

```

## 輸出
資料夾結構：
```
├── save_csv
│   ├── 體育.csv
│   ├── 國文.csv
│   ├── 外文.csv
│   ├── 9010.csv
│   └── ...
├── out
│   ├── Ray.xlsx
│   ├── name.xlsx
│   └── ...
└── app
```
輸出.xlsx檔於`.out/name.xlsx`
