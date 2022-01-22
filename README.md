# happig-UHF 

## 運作過程

1. 送出 UDP 指令，連到與 server（執行此程式的裝置） 同個網域下的 RFID 裝置會回應
2. 對有回應的 RFID 裝置進行設定，使其與 server 建立 TCP 連線
3. 對所有 TCP 連線開一個獨立的線程，若該裝置傳來資料則上傳至雲端 database，若雲端 database 失連將備份到本地 database，待自動偵測到重連時自動將本地備份上傳回 server


## 執行

```shell
# （依環境不同可能需要）先設定 server.py 第一行的 python3 路徑
$ sudo ./server.py
```

## 注意事項

1. 確保執行以下指令可以進入 mysql（不需 root 權限）：
	```
	$ mysql -u root -p
	```
2. Mysql 密碼等資訊記錄在 `setting.json` ，和其他 .py 檔放在同一個資料夾裡面。
   
