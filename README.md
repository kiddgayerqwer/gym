# 教練學員管理

單頁網頁 APP（`index.html`），功能：

- 學員基本資料建置
- 教練上課紀錄
- 學員簽到
- 合約上傳與保存（含到期提醒）

直接用瀏覽器開啟 `index.html` 即可使用；資料預設存在該瀏覽器本機（localStorage / IndexedDB）。

## GitHub 備份（選用）

APP 右上角「備份」可設定把資料自動上傳到**另一個私人 repo**（`data.json` 與合約檔案）。
程式會檢查對方 repo 是否為私人，公開 repo 一律拒絕。請勿把學員資料放進本 repo。
