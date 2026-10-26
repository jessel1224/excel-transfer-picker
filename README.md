# 轉出狀況背景模式 v3

## 更新安裝

1. 將本資料夾中的所有網站檔案上傳 GitHub 儲存庫最外層，覆蓋 manifest.xml、help.html、index.html 等同名檔案。
2. 等待 GitHub Pages 發布完成。可開啟 background-v3.html 確認新檔案可存取（直接瀏覽不會連接 Excel）。
3. 在 Excel 測試副本中移除舊版增益集，再上傳本包的 manifest.xml。新舊版 ID 相同，版本升至 3.0.0.0，避免同時安裝兩個監聽器。
4. 首次啟用可能短暫顯示設定面板，初始化成功後自動收起。平常點 O 欄即開框，無須常駐側邊欄。
5. 如未啟動，常用頁籤 → 轉出狀況 → 啟用自動跳框。每份新的活頁簿需首次手動啟用一次。

## 行為

SharedRuntime 1.1、long lifetime；以 Office.addin.hide() 收起面板，不解除事件訂閱。setStartupBehavior(load) 設定此活頁簿重開時載入。程式在 Office.onReady 重新註冊事件。功能區提供停用（同時取消下次啟動）、重開選擇框及診斷設定入口。停止後不跳框，手動再啟用可恢復。設定頁手動開啟不會自動收起，直到按啟用。

仍只支援 ICU、RCC、WARD 的 O5:O1048576，13 個文字選項不變。相同格再點不觸發 selectionChanged，可換格再點回來或用功能區重開。新增/重建/重新命名工作表後請重開活頁簿。填寫成功後覆寫原文字，含公式時拒絕。原值變更時停止，但非跨使用者原子鎖。儲存格必須可編輯。

開框或寫入錯誤可在「查看設定」讀取。當未支援背景模式時，不宣稱成功。未添加病人資料上傳、分析追蹤或外部後端；保留 ReadWriteDocument 權限，因此仍由院方審查後才正式使用。

## 驗收

離線檢查不等於 Excel 整合測試，以下須實測：

- 上傳新版 XML 可安裝，功能區顯示轉出狀況群組。
- 第一次啟用後面板收起；不開面板直接切換至 O5，出現一個選擇框。
- 關閉設定面板後，O6/O7 仍可開框並正確填回。
- 重開活頁簿，無需打開面板即可於 O 欄跳框。
- 停用後不跳框；重新啟用不產生重複訂閱。
- Q/S 欄、標題列、多格、公式格不寫入；取消不變。
- 保護/唯讀及共同編輯值變更情況正確處理。

## 官方依據

https://learn.microsoft.com/en-us/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime

https://learn.microsoft.com/en-us/office/dev/add-ins/develop/run-code-on-document-open

https://learn.microsoft.com/en-us/office/dev/add-ins/develop/show-hide-add-in
