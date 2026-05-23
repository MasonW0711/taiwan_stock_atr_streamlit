# 台股 ATR 移動停利與融資風險追蹤系統

## 專案用途

這是一個可部署到 Streamlit Community Cloud 的台股持股追蹤工具。

你可以每天在手機、平板、公司電腦或家裡電腦上，透過瀏覽器上傳自己的持股資料，快速查看：

- ATR14 與最新 TR
- ATR 移動停利價
- 停利線是否被跌破
- 持股未實現損益
- 融資餘額風險
- 操作建議

## 功能說明

本專案提供以下功能：

1. 上傳 portfolio.csv 作為持股來源。
2. 若券商系統不能匯出 CSV，可改上傳持股截圖做 OCR 辨識。
3. OCR 辨識後可先在畫面上的草稿表手動修正，再進入 ATR 計算。
4. 選擇性上傳 margin.csv 做融資風險判斷。
5. 選擇性上傳 stop_history.csv 延續既有移動停利歷史。
6. 自訂 ATR 週期與最近高點期間。
7. 依股票類型套用 ATR 倍數，並支援手動調整。
8. 計算持股總市值、投入成本、未實現損益與未實現損益率。
9. 判斷是否跌破 ATR 移動停利與融資風險是否升高。
10. 下載 atr_report.csv、atr_report.xlsx、updated_stop_history.csv。

## 安裝套件方法

在專案根目錄執行：

```bash
pip install -r requirements.txt
```

## 本機執行方式

安裝完成後，在專案根目錄執行：

```bash
streamlit run app.py
```

啟動後，瀏覽器會開啟本機 Streamlit 畫面。

## 如何編輯 portfolio_sample.csv

請以 portfolio_sample.csv 為基礎，複製一份改成自己的 portfolio.csv。

欄位如下：

- symbol：台股代號，使用 yfinance 格式，例如 2382.TW
- name：股票名稱
- shares：持股股數
- cost：平均成本或損益平衡價
- category：股票類型，決定 ATR 倍數

預設支援的類型：

- ETF
- financial
- low_volatility
- normal
- high_volatility

如果 category 無法辨識，系統會預設使用 2.0 倍 ATR。

## 如果券商系統不能下載 CSV

你可以改用持股截圖辨識模式。

使用方式如下：

1. 在 App 側邊欄上傳一張或多張持股截圖。
2. 盡量讓截圖包含股票代號、股票名稱、股數與平均成本。
3. 按下「辨識持股截圖」。
4. 系統會先產生 OCR 草稿表。
5. 在草稿表中手動修正股票代號、股數、成本價與 category。
6. 確認後再按「開始計算」。

建議：

- 盡量裁成只有持股表格的區域。
- 畫面文字要清楚，避免反光與模糊。
- 若券商畫面顯示的是「張」而不是「股」，請在 App 中切換正確單位。
- 若 OCR 沒抓到成本價，請手動補上後再計算。

## 如何準備 margin.csv

margin.csv 是選填檔案。

欄位如下：

- symbol
- date
- margin_balance
- margin_change
- margin_change_rate
- foreign_buy_sell
- investment_trust_buy_sell
- price_change_rate

若未上傳 margin.csv，系統仍可正常執行，只是融資風險欄位會顯示「尚未提供融資資料」。

## 如何使用 stop_history.csv

stop_history.csv 用來記錄每次計算後的最高 ATR 移動停利價。

欄位如下：

- symbol
- name
- last_trailing_stop
- last_update_date

使用方式：

1. 第一次使用時可不提供 stop_history.csv。
2. 系統會從空白歷史開始計算。
3. 本次計算完成後，下載 updated_stop_history.csv。
4. 下次使用時再把 updated_stop_history.csv 上傳回 App。

## ATR 是什麼

ATR 是 Average True Range，常譯為平均真實波幅。

它用來描述一段期間內股價波動幅度。ATR 越大，代表波動越大；ATR 越小，代表波動越小。

本專案使用標準 ATR 計算方式：

```text
TR = max(
    今日最高價 - 今日最低價,
    abs(今日最高價 - 昨日收盤價),
    abs(今日最低價 - 昨日收盤價)
)

ATR = 最近 N 日 TR 的平均值
```

## ATR 移動停利如何解讀

系統會先抓最近 M 個交易日最高價，再扣掉 ATR 乘上倍數，得到原始 ATR 移動停利價。

```text
recent_high = 最近 M 個交易日最高價
raw_trailing_stop = recent_high - ATR * atr_multiplier
```

如果目前股價高於最終 ATR 移動停利價，代表尚未跌破防守位置。

如果目前股價低於或等於最終 ATR 移動停利價，代表已跌破移動停利，應提高警覺。

## 為什麼移動停利線只能上移不能下移

移動停利的目的，是把已經建立的保護水位留住。

如果停利線可以往下調，就會讓風險控管失真，原本已經保住的部位又暴露回去。因此本專案採用以下規則：

1. 沒有歷史停利價時，直接採用本次計算出的原始停利價。
2. 有歷史停利價時，只要本次新算出的停利價較低，就維持舊值不變。
3. 只有在本次停利價更高時，才更新停利線。

## 如何部署到 Streamlit Community Cloud

部署流程如下：

1. 建立 GitHub repository。
2. 將 app.py、requirements.txt、README.md 與 sample CSV 上傳到 GitHub。
3. 不要把真實 portfolio.csv、margin.csv、stop_history.csv 提交到公開 repository。
4. 前往 Streamlit Community Cloud。
5. 連接你的 GitHub repository。
6. 在 App entry point 選擇 app.py。
7. 點選 Deploy。
8. 等待部署完成後，用瀏覽器打開 App。
9. 在 App 中上傳自己的 portfolio.csv、margin.csv、stop_history.csv。
10. 計算完成後下載 atr_report.csv、atr_report.xlsx 與 updated_stop_history.csv。

## 如何把專案上傳到 GitHub

以下為常見流程：

```bash
git init
git add .
git commit -m "Initial Streamlit ATR app"
git branch -M main
git remote add origin <你的 GitHub repository URL>
git push -u origin main
```

## 如何在 Streamlit Community Cloud 選擇 app.py

建立 App 時，在部署畫面的 App entry point 欄位選擇 app.py。

本專案已經使用 app.py 作為根目錄入口檔，符合 Streamlit Community Cloud 的常見部署方式。

## 如何避免把真實持股資料上傳到公開 GitHub

請務必注意以下事項：

1. 如果 GitHub repository 是公開的，不要把自己的真實 portfolio.csv、margin.csv、stop_history.csv 上傳到 GitHub。
2. 建議只提交 sample CSV。
3. 真實資料請透過 Streamlit App 的檔案上傳功能處理。
4. Streamlit Community Cloud 不應依賴本機 CSV 永久保存資料。
5. 每次使用後請下載最新 stop_history.csv，下次再重新上傳。

## Secrets 與 API 金鑰管理

如果未來加入券商 API、資料 API 或其他外部金鑰，請使用 Streamlit Community Cloud 的 Advanced settings 設定 secrets。

本機開發通常會使用 .streamlit/secrets.toml，但這個檔案不應提交到 GitHub。

也不要把 API Key、Token、帳密寫死在程式碼裡。

## 常見錯誤排除

### 1. App 啟動失敗

請先確認已執行：

```bash
pip install -r requirements.txt
```

並使用：

```bash
streamlit run app.py
```

### 2. 上傳 portfolio.csv 後顯示缺少欄位

請確認檔案至少有以下欄位：

- symbol
- name
- shares
- cost
- category

### 3. 某些股票顯示資料下載失敗

yfinance 偶爾可能遇到短暫連線問題或資料源異常。系統會保留該股票列，但顯示「資料下載失敗，暫不判斷」，不影響其他股票計算。

### 4. 某些股票顯示資料不足

如果股票歷史日線資料不足 ATR 週期或最近高點期間，系統會標示「資料不足，暫不判斷」。

### 5. 沒有上傳 margin.csv

這是正常情況，系統仍可運作，只是融資風險欄位會顯示「尚未提供融資資料」。

### 6. 沒有上傳 stop_history.csv

系統會從空白歷史開始計算，並在本次完成後提供新的 updated_stop_history.csv 下載。

### 7. Excel 下載按鈕無法使用

若 openpyxl 產生 Excel 失敗，畫面會顯示錯誤訊息，但 CSV 下載仍可繼續使用。
