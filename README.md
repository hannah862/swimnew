# 🏊 游泳課報名與級數管理（最終版，Streamlit）

上傳學生名單（Excel），逐一設定 **參加/不參加** 與（若參加）**級數 0–5**。
一鍵匯出各班級 Excel，欄位與合計完全符合學校需求。

## 匯出欄位
```
班級｜座號｜姓名｜參加意願｜0級｜1級｜2級｜3級｜4級｜5級
```
- 參加意願：1=參加，0=不參加
- 0級～5級：符合條件之學生=1，其餘=0
- 最後一列「合計」：參加總人數與各級人數

## 本機執行
```bash
pip install -r requirements.txt
streamlit run swim_enroll_web_final.py
```

## 部署到 Streamlit Cloud
1. 將本專案上傳到 GitHub（整個資料夾）。
2. 到 https://share.streamlit.io → New app → 選擇此 repo 與 branch。
3. **Main file path** 輸入：`swim_enroll_web_final.py` → Deploy。
4. 取得網址後分享給老師使用。

### （可選）加入簡易密碼保護
在 Streamlit Cloud 的 **Settings → Secrets** 新增：
```
APP_PASSWORD="你的密碼"
```
然後將下列程式碼加到 `swim_enroll_web_final.py` 最上方（`st.set_page_config` 前）：
```python
import os, streamlit as st
pwd = st.sidebar.text_input("請輸入密碼", type="password")
if os.environ.get("APP_PASSWORD") and pwd != os.environ["APP_PASSWORD"]:
    st.warning("請輸入正確密碼以使用本系統")
    st.stop()
```

## 名單樣板
本專案包含 `學生名單_樣板.xlsx`，欄位：**班級、座號、姓名、參加意願、級數**。
