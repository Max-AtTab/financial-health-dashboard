# 曜析數據｜企業健檢與陪跑交付系統

本專案保存可公開的網站與建置工具；客戶問卷、健檢數據、服務帳戶憑證與其他敏感資料不納入版本控制。

## 交付內容

- `site/`：曜析企業健檢網站靜態頁面。
- `scripts/`：13 週現金流量（TWCF）相關的公式建置與驗證腳本。
- `tools/`：陪跑訂閱月度儀表板模板的建置腳本。
- `docs/`：交接與使用說明。

## 目前的正式 Google Sheets 範本

- [健檢問卷標準底稿](https://docs.google.com/spreadsheets/d/1WVWpwp-9Gkz6N2HYayqXq43MeqJY0FDG_d1yR8U_3Ig/edit)
- [健檢系統 v3 標準底稿](https://docs.google.com/spreadsheets/d/1h0zVgEqGDBWLKqY9JNcUCcrCWJLeeiNT1t42uRD3Rfg/edit)
- [陪跑訂閱月度營運儀表板模板](https://docs.google.com/spreadsheets/d/1BF31OFjm4Sb6spUSCSG_aBnyIC0zx9zsn-PSfwD0EqM/edit)

## 資料流程

`健檢問卷` → `健檢系統_v3` → `陪跑訂閱月度營運儀表板`

問卷的「⑤逐月業績」包含年、月、營收、目標、毛利、營業費用、期末現金、應收帳款、應付帳款、存貨、新客戶與流失客戶。陪跑模板透過健檢系統的資料匯入頁自動讀取；客戶副本首次連結時須在 Google Sheets 允許 `IMPORTRANGE` 存取。
