# Tech Feasibility Evaluator

批次申請案技術可行性自動化評估系統

> ⚠️ **重要說明**：本工具僅作為初步篩選輔助，專業領域的最終評審仍會由領域專家進行人工審查。

---

## 📖 專案簡介

基於 Google Apps Script 的自動化評估工具，整合 OpenAI API 進行大量申請案的技術可行性初步審查。支援自訂評估框架、批次處理、智慧跳過已評估項目，協助組織提高初步篩選效率。

---

## 🚀 安裝設定

### 1. 準備試算表

建立 Google Sheets，包含以下工作表：
- `原始申請資料`
- `技術可行性評估框架_Claude`
- `技術可行性評估框架_GPT`

### 2. 設定欄位

在「原始申請資料」第一列設定欄位：

`方案別` | `設備數` | `用途` | `背景資訊` | `Claude評估` | `Claude說明` | `chatGPT評估` | `chatGPT說明`

### 3. 部署腳本

Google Sheets → 擴充功能 → Apps Script → 複製貼上程式碼

### 4. 設定 API Key

Apps Script → 專案設定 → 指令碼屬性 → 新增 `OPENAI_API_KEY`

### 5. 設定評估框架

在評估框架工作表設定（A 欄：欄位，B 欄：內容）：

`系統角色` | `設備規格說明` | `評估框架` | `評估範例` | `輸出格式要求`

---

## 📝 使用方法

**選取評估**：選取列 → 選單「🔍 自動審查」→「評估選取的列」

**批次評估**：選單「🔍 自動審查」→「批次評估空白（兩者）」

---

## ⚠️ 注意事項

- API Key 絕不寫在程式碼中，使用指令碼屬性
- 每次 API 呼叫產生費用，注意成本控制
- 欄位名稱必須與程式碼完全一致
- AI 評估僅供參考，重要決策需人工覆核

---

---

# Tech Feasibility Evaluator

Automated Batch Application Technical Feasibility Assessment System

> ⚠️ **Important Notice**: This tool is for preliminary screening assistance only. Final professional evaluation will be conducted by domain experts through manual review.

---

## 📖 Project Overview

A Google Apps Script-based automation tool that integrates OpenAI API for preliminary technical feasibility reviews of large application volumes. Supports customizable evaluation frameworks, batch processing, and smart skip of evaluated items to improve organizational screening efficiency.

---

## 🚀 Installation & Setup

### 1. Prepare Spreadsheet

Create Google Sheets with these sheets:
- `原始申請資料`
- `技術可行性評估框架_Claude`
- `技術可行性評估框架_GPT`

### 2. Configure Columns

Set columns in first row of applications sheet:

`方案別` | `設備數` | `用途` | `背景資訊` | `Claude評估` | `Claude說明` | `chatGPT評估` | `chatGPT說明`

### 3. Deploy Script

Google Sheets → Extensions → Apps Script → Copy and paste code

### 4. Configure API Key

Apps Script → Project Settings → Script Properties → Add `OPENAI_API_KEY`

### 5. Set Up Frameworks

In framework sheets (Column A: Field, Column B: Content):

`系統角色` | `設備規格說明` | `評估框架` | `評估範例` | `輸出格式要求`

---

## 📝 Usage

**Selected Rows**: Select rows → Menu "🔍 自動審查" → "評估選取的列"

**Batch Process**: Menu "🔍 自動審查" → "批次評估空白（兩者）"

---

## ⚠️ Important Notes

- Never hardcode API keys; use script properties
- Each API call incurs costs; monitor usage
- Column names must exactly match code configuration
- AI assessments are for reference only; critical decisions require manual review
