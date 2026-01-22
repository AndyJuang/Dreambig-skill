# Dream Big 元大公益圓夢計畫 Skills

這是一個 MCP (Model Context Protocol) 伺服器與技能庫，提供生成「Dream Big 元大公益圓夢計畫」申請書與提案簡報的功能。

## 包含技能

### 1. 申請書生成器 (dream-big-application)
**用途**：生成 Word 格式的計畫申請書
- 支援：單位概況表、計畫內容、預算規劃、SDGs 對應、企業連結度說明。
- 檔案：`dream-big-application.skill`
- 執行方式：透過 MCP Server (`index.js`) 或直接匯入 Skill。

### 2. 提案簡報生成器 (dream-big-pitch-deck)
**用途**：生成 7 分鐘提案簡報 (PPTX)
- 支援：13 頁標準結構、數據視覺化、甘特圖、企業連結專頁。
- 檔案：`dream-big-pitch-deck.skill`
- 結構：封面、問題痛點、我們是誰、計畫目標、服務對象、執行策略、時程規劃、企業志工合作、資源運用、量化成效、質化成效、永續擴散、結語。

## 安裝與使用

### 前置需求
- [Node.js](https://nodejs.org/) (建議 v18 以上)
- Git

### 1. 下載與安裝

```bash
git clone https://github.com/AndyJuang/Dreambig-skill.git
cd Dreambig-skill
npm install
```

### 2. 在 Claude Desktop 中使用 (作為 MCP Server)

要讓 Claude 使用申請書生成功能，請編輯您的 Claude Desktop 設定檔：

- **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`

在 `mcpServers` 區塊中加入以下設定（請修改 `/path/to/Dreambig-skill` 為您實際下載的絕對路徑）：

```json
{
  "mcpServers": {
    "dream-big": {
      "command": "node",
      "args": [
        "/path/to/Dreambig-skill/index.js"
      ]
    }
  }
}
```

儲存檔案後，重新啟動 Claude Desktop。

### 3. 使用方式

在對話中，您可以說：

- **申請書**：「幫我生成 Dream Big 申請書」、「填寫公益計畫提案」
- **簡報**：「製作 Dream Big 提案簡報」、「圓夢計畫簡報設計」

## 專案結構

- `SKILL.md`: 申請書生成技能主說明
- `dream-big-application.skill`: 申請書技能包
- `dream-big-pitch-deck.skill`: 簡報生成技能包
- `scripts/`: 生成腳本 (如 `generate-application.js`)
- `assets/`: 範例資料與模板
- `references/`: 詳細指南與參考文件

## 參考資料

- **欄位說明**: `references/fields-guide.md`
- **範例資料**: `assets/sample-data.json`
