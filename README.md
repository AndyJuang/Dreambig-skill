# Dream Big Application Generator Skill

這是一個 MCP (Model Context Protocol) 伺服器，提供生成「Dream Big 元大公益圓夢計畫」申請書的功能。它可以被 Claude Desktop App、Antigravity 或其他支援 MCP 的工具使用。

## 功能

- **generate_application**: 根據提供的 JSON 資料生成完整的 Word (.docx) 申請書，包含表格、預算、計畫內容等。

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

### 2. 在 Claude Desktop 中使用

要讓 Claude 使用此 Skill，請編輯您的 Claude Desktop 設定檔：

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

儲存檔案後，重新啟動 Claude Desktop。接著您就可以在對話中要求 Claude：「幫我生成 Dream Big 申請書」或「填寫計畫資料」。

### 3. 在 Antigravity 中使用

若 Antigravity 支援標準 MCP 協定，請參照其 MCP 伺服器設定方式，加入同上的 Command 與 Args。

通常設定位於 `~/.antigravity/config.json` 或類似位置 (視版本而定)，格式與 Claude 類似。

## 參考資料

- **欄位說明**: `references/fields-guide.md`
- **範例資料**: `assets/sample-data.json`
