# Gemini 助手 · Excel 加载项

一个类似 “Claude for Microsoft 365” 的 Excel 侧边栏 AI 助手，模型使用 Google Gemini（自带 API Key，个人使用）。

- 侧边栏聊天，流式输出
- Gemini Function Calling 直接读写工作簿：读取区域 / 写入值和公式 / 格式 / 插入删除 / 新建工作表 / 表格 / 图表 / 排序 / 搜索
- 自动附带当前选区作为上下文
- **发票识别录入**：📎 选择 / 粘贴截图 / 拖入发票图片或 PDF（可多张），自动录入“发票台账”工作表，含金额校验公式、重复发票检查、需核对标记
- 每轮修改可一键撤销（基于修改前快照）
- 可选“修改前需确认”
- 模型可选：`gemini-4-hydrogen`、`gemini-4-neon`、`gemini-4-carbon`，或从 API 拉取 / 手动输入任意模型 ID
- 思考深度：默认 / low / medium / high / xhigh

## 安装（侧载）

清单地址：https://swindlershen.github.io/excel-gemini-assistant/manifest.xml

**Excel 网页版 / Microsoft 365 桌面版：**
1. 下载上面的 `manifest.xml` 到本地。
2. Excel →「开始」→「加载项」→「更多加载项」/「高级」→「我的加载项」→「上传我的加载项」，选择 `manifest.xml`。
3. 「开始」选项卡出现 **Gemini 助手** 按钮，点击打开侧边栏，在设置中填入 Gemini API Key。

**Windows 桌面版（共享文件夹方式，持久生效）：**
1. 把 `manifest.xml` 放到一个文件夹，右键 → 属性 → 共享，得到 `\\电脑名\文件夹`。
2. Excel → 文件 → 选项 → 信任中心 → 信任中心设置 → 受信任的加载项目录，添加该网络路径并勾选“显示在菜单中”。
3. 重启 Excel →「加载项」→「共享文件夹」中添加。

## 开发

```bash
npm install
npm run build   # 输出到 dist/
```

部署：`dist/` 推送到 `gh-pages` 分支（GitHub Pages）。

## 隐私

API Key 仅保存在加载项本地存储（localStorage）中，请求直接从 Excel 发往 `generativelanguage.googleapis.com`，没有任何中转服务器。
