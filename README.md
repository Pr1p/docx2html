# DOCX to HTML Converter

基于 [Mammoth.js](https://github.com/mwilliamson/mammoth.js) 的浏览器端 DOCX → HTML 转换工具，零服务端依赖，开箱即用。

## 功能

### 文件转换
- **拖拽/点击上传** — 支持 `.docx` 文件，批量导入
- **实时预览** — A4 纸张样式渲染，所见即所得
- **图片内嵌** — 自动提取 DOCX 内嵌图片为 Base64 Data URI
- **格式校验** — 通过 Magic Bytes 检测文件类型，拒绝旧版 `.doc` 和无效文件

### 多视图与编辑
- **Preview** — A4 纸张仿真预览
- **HTML Source** — 可直接编辑 HTML，Apply Changes 即时更新预览，Reset 恢复原始内容
- **Split View** — 左侧源码、右侧预览同屏

### 模板变量标记
- **手动标记** — 在 Preview 中选中文字，弹出浮动工具栏，输入变量名后替换为 `${varName}`
- **智能建议** — 自动识别中文合同术语（80+ 词汇），推荐英文变量名（如"甲方" → `partyA`）
- **一键全部替换** — 自动统计选中文字在文档中的出现次数，Replace All 一键替换所有相同文本
- **匹配预览** — 点击出现次数可展开查看所有匹配位置的上下文（关键字高亮）

### 变量汇总
- **Variables 标签页** — 自动检测 `${...}` 变量 + 手动标记变量，分表展示
- **信息表** — 变量名、完整语法、出现次数、上下文、来源（auto/manual）
- **Copy All** — 一键复制所有变量列表

### 导出
- **单文件下载** — 下载当前文件的完整 HTML
- **批量下载** — 一键下载所有已转换文件
- **Copy HTML** — 复制当前 HTML 到剪贴板

## 使用

直接用浏览器打开 `index.html` 即可，无需安装任何依赖。

```
# 或者用任意静态服务器
npx serve .
python -m http.server 8080
```

## 技术栈

| 组件 | 说明 |
|------|------|
| Mammoth.js 1.8.0 | DOCX → 语义 HTML 转换（CDN 加载） |
| 原生 HTML/CSS/JS | 无框架、无构建工具、单文件应用 |

## HTML → PDF

转换后的 HTML 可直接用于 OpenHTMLToPDF（Java）生成 PDF。注意：

- Mammoth 输出语义 HTML（无 Flexbox/Grid），与 OpenHTMLToPDF 兼容性好
- 图片已内嵌为 Data URI，OpenHTMLToPDF 开箱支持
- 需要注册 TrueType 字体（如 SimSun）以支持中文渲染
- 建议用 JSoup 将 HTML5 转 XHTML 后再传入 PDF 渲染器

## License

MIT
