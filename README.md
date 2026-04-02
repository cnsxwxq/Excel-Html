# Excel-Html

将 Excel 文件转换为响应式 HTML 表格的在线工具。

## 功能特点

- 支持 `.xlsx` 和 `.xls` 格式
- 自动识别链接和图片列，转换为可点击链接
- 响应式设计，移动端自动切换为卡片视图
- 内置搜索功能，支持多关键词搜索
- 自动隐藏指定列（拍卖平台、标的类型、省份、城市）
- 拖拽上传支持
- 纯前端实现，无需后端服务

## 使用方法

1. 打开 `index.html` 页面
2. 上传 Excel 文件（支持拖拽）
3. 点击"转换为 HTML"按钮
4. 点击"另存为 html"下载生成的 HTML 文件

## 技术栈

- HTML5 + CSS3 + JavaScript
- [SheetJS (xlsx)](https://github.com/SheetJS/sheetjs) - Excel 文件解析

## 项目结构

```
Excel-Html/
├── index.html      # 主页面
├── script.js       # 核心逻辑
├── style.css       # 样式文件
└── README.md       # 说明文档
```

## License

MIT
