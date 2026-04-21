# PDF 工具

一个轻量的本地 PDF 工具，用 Python + Tkinter 构建，提供查看、拆分、合并、单页旋转与保存能力。界面参考 Adobe PDF 的布局，支持文件拖拽。

## 功能

### 查看 / 编辑
- 打开 PDF 并显示；左侧缩略图导航，可点击跳转
- 工具栏：打开 / 保存 / 另存为 / 翻页 / 缩放 / 适合宽度 / 左旋 / 右旋
- 鼠标滚轮翻页（页面完全可见时直接翻页；超出视口时先在页内滚动，滚到边缘再翻）
- `Ctrl + 滚轮` 缩放；页码输入框回车跳转
- 单页可视化旋转：`↺ 左旋` / `↻ 右旋` 实时更新主视图和缩略图
- 保存：覆盖原文件或另存为，旋转修改会写入输出 PDF

### 拆分
- 按页码规则拆分：
  - 留空 = 每页一个文件
  - `1-3,5` = 把第 1~3 和第 5 页合为一个文件
  - `1-3;4-6` = 输出两个文件

### 合并（PDF + 图片）
- 支持 PDF 与图片混合：`png / jpg / jpeg / bmp / tif / tiff / gif / webp`
- 图片按 150 DPI 转为 PDF 页后按列表顺序并入
- 列表支持上移 / 下移 / 移除 / 清空，多选操作
- 📄 / 🖼 图标区分文件类型

### 文件拖拽
- 「查看 / 编辑」页：拖入 PDF 直接打开
- 「拆分」页：拖入 PDF 自动填路径
- 「合并」页：拖入 PDF / 图片进列表

## 使用

### 直接运行 EXE
构建产物位于 `dist/PDF工具.exe`（单文件，约 21 MB），双击即可运行，无需安装 Python。

### 源码运行

```bash
pip install -r requirements.txt
python pdf_tool.py
```

### 自行打包

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "PDF工具" --collect-data tkinterdnd2 pdf_tool.py
```

## 依赖

- `pypdf` — PDF 读写 / 旋转
- `pypdfium2` — 页面渲染（基于 PDFium）
- `Pillow` — 图片处理与 Tk 显示
- `tkinterdnd2` — 文件拖拽

## 项目结构

```
.
├── pdf_tool.py        # 主程序
├── requirements.txt   # Python 依赖
├── .gitignore
└── README.md
```

## 许可

个人使用，无限制。
