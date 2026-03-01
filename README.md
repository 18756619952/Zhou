# Zhou

将 PPT 转成可展示网页的小工具（`.pptx` → `index.html` + `images/`）。

## 功能

- 读取 `.pptx` 并生成网页幻灯片。
- 自动提取每页文字与图片。
- 支持按钮和键盘翻页（左右方向键 / PageUp / PageDown）。
- 生成纯静态文件，方便部署到任意静态站点服务。

## 安装依赖

```bash
python3 -m pip install python-pptx
```

## 使用方法

```bash
python3 convert_ppt_to_web.py 你的文件.pptx -o output_dir
```

示例：

```bash
python3 convert_ppt_to_web.py demo.pptx -o dist
```

默认会输出到：

- `dist/demo/index.html`
- `dist/demo/images/*`

## 本地预览

```bash
python3 -m http.server 8000 -d dist
```

打开浏览器访问：<http://localhost:8000/demo/index.html>

## 说明

- 该脚本更适合“展示型”网页转换，保留主要内容结构（标题、文字、图片）。
- 复杂动画、切换特效、嵌入字体、部分图表样式可能无法 1:1 还原。
