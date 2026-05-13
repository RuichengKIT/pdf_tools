# pdf_tools

一个简单的 Python GUI 工具，用来把 PDF 页面导出为 PNG 图片。

## 功能

- 选择 PDF 文件
- 选择输出文件夹
- 支持导出全部页面或指定页码
- 支持中文 / English 界面切换
- 输出文件按 PDF 页码命名，例如 `1.png`、`2.png`、`3.png`

## 页码格式

在 `导出页码` 输入框里选择要导出的页码：

```text
1,3,5-8
```

留空表示导出全部页面。

## 安装

```powershell
py -m pip install -r requirements.txt
```

如果 `py` 不可用，请先安装 Python：https://www.python.org/downloads/

## 运行

```powershell
py pdf_to_images_gui.py
```
