# pdf_tools

一个 Windows 友好的 Python GUI PDF 工具箱，支持中文 / English 界面切换。

## 功能

- PDF 转图片：按 PDF 原页码输出 `1.png`、`2.png`、`3.png`
- 合并 PDF：按列表顺序合并多个 PDF，可选择输出路径和文件名
- 拆分 PDF：每页一个 PDF，或把选中页合成一个 PDF
- PDF 转 Word：把每页渲染成图片并写入 `.docx`
- PDF 转 PPT：把每页渲染成图片并写入 `.pptx`
- Word/PPT 转 PDF：调用本机 Microsoft Office 或 LibreOffice 转换
- 加密 PDF：输入密码后可用于导出、合并、拆分、转换，也可单独解密成新 PDF

## 页码格式

需要选择页码的地方都支持：

```text
1,3,5-8
```

留空表示导出全部页面。

## 合并 PDF 默认命名

合并 PDF 时：

- `输出路径` 留空：默认保存到第一个待合并 PDF 所在文件夹
- `合并文件名` 留空：默认使用 `合并PDF_YYYYMMDD_HHMM.pdf`
- 文件名可以不写 `.pdf` 后缀，程序会自动补上

## 转换/解密默认命名

PDF 转 Word/PPT、Word/PPT 转 PDF、解密 PDF 时：

- `输出路径` 留空：默认保存到源文件所在文件夹
- `Word 文件名` 留空：默认使用 `PDF转Word_YYYYMMDD_HHMM.docx`
- `PPT 文件名` 留空：默认使用 `PDF转PPT_YYYYMMDD_HHMM.pptx`
- Word 转 PDF 的 `PDF 文件名` 留空：默认使用 `Word转PDF_YYYYMMDD_HHMM.pdf`
- PPT 转 PDF 的 `PDF 文件名` 留空：默认使用 `PPT转PDF_YYYYMMDD_HHMM.pdf`
- 解密 PDF 的 `PDF 文件名` 留空：默认使用 `解密PDF_YYYYMMDD_HHMM.pdf`
- 文件名可以不写后缀，程序会自动补上

## 转换说明

PDF 转 Word/PPT 使用“页面图片化”的方式，优点是版式稳定，缺点是 Word/PPT 里的文字不可直接编辑。

Word/PPT 转 PDF 需要本机安装以下任意一种软件：

- Microsoft Office
- LibreOffice

## 安装

```powershell
py -m pip install -r requirements.txt
```

如果 `py` 不可用，请先安装 Python：https://www.python.org/downloads/

## 运行

```powershell
py pdf_to_images_gui.py
```
