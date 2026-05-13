# pdf_tools

一个 Windows 友好的 Python GUI PDF 工具箱，支持中文 / English 界面切换。

## 功能

- PDF 转图片：按 PDF 原页码输出 `1.png`、`2.png`、`3.png`
- 合并 PDF：按列表顺序合并多个 PDF
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
