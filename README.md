# PDF 转换拼接工具 v3.0 · pdf-convert-merge

把一批 **Word / Excel**（Windows 还有 CAD、标签）批量转成 PDF；也能给一批 PDF **一键换掉首页** —— 把签好字的那份扫描件贴回去。

**原始文件只读，永不改动。**

> 两个模式，窗口顶部两个大按钮，点一下切换：
>
> - **导出 PDF** —— 一批 Office 文件批量转 PDF，打印签字之前那一步。每份原件 → 各自一份 PDF，落在原件所在文件夹里。
> - **PDF 首页替换** —— 一份扫好的多页签字页 PDF（一页对应一份原件），把第 N 页贴到第 N 份原件的第 1 页，后面几页原样保留。签字页一次扫成册，靠顺序一一对应，不用改名。

---

## 它解决什么

现实里的痛点：

1. **批量转 PDF 麻烦** —— 十几二十份 Word/Excel，一份一份「另存为 PDF」要点几十次鼠标。
2. **签字首页归位麻烦** —— 签完字的首页是扫描件，一份一份贴回原件第 1 页更费劲。

这个工具两步全包：先「导出 PDF」批量转，再「PDF 首页替换」把签好字的扫描页贴回去。

## 目录结构

```
mac/          macOS 版（含开发期的自测与回归脚本）
  seq_core.py                    核心逻辑层（无 GUI 依赖）
  seq_app.py                     tkinter 图形界面
  setup_env.py                   环境自检 / 按需重建虚拟环境 / 启动（两平台共用同一份）
  启动PDF转换拼接.command         macOS 双击入口
  使用说明书.txt                  完整使用文档
  _regress/                      回归与验收脚本

windows/      Windows 交付包
  seq_core.py  seq_app.py  setup_env.py    ← 与 mac/ 逐字节一致
  启动PDF转换拼接.bat            日常入口（CRLF + 纯 ASCII + 无 BOM）
  运行前的环境准备.bat            首次运行 / 重装环境
  使用说明书.txt                  同上
```

> `seq_core.py` / `seq_app.py` / `setup_env.py` 属于**程序本体，两份包逐字节一致**。
> 改代码请两边同步（`mac/_regress/sync_win_package.py` 是为此写的同步脚本）。

## 怎么跑

两个平台都是**双击启动器**即可，首次会自动准备运行环境（建 `.venv` → 装依赖 → 拉起界面）：

| 平台 | 入口 | 备注 |
|---|---|---|
| macOS | `mac/启动PDF转换拼接.command` | 首次自动准备环境（约 1-2 分钟），以后秒开 |
| Windows | `windows/启动PDF转换拼接.bat` | 第一次先双击 `运行前的环境准备.bat` |

**macOS 首次转 Word / Excel 会弹系统授权**，点「允许」—— 授权按应用记，允许一次以后不再问。

## 支持的格式

| 格式 | macOS | Windows |
|---|---|---|
| Word（.doc .docx .rtf） | ✓ | ✓ |
| Excel（.xls .xlsx） | ✓ | ✓ |
| PDF（首页替换） | ✓ | ✓ |
| CAD 图纸 | — | ✓ |
| 标签文件 | — | ✓ |

Word / Excel 转 PDF 需要本机装有 **Microsoft Office** 或 **LibreOffice**。

## 技术要点

- **Python + tkinter**，无重量级框架依赖
- **虚拟环境自动管理**：`setup_env.py` 探测 → 按需创建 `.venv` → 装依赖 → 拉起界面
- **配置记忆**：输出路径、后缀等设定自动保存，下次打开自动恢复
- **文件拖入**：支持从访达 / 资源管理器直接拖文件进窗口

## 同系列版本

本工具与 [pdf-firstpage-swap-1.0](https://github.com/PRLODIGJAL/pdf-firstpage-swap-1.0) / [pdf-firstpage-swap-2.0](https://github.com/PRLODIGJAL/pdf-firstpage-swap-2.0) 同源。那两个是独立的「首页替换」工具（v1 靠文件名相似度配对，v2 靠顺序硬对应），本仓库是「转换 + 首页替换」的整合版。
