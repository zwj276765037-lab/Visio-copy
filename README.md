<p align="right">
  <strong>简体中文</strong> | <a href="README.en.md">English</a>
</p>

# Visio Copy

`visio-copy` 是一个 Codex skill，用于将论文、PPT 或截图中的技术图复刻到 Microsoft Visio 中，并尽量生成可编辑的矢量形状，而不是简单粘贴一张图片。

它主要面向体系结构图、硬件框图、数据流图和论文中的模块示意图，支持通过 Visio COM 自动化创建矩形、箭头、总线、表格、网格、公式标签和堆叠结构，并通过预览图与局部 crop 对比进行迭代修复。

## 功能特点

- 生成可编辑的 Visio shape，而不是截图式复制。
- 使用源图像像素坐标作为统一坐标系，便于精确复刻布局。
- 迭代时使用锁定的底层参考图，最终交付时移除底图。
- 提供 PowerShell Visio COM 绘图脚手架。
- 提供 Python 工具用于颜色区域提取和局部 crop 对比。
- 包含针对堆叠网格、Key Matrix、小方块阵列等复杂结构的复刻规则。

## 部分效果展示：原图 vs Copy

下面展示的是参考原图与 `visio-copy` 复刻后的 Visio 结果。Copy 一侧是可编辑 Visio 图形的截图，不是直接贴图替代。

**版本说明：** 这是 `1.0` 版本。当前 `visio-copy` 对常规模块框图和架构图效果较好，但对堆叠密集图形的绘制效果还不够好。尤其是大量小方块、遮挡层级、重复单元计数、斜线纹理和密集文字排版，仍然需要人工 crop 级检查和多轮修正。

| 案例 | 原图 | Copy |
| --- | --- | --- |
| 硬件架构图 | <img src="assets/showcase/hardware-original.png" width="420" alt="Hardware original"> | <img src="assets/showcase/hardware-copy.png" width="420" alt="Hardware Visio copy"> |
| PADE 架构图 | <img src="assets/showcase/pade-original.png" width="420" alt="PADE original"> | <img src="assets/showcase/pade-copy.png" width="420" alt="PADE Visio copy"> |
| Bit-serial stage-fusion 图 | <img src="assets/showcase/bsf-original.png" width="420" alt="BSF original"> | <img src="assets/showcase/bsf-copy.png" width="420" alt="BSF Visio copy"> |

## 仓库结构

```text
.
|-- README.md
|-- README.en.md
|-- SKILL.md
|-- assets/
|   `-- showcase/
|-- agents/
|-- references/
|   |-- redraw-checklist.md
|   `-- stacked-grid-mode.md
|-- scripts/
|   |-- crop_compare.py
|   |-- extract_color_components.py
|   |-- finalize_visio_copy_page.ps1
|   `-- visio_manual_redraw_scaffold.ps1
|-- requirements.txt
`-- LICENSE
```

## 环境要求

- Windows
- Microsoft Visio 桌面版
- PowerShell
- Python 3.10+
- `requirements.txt` 中的 Python 包

安装 Python 依赖：

```powershell
python -m pip install -r requirements.txt
```

## 安装为 Codex Skill

将本仓库 clone 到 Codex skills 目录：

```powershell
git clone https://github.com/zwj276765037-lab/Visio-copy.git "$env:USERPROFILE\.codex\skills\visio-copy"
```

然后在 Codex 中调用：

```text
$visio-copy
```

## 基本流程

1. 准备目标 `.vsdx` 文件和参考图片路径。
2. 基于 `scripts/visio_manual_redraw_scaffold.ps1` 创建项目专用绘图脚本。
3. 设置源图像宽度、高度和像素到 Visio 的比例映射。
4. 使用 `Add-RectPx`、`Add-LinePx`、`Add-TextPx`、`Add-PolygonPx` 等 helper 按像素坐标绘制。
5. 从 Visio 导出预览 PNG。
6. 对原图和预览图生成相同 bbox 的局部 crop。
7. 批量修复几何位置、文字、箭头、线宽、图层顺序和漏画元素。
8. 用户确认后清理底层参考图，只保留最终可编辑矢量层。

## 局部 Crop 对比

生成局部对比图：

```powershell
python scripts/crop_compare.py reference.png preview.png --out crops `
  --component left_stack:145,30,300,190 `
  --component right_table:560,40,220,120
```

组件格式：

```text
name:x,y,w,h
```

坐标均为参考图像中的像素坐标。

## 颜色区域提取

提取常见论文图颜色区域的粗略 bbox：

```powershell
python scripts/extract_color_components.py reference.png --min-area 100 --top 20
```

该工具只用于辅助定位。最终复刻质量仍然依赖人工组件审查和 Visio shape 级绘制。

## 最终清理

用户确认结果后，运行：

```powershell
powershell -ExecutionPolicy Bypass -File scripts/finalize_visio_copy_page.ps1 `
  -TargetPath "path\to\diagram.vsdx" `
  -PageName "VisioCopy_Trace" `
  -FinalLayerName "ManualRedraw_HiRes"
```

该脚本会备份文件，删除非最终层 shape，移除临摹底图和临时图层，并保存 `.vsdx`。

## 堆叠密集图形说明

不要把分离的小方块阵列画成一个整体立方体。应先从 crop 中统计可见单元数量，先画后层提示线，再画带不透明填充的前景单元，避免隐藏线穿过小方块缝隙。对于密集堆叠图，必须使用 2x 或 3x 局部 crop 检查。

更多规则见 `references/stacked-grid-mode.md`。

## License

MIT License. See `LICENSE`.
