# PowerPoint 实测案例

四个案例使用确定性的合成数据，测试文字、几何裁切和编辑能力，不代表实验或有限元计算结果。误差棒是用于检验图形元素的合成宽度，不表示置信区间或样本统计。

## 比较方法

1. 用 `build_cases.py` 绘制源图，导出 PNG 和保留文字的 SVG（`svg.fonttype='none'`）。
2. **普通路径**：把同一 SVG 插入 PowerPoint，执行“转为形状”，直接导出幻灯片 PNG。
3. **Skill 路径**：解析这组源图的文字与路径，预先裁切几何，生成原生文本框和自由曲线，再由 PowerPoint 导出 PNG。
4. 将两份 PNG 按同一比例并排排版，不修改图内像素。另存原图和转换前的 PowerPoint 渲染，便于核对。

左右差异来自导入、转换与重建路径，不是更换了数据或故意把左图画坏。部分普通转换结果基本正常；这组测试也不代表所有 Office 版本的表现。

## 案例与文件

| 案例 | 本次观察 | 原生文本框 | 原生矢量形状 | 位图 |
|---|---|---:|---:|---:|
| [双对数曲线](results/01_math/) | 符号字体、上下标和数学标签排版变化 | 17 | 85 | 0 |
| [网格局部放大](results/02_mesh/) | 转形状后网格越过子图边界 | 29 | 34 | 0 |
| [云图与色标](results/03_field/) | 单位的混合文本发生重叠 | 23 | 29 | 2 |
| [误差棒与图例](results/04_errorbars/) | 主体曲线保留，混合文本间距变化 | 18 | 136 | 0 |

原生形状数量包含坐标轴、刻度等；网格按层合并，数量不等于网格单元数。云图的两张位图分别为色彩图和色条。

每个目录包含：

| 文件 | 内容 |
|---|---|
| `source.svg` / `source.png` | Matplotlib 原始输出 |
| `svg-inserted.png` | SVG 插入 PowerPoint、尚未转形状的实际显示 |
| `svg-converted.pptx` / `.png` | 普通“转为形状”后的文件及渲染 |
| `native.pptx` / `.png` | 原生可编辑版本及 PowerPoint 渲染 |
| `native-edited.pptx` / `.png` | 修改字号和线宽、保存并重开后的文件及渲染 |
| `comparison.png` | 首页使用的并排对比图 |
| `native_geometry_check.json` | 裁切后的几何边界检查 |

公开 PPTX 仅清空了 Office 作者档案字段；幻灯片 XML 和内容保持不变。

## 编辑验证

四个原生文件分别修改一个标题为 **14 pt**，并将一条曲线或一个网格/误差棒层的线宽改为 **2.5 pt**。保存副本、关闭并重新打开后，再读取属性并导出图片。四例均通过。修改对象名称和结构统计记录在 [report.json](results/report.json)。

这验证了文字和形状可以直接修改，不需要“转为形状”。它不等于每个对象、每种字号和所有 PowerPoint 版本都已测试。

## 复现

实测日期：2026-11-22。Windows 桌面版 PowerPoint 16.0，Office Build 16.0.20527.20196；Python 3.9.13，Matplotlib 3.9.2，NumPy 2.0.2，lxml 4.9.1，Pillow 9.2.0。常规文字为 Arial，比例符号使用 Cambria Math；字体不随仓库分发。

在仓库根目录运行，输出目录须为新目录：

```sh
python -m pip install -r examples/gallery/requirements.txt
python examples/gallery/build_cases.py --output build/gallery-run
```

在装有 PowerPoint 的 Windows 桌面会话中运行：

```powershell
powershell -File examples/gallery/render_cases.ps1 -Artifacts build/gallery-run
```

脚本会打开专用测试演示文稿。“转为形状”可能弹出大量形状的性能提示，选择“是”后继续。脚本只关闭自己创建或打开的测试文件，不退出 PowerPoint。未安装 PowerPoint 时，可生成文件，但不能复现本页的实际 Office 渲染和编辑验证。

生成比较图及可分享文件（同样使用新输出目录）：

```sh
python examples/gallery/publish_results.py build/gallery-run --output build/gallery-public
```

`matplotlib_scene.py` **只适配这四个 Matplotlib 示例**，不是通用 SVG 转换器。它处理这里使用的矩形裁切、文字、路径和内嵌图像；其他输入应从绘图源代码重建，或另行验证转换。曲线形状不附带 Excel 数据表。

核心导出器不需要本目录的第三方依赖。使用显式场景格式时，参见[场景接口](../../references/scene-format.md)。
