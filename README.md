# Editable Scientific PPT

把科研图制作成原生可编辑 PowerPoint：文字是真正的文本框，曲线和网格是可编辑形状，子图边界在导出前完成几何裁切。适用于 SVG 转形状后出现文字堆叠、字体变化和跨子图长线的问题。

**版本：0.1.0；作者：Tanwu2000；许可证：MIT。** 不包含论文数据、Abaqus 模型、个人路径、凭据或私有运行库。

## 调用

安装后在 Codex 中输入：

> 使用 $editable-scientific-ppt，把这张科研图做成原生可编辑 PPTX。文字使用 Arial，可修改字号；网格与曲线可修改线宽；保留原布局，裁切越界线条，同时输出 PNG 和 600 dpi TIFF。

也可以自然语言请求“可编辑的 PPT 科研图”。自动匹配取决于宿主是否支持技能发现。

## 安装

将**整个仓库文件夹**复制或克隆为个人技能目录下的 `editable-scientific-ppt`，确保该目录直接包含 `SKILL.md`。此用户环境的技能目录是 `~/.codex/skills`；不同宿主可按自己的技能发现规则选择目录。不要把仓库父目录当成技能目录。

GitHub 发布后的仓库地址为 `https://github.com/Tanwu2000/editable-scientific-ppt`。可让 Codex 的 `$skill-installer` 从该地址安装，或使用官方安装脚本的 `--repo Tanwu2000/editable-scientific-ppt --path . --name editable-scientific-ppt`。已有同名目录时，官方安装器会拒绝覆盖。

本地复制安装：

```sh
python tools/install_local.py --dest ~/.codex/skills
```

安装后在下一次请求中调用；若宿主尚未发现，可重新打开会话或重启宿主。

## 更新会自动同步吗？

通常不会。Skill 是文件，发布 GitHub 更新不会直接改写别人已经安装的本地副本，也不会修改以前生成的 PPTX。

| 安装方式 | 更新方式 |
|---|---|
| 官方安装器或复制安装 | 先保留本地修改和旧版本，再安装新版本 |
| Git 克隆进技能目录 | 在该目录执行 `git pull --ff-only`；有本地改动先检查，不强制覆盖 |
| 固定 tag/commit | 只有主动切换版本才更新，适合可复现科研流程 |
| 托管插件 | 取决于该平台的分发与更新机制，本仓库不承诺自动更新 |

本仓库的本地安装工具默认不覆盖。对复制安装进行更新时，从已下载/审阅的新版本执行：

```sh
python tools/install_local.py --dest ~/.codex/skills --update
```

工具会先把旧技能移动到技能目录旁的备份目录，再安装新版本，失败时回滚。不联网、不自动拉取、不定时更新。Git 安装请通过 Git 更新，不用此工具替换。

## 脚本与能力边界

- `scripts/native_pptx.py`：从显式 JSON 场景生成单页原生 PPTX，使用 Python 标准库。
- `scripts/inspect_pptx.py`：结构审查，统计原生文本、自由曲线与图片。
- `examples/minimal_scene.json`：合成示例，不代表实验或有限元结果。
- `references/`：几何、文字、场景格式及 PowerPoint 验证说明。

这不是任意 SVG/PDF 一键完美转换器。复杂公式、蒙版、渐变、带孔多边形等需要其他合适的后端或从源绘图代码重建。云图可以是图像，其轴标签和图例文字仍可独立编辑。文字编辑能力、形状编辑能力和数据表驱动的图表编辑能力是不同概念。

## 验证

```sh
python -m unittest discover -s tests -v
python scripts/native_pptx.py examples/minimal_scene.json build/example.pptx
python scripts/inspect_pptx.py build/example.pptx
```

结构验证不能替代 PowerPoint 视觉与编辑验证，见 `references/powerpoint-qa.md`。

## 发布与维护

后续修改先更新 `VERSION` 和 `SKILL.md` 的版本，运行测试，再提交并打 tag（如 `v0.1.1`）。保持核心技能简短，新增已验证的规则与案例，避免把某一张图的偏好写成所有图都必须遵守的规则。

参考：[官方技能说明](https://learn.chatgpt.com/docs/build-skills)；[官方安装器](https://github.com/openai/skills/tree/main/skills/.system/skill-installer)。
