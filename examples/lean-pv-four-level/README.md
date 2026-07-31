# 光伏精益工具方法论 · 四级联动一页总览

适配硅片 / 组件（含广州组件基地）的精益工具系统化整合一页演示页。

## 交付物

| 文件 | 说明 |
|------|------|
| `output/slide-01.svg` | 1280×720 一页总览（对齐 ppt-agent 输出规范） |
| `output/index.html` | 交互预览页 |
| `output/精益工具方法论系统化整合_四级联动一页PPT.pptx` | PowerPoint 一页版 |
| `create_lean_ppt.py` | PPTX 生成脚本（python-pptx） |

## 页面结构

1. **核心目标**：提质、降损耗、提设备效率、均衡多基地产能、低碳降耗保交付
2. **五大集成模块**：价值流诊断 / 现场精益 / 品质管控 / 设备能耗 / 长效固化
3. **四级联动**：总部 → 基地（广州/珠海/西宁/宜宾）→ 车间 → 班组
4. **统一闭环**：诊断 → 拆解 → 落地 → 验证 → 固化 → 跨基地复制 → 持续迭代

## 预览

打开 `output/index.html`，或直接查看 `output/slide-01.svg`。

```bash
# 重新生成 PPTX（需 python-pptx）
pip install python-pptx
python create_lean_ppt.py
```
