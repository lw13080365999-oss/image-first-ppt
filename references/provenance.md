# 来源与复用范围

研究日期：2026-10-31。按优秀通用 PPT skills 的内容、结构、布局、风格和检查价值选样，未按是否使用图片路线筛选。这个技能重新组织这些工作方法，实际制作统一采用图片模型设计整页与可编辑对象重建。运行时不需要下载或安装以下项目。

本技能的脚本和八种风格标本为本次新编写；参考文档是结合用户要求重新组织的指导。原始研究快照保存在开发资料目录，未将上游大仓库或付费后端打包入技能。

| 来源与固定版本 | 本技能参考的部分 | 许可 |
|---|---|---|
| [hugohe3/ppt-master](https://github.com/hugohe3/ppt-master/tree/5e8856be0da8e6bc85c4aa28acc614478f7b7137) | 沟通任务、论述模式、语义关系、整套节奏和具体风格 | [MIT / Hugo He](licenses/ppt-master-MIT.txt) |
| [JimLiu/baoyu-skills: slide-deck](https://github.com/JimLiu/baoyu-skills/tree/1567581c26ec29f4216c6e6835415bf30343b0e3/skills/baoyu-slide-deck) | 内容分析、布局词汇、维度化风格、全局与逐页提示分离 | [MIT / Jim Liu](licenses/baoyu-skills-MIT.txt) |
| [ningzimu/codex-ppt-skill](https://github.com/ningzimu/codex-ppt-skill/tree/941dcafe2f23ea4ae61de99e9b22fb0ef3d9e06c) | 上下文、逐页任务、严格素材与风格参考区分、状态与局部修改 | [MIT / ningzimu](licenses/codex-ppt-MIT.txt) |
| [siril9/presentation-skill](https://github.com/siril9/presentation-skill/tree/efd899c5ddf866af2ce0001c4c9340af775d4d55) | 内容/证据/素材分工、实际渲染审阅和产物对应 | [MIT / Siril Sengolraj](licenses/presentation-skill-MIT.txt) |
| [OpenAI historical slides](https://github.com/openai/skills/tree/e6afb0d74cc75d220df2faf3dd6c635c2dc6a108/skills/.curated/slides) | 字体、画布、渲染、总览和越界检查思路 | [Apache-2.0](licenses/openai-slides-Apache-2.0.txt) |

上表许可证原文随包保留。没有复制上游整页图装配器、SVG生产器、视觉审核脚本或品牌资产；本技能辅助脚本为原创实现，也不代表上游项目提供支持。没有采用上游的强制问卷、审批、广告、固定模型切换或“必须找到错误”等产品规则。

Anthropic 的 `pptx` 作为功能覆盖比较样本，但其独立许可为 Proprietary；其提示词、代码和素材未并入本技能。OpenAI slides 指删除前的历史版本，不是当前 main 中仍提供的技能。不要将研究到的历史依赖当成当前环境已安装能力。

八种风格示意使用虚构内容和数据，由 SVG 绘制，用于比较设计参数；不是模型生成结果，也不是证据材料。实际 PPT 仍须经过真正的图片模型设计阶段。
