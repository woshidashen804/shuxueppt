---
name: shuxueppt
description: 生成与检修中文数学教学 PPT(.pptx)。覆盖两种任务：从教案/练习册批量生产数学课件，以及对既有 18 份级别课件做质量检查与修复（内容缺失、lim 极限表示、分式、指数上标、答案与题干重叠、排版越界）。内置纯 XML 检查器、批量修复脚本、PowerPoint/WPS 双引擎渲染比对与回归验收，重点解决"PowerPoint 正常但 WPS 里公式和文字整块消失"这类兼容性缺陷。
---

# 数学教学 PPT 生成与检修

面向中文数学课件（导数、极限、积分等）的生产与检修。核心判断标准只有一条：**同一份 pptx 在 PowerPoint 和 WPS 里都要显示完整、公式结构正确、答案不压题干**。只看 PowerPoint 通过不算完成。

## 何时用这个技能

- 要按课次批量生成数学课件（一课一份，含导入/概念/例题/随堂练习/小结）。
- 已有课件"内容缺失""极限写得不对""分式是斜杠""指数掉下来了""答案压住题目""排版乱"。
- 需要交付前做一次全量体检和回归验证。

## 第一原则：先量，再改

不要凭肉眼翻页判断。先用检查器把所有缺陷量化成清单，按缺陷族归类，再逐族修。改完必须回归验收，证明"该修的都修了，没修的没坏"。

```bash
python "<SKILL_DIR>/scripts/inspect_pptx.py" --decks "<课件目录>" --out "<输出目录>"
```

产出 `issues.json`、`report.md`（人读）。先把这个清单交给用户确认优先级，再动手。

## 缺陷族与判定标准

按严重度排序，完整判定规则与真实样例见 [缺陷目录](references/defect-catalogue.md)。

| 族 | 代码 | 现象 | 根因 |
|---|---|---|---|
| 公式结构非法 | `M-BADNEST` | WPS 里公式整块消失，PowerPoint 正常 | `m:oMathPara` 里嵌了 `a14:m` |
| 内容为空 | `M-EMPTYMATH` / `M-FUNCARG` / `M-EMPTYFRAC` | 公式空白、`sin` 无自变量、分式分子分母空 | 转换时结构丢失 |
| 极限 | `L-AS-SUB` | `lim` 的趋近条件缩成右下角小脚标 | 用 `m:sSub` 而非 `m:limLow` |
| 指数 | `E-SPLITRUN` | `x3/2`、`e−x` 平排 | 指数被切成独立 run，靠 `baseline` 硬抬 |
| 分式 | `F-LINEAR` / `F-IN-TEXT` | `1/x`、`π/2` 斜杠平写 | 该做分式对象的地方留了 `/` |
| 版式 | `O-COLLIDE` / `O-OFFCANVAS` | 答案压题干、形状越出画布 | 坐标错位 |
| 潜在兼容风险 | `M-BLANKFB` | **对 WPS 无实测影响** | 见下方警告 |

**警告：不要假设"空 Fallback 导致 WPS 缺内容"。** 实测 5 份课件、单变量对照，展开空 Fallback 对 WPS 渲染的墨迹变化恒为 `+0`；同一批课件的改善 100% 来自拆 `oMathPara` 非法嵌套。WPS 实际读的是 Choice 分支。空 Fallback 仍是其他阅读器的真实风险，但**不是 WPS 症状的主因**，不要用它解释用户报告。数据见 [缺陷目录](references/defect-catalogue.md)。

**必须做单变量实验再下结论。** 只改一处结构、其余不动，两个引擎各渲染同一页，量该公式所在行的墨迹。不做这一步就改代码，很容易修错地方、还破坏了原本正常的画面。

**不要误判的两类：**
- 答案形状与题干完全重叠，有时是**点击揭示**的有意设计（答案出现时盖住题干）。改之前先看动画时间轴：两者都是 `entr` 且位置一致，通常是设计。仅当答案偏到别处、留下 2 英寸空隙或压住**另一道**题时才是缺陷。
- `a14:m` 数量减少不一定是内容丢失——展开非法嵌套会让它合法减少。**永远用 `<m:oMath[ >]` 计数做不变量**，不要用 `a14:m`。

## 修复顺序

由"看不见"到"不好看"，且每一步都要过数学计数守护。

1. **公式结构非法** → `scripts/fix_invisible.py --nest`（拆 `oMathPara` 内的 `a14:m`）**← 真正修复 WPS 缺内容的一步**
2. **极限表示** → `scripts/fix_limit.py`（`m:sSub` 重建为 `m:limLow`）
3. **空结构** → `scripts/fix_empty_math.py`（空 `m:func` 外壳、空分式）
4. **指数** → `scripts/fix_exponent.py`（拆散 run 合成真 `m:sSup`）
5. **版式** → `scripts/audit_overlap.py` 定位，再手工或改坐标挪位
6. **可选** → `scripts/fix_invisible.py --unwrap`（展开空 Fallback，仅当目标读者可能是老版本 Office 等不支持 a14 的软件时）

所有修复脚本默认 **dry-run**，加 `--write` 才落盘；每个脚本内置"每页 `<m:oMath` 数不得减少"断言，触发即中止。

细节配方与代码要点见 [修复配方](references/repair-recipes.md)。

## 生成新课件

从教案/练习册产出时的硬性要求见 [设计规范](references/design-spec.md)，要点：

- 一课一份，页序固定：导入 → 概念 → 例题 → 随堂练习 → 小结。
- 练习页题号必须与练习册一致；整题略讲要在页面上写明"第 X 题与第 Y 题相同，略"。
- 答案用独立形状 + `entr` 点击动画，初始 `hidden="1"`，动画触发后显示。
- 公式一律真 OMML，不用 Unicode 上标字符拼凑，不用 `/` 写分式，不用文字写 `lim`。
- 颜色、字号、卡片尺寸沿用同套课件既有值，不新造风格。

## 验收（必须做，不可跳过）

```bash
# 1. 结构合法 + 数学不丢失（与备份逐页比 oMath 计数）
python "<SKILL_DIR>/scripts/verify_integrity.py" --decks "<课件目录>" --baseline "<备份目录>"

# 2. 双引擎渲染比对（需要 Windows + PowerPoint + WPS）
python "<SKILL_DIR>/scripts/compare_render.py" --decks "<课件目录>" --baseline "<备份目录>" \
       --engines ppt,wps --out "<输出目录>"
```

通过标准：

| 指标 | 要求 |
|---|---|
| PowerPoint 逐页墨迹变化 | 0 页超过 2%（证明没有改坏原有画面） |
| WPS 逐页墨迹 | 修复后 ≥ 修复前（证明内容真的回来了） |
| 每页 `<m:oMath` 计数 | 不低于基线 |
| 文本重叠 | 无"答案压住另一道题"的实例 |

`compare_render.py` 依赖 `win32com` 与 `Pillow`；无 Windows 环境时至少跑 `verify_integrity.py` 的纯 XML 校验，并明确告知用户渲染未验证。

**如果 WPS 侧墨迹没有上升，说明没修到根因。** 回到单变量实验，用 `build_variant.py` + `render_ink.py` 分别验证每个假设，不要继续叠加改动。

## 做单变量实验（本技能的核心方法）

任何"改了 X 就能修好 Y"的判断都必须先验证：

```bash
# 从备份构造只含单一变换的变体
python "<SKILL_DIR>/scripts/build_variant.py" "<备份.pptx>" "<输出.pptx>" u   # 仅展开 Fallback
python "<SKILL_DIR>/scripts/build_variant.py" "<备份.pptx>" "<输出.pptx>" f   # 仅拆嵌套
python "<SKILL_DIR>/scripts/build_variant.py" "<备份.pptx>" "<输出.pptx>" uf  # 两者

# 每引擎单进程渲染，避免 COM 冲突
python "<SKILL_DIR>/scripts/render_ink.py" KWPP.Application "<文件.pptx>" "<out.json>"
python "<SKILL_DIR>/scripts/render_ink.py" PowerPoint.Application "<文件.pptx>" "<out.json>"
```

判据：目标引擎墨迹上升、另一引擎墨迹不变的那个变体，才是真正的修复。

**踩过的坑**：曾判定"空 Fallback 导致 WPS 缺内容"，理由是看到 1265 处可疑结构且 WPS 确实缺内容。实测后发现展开 Fallback 的墨迹变化恒为 `+0`，真正原因是同批文件里 207 处 `oMathPara` 非法嵌套。结构可疑不等于它导致问题——这个错误浪费了一整轮改动。

## 交付

- 修改前的完整备份目录（含时间戳），并在交付说明里给出回滚方法。
- `report.md`：修复前后对比、逐页清单、遗留项及原因。
- 不要声称"已验证"除非真的跑过对应引擎。

## 参考

- [缺陷目录](references/defect-catalogue.md) — 每个缺陷族的判定式、真实样例、误判陷阱
- [修复配方](references/repair-recipes.md) — 每类修复的 XML 前后对照与实现要点
- [设计规范](references/design-spec.md) — 生成新课件时的页面结构、命名、动画、公式要求
- [OOXML 结构速查](references/ooxml-notes.md) — 公式容器层级、AlternateContent 语义、验证清单
