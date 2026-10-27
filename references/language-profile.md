# 科技论文语言风格规范（计算力学 / 工程数值方法类期刊）

> 说明：这是从计算力学类学术稿件中抽象出来的**语言模式**，不是任何稿件的复写。风格目标是平实、客观、技术导向的英文，同时保证语法和逻辑准确。修改用户稿件时，以用户本人的文风为准；**不得把原稿的缺陷当成风格学习**。

## 1. 风格身份（六个核心特征）

1. **问题驱动、机制先行**：以具体的物理效应或数值困难开篇，指出经典模型的适用边界，再说明为什么需要所提出的理论或方法。少用空泛的应用愿景。
2. **客观陈述为主**：使用被动式或中性学术主语（`The governing equation ...`、`The numerical results ...`、`The method ...`）。可以使用 `we`，但不要每段都用。
3. **模型—方程—算法—验证的直线叙事**：先介绍物理量和假定，再列方程、边界条件、离散与求解方式，最后通过算例检查精度和行为。
4. **数学符号嵌入文字**：公式前交代其作用，公式后用 `where ...` 定义符号，再用 `According to Eq. (n), ...` 做下一步推导。避免"只有公式没有解释"或机械重复定义。
5. **证据对应结论**：数值部分明确几何、材料、边界、控制参数、对照标准与观测量。按"Figure/Table → 具体趋势 → 数值对比 → 原因/限制"写。
6. **低修辞密度**：以短到中等长度的完整陈述句为主，少用夸张形容词、堆叠的抽象名词或泛化的 AI 式总结；修改已有稿件时尽量少改。

## 2. 各章节写法

### Abstract
- 主线：**研究问题/限制 → 本文方法 → 技术实现 → 验证场景 → 有证据的结论**。
- 一般 150–220 词，或按期刊要求；不要在摘要中罗列章节，也不要堆砌无依据的 "significant"。
- 开头可用 `A ... method is developed for ...`、`This study presents ...`，全文只用一种。
- 方法要落到可观察的计算步骤；不要把它简单写成 "novel approach"。

### Introduction
- 先指出现象或需求，再简述传统理论和代表性方法，明确困难（例如高阶导数所需的连续性、非线性算子的求解、尺寸效应）。
- 最后的贡献段要回答：现有方法具体差在哪里、本文改了什么、可验证的收益是什么；不要泛称"提高精度/效率"。
- 文献按问题类型分组比较，不要罗列。找不到的引用记为 `[CITATION NEEDED]`，不得编造 DOI、年份、作者。
- 篇幅紧时章节简介压缩为 1–2 句；不必每次都写 `The remainder of this paper ...`。

### Governing equations / Method
- 顺序：**定义研究域与材料 → 守恒/能量/本构关系 → 边界条件 → 无量纲化/弱形式/离散 → 算法实现**。
- 公式前用一句话说明目的：`The constitutive relation is written as`。
- 公式后紧接定义：`where ... denotes ...`；缩写第一次出现时展开。
- 只有确实给出推导时才写 `It follows from ... that ...`；不要用措辞暗示未经证明的等价性。
- 避免 `Obviously`、`It is clear`、`It can easily be seen`；需要推理时写出关键条件。

### Numerical examples / Results
- 每个算例建议四段：①几何、材料、载荷与边界；②离散、收敛设置和参数；③结果趋势与比较（给出数值）；④成立条件与适用范围。
- 典型衔接：`Figure 3 shows ...` → `As ... increases, ...` → `The maximum difference ...`（仅在有数值时）→ `This behaviour can be attributed to ...`（有物理依据时）。
- 理论验证、收敛性、参数影响和工程应用分段写，不要把不同层面的结果混成"验证了有效性"。
- 写明横纵坐标、单位、观测路径和固定参数；没有数据就不写百分比或加速倍数。

### Conclusions
- 先用一句话说明完成了什么，再按**方法贡献 / 可核查的结果 / 适用条件与局限**写 2–4 条事实陈述。
- 不重复摘要的全部细节，不声称未经验证的工程价值；对待验证的推断加限定语。

## 3. 可用的句式模块（填入真实事实，不要照抄例句）

| 作用 | 句式 | 注意 |
| --- | --- | --- |
| 引入物理问题 | `The response of [system] depends on [two or three named factors].` | 不写 "plays an increasingly important role" |
| 传统模型局限 | `The classical [model] does not account for [specific phenomenon].` | 无依据时用 `may be insufficient` |
| 研究目的 | `In this study, [method] is employed to investigate [quantity] under [conditions].` | 主语与作用明确 |
| 方法实现 | `The domain is divided into [subdomains], and [field] is approximated using [basis].` | 只描述实际采用的方法 |
| 引出公式 | `The weak form of the governing equation is written as` | 前后不重复 |
| 符号解释 | `where [symbol] denotes [quantity] and [symbol] is [parameter].` | 新符号首次出现都要解释 |
| 对比 | `The results are compared with [analytical/reference solution] in Figure [n].` | 对照来源必须真实 |
| 趋势 | `The [quantity] decreases with increasing [parameter] over the range [values].` | 范围和量纲正确 |
| 收敛 | `The relative error decreases as the number of nodes increases.` | 只有数据确实如此才写 |
| 不确定性 | `Within the considered parameter range, the influence of [factor] remains limited.` | 限定范围 |
| 结论 | `The formulation provides [specific capability] without [specific operation].` | 不写成无条件高效 |

## 4. 示意句（新写，不对应任何数据）

**Theory**
> The displacement and its spatial derivatives are taken as the primary variables. The resulting boundary-value problem involves additional material coefficients associated with an internal length.

**Numerical setup**
> A square plate containing a circular hole is considered. The hole radius is varied while the material parameters are kept fixed.

**Results（无数值时只作描述）**
> Figure 4 illustrates the change in the stress distribution as the hole radius increases. The difference between the two models is most pronounced near the boundary of the hole.

**Limits**
> These results are obtained for the selected boundary conditions and geometries. Further calculations are required before the conclusion is extended to general three-dimensional configurations.

## 5. 不要学习的草稿缺点

- 协作草稿标记（如 "this section is to be rewritten"）不得出现在新稿中。
- Word 中显示为 `Equation Chapter ...`、`MERGEFORMAT` 的是损坏的域代码，不是编号格式，不要复用。
- `significantly`、`excellent agreement`、`robust`、`novel` 等缺乏定量支撑时，删除或弱化，或补上误差与收敛指标。
- 主谓不一致、缺冠词、`it is requirement to ...` 一类的语法错误要改正，不能照搬。
- 不要连续使用 `Moreover`、`Furthermore`、`It is worth noting that`；按逻辑直接陈述。
- 结论不要重复摘要结构；不要把近似和数值验证包装成严格证明。
- 修改已有英文稿时**尽量少改**：保留科学含义、符号、章节结构、数值和引文位置，优先改语法、时态、搭配和逻辑连接。

## 6. 术语与表达偏好

- 以 British English 为基线（`modelling`、`normalised`、`behaviour`）；期刊要求美式拼写时全文统一替换，不要混用。
- 缩写首次出现时展开，如 `boundary-value problem (BVP)`、`finite element method (FEM)`，之后用简称。
- 常用准确术语：`governing equation`、`constitutive relation`、`boundary-value problem`、`effective coefficients`、`convergence`、`relative error`。按具体领域选用，不要为了贴近样稿而硬加术语。
- 正文写 `Eq. (n)`、`Figure n`、`Table n`（或期刊规定的形式）；标题一般用 sentence case。
- 报告数值时，变量、误差定义、有效位数、单位和参考解要前后一致；没有原始数据时不推断显著性。

## 7. 交付前的语言自查

1. 每段主题句是否具体？有没有不含科学内容的空泛句？
2. 每个强形容词后面有没有量化证据？
3. 每条结论能否追溯到推导、实验、图或表？
4. 首次出现的符号、缩写和边界条件是否都解释了？
5. 算法描述是否区分了 `computed to machine precision` 和不恰当的 `exact solution`？
6. 观点、分析和文献事实是否分开？有没有未经核实的引用？
7. 修改已有稿件时，是否最大限度保留了原作者的逻辑、句长和术语？
