# Deloitte Task 2 – Excel 薪酬平等分类（Equality Classification）学习笔记

**来源**：Deloitte Australia Data Analytics Job Simulation（Forage）Task 2 参考答案，见 `Deloitte_Forage\参考答案_Reference_Answers\Task2_Excel平等分类_参考答案.md`
**类型**：Deloitte 案例
**一句话价值**：一个 5 分钟就能做完的 Excel 分类公式题，却精准训练了"读懂评分规则 → 写对嵌套 IF → 用 ABS 处理正负值"的入门数据清洗与分类技能。

## 核心知识点

- **任务背景**：Daikibo 因内部薪酬性别不平等投诉，法证技术团队用算法算出各工厂各岗位的 **Equality Score**（整数，-100 ~ +100，0 为最理想），给出 `Equality Table.xlsx`（3 列：Factory / Job Role / Equality Score）。
- **要求**：新增第 4 列 `Equality class`，按规则分类：
  | 类别 | 规则 |
  |---|---|
  | Fair | 分数在 ±10 之间（含） |
  | Unfair | 分数 < -10 或 > 10 |
  | Highly Discriminative | 分数 < -20 或 > 20 |
- **标准答案公式**（D2 单元格，下拉填充到全部 37 行）：
  ```
  =IF(ABS(C2)>20, "Highly Discriminative", IF(ABS(C2)>10, "Unfair", "Fair"))
  ```
  逻辑：`ABS(分数)>20` → Highly Discriminative；`ABS(分数)>10` → Unfair；其余（|分数|≤10）→ Fair。
- **关键陷阱**：官方任务示例里写 "-9 → Unfair"，但按公式 |-9|=9 ≤ 10 应判 **Fair**——这是平台示例的笔误；所有公开通过的解答均用上面的 ABS 公式（Forage 官方示例答案即此公式）。

## 学习要点

1. 涉及正负对称的评分分类，优先用 **ABS()** 一次处理两个方向，避免写死正负分支。
2. 嵌套 IF 顺序要"从苛刻到宽松"：先判 ±20，再判 ±10，最后兜底 Fair，否则逻辑会被吞掉。
3. 动手前先读懂评分规则本身（阈值含不含边界：±10 含在内为 Fair）。
4. 提交时保留公式版（官方要求的格式），另存数值版便于核对——"底稿与成品分离"是好习惯。
5. 对平台示例与规则矛盾时，以规则/公式为准，并在说明中标注该疑点（职业素养体现）。
6. 分类结果要能自检：37 行 = 19+11+7，分类列无空值、无非法值。

## 实践建议

1. 在 Excel 里重做一遍并下拉填充，验证分类分布：Fair 19 / Unfair 11 / Highly Discriminative 7。
2. 试着用 `IF(ABS(...)...)` 之外的方法（如 IFS 或 lookup 表）实现同结果，比较可读性。
3. 把"ABS + 嵌套 IF 阈值分类"存为模板，处理任何评分/阈值分档任务都可复用。

## 关键数据/答案

- 公式：`=IF(ABS(C2)>20,"Highly Discriminative",IF(ABS(C2)>10,"Unfair","Fair"))`
- 结果分布（37 行）：**Fair 19、Unfair 11、Highly Discriminative 7**
- 抽样核对：Meiyo C-Level -25 → Highly Discriminative；Meiyo Director -19 → Unfair；Meiyo Manager -14 → Unfair
- 示例笔误说明：官方示例 "-9 → Unfair" 与公式矛盾，按公式 -9（|-9|=9≤10）应为 Fair
- 提交文件：`Task2_Equality_Table_公式版.xlsx`（推荐）/ 数值版
