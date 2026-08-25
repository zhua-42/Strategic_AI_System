# Deloitte Task 1 – Tableau 停机分析（Machine Down Time Analysis）学习笔记

**来源**：Deloitte Australia Data Analytics Job Simulation（Forage）Task 1 参考答案，见 `Deloitte_Forage\参考答案_Reference_Answers\Task1_Tableau停机分析_参考答案.md`
**类型**：Deloitte 案例
**一句话价值**：用 Tableau 从 16 万条工厂遥测数据中定位"故障最多的工厂和机器"，是零基础学 BI 看板 + 数据探索的绝佳小案例。

## 核心知识点

- **任务背景**：客户 Daikibo 把 4 家工厂（东京 Meiyo、大阪 Seiko、柏林 Berlin、深圳 Shenzhen）× 9 种设备的遥测数据合并为一个 JSON（2021 年 5 月，每台设备每 10 分钟发一条消息，共 160,704 条记录）。
- **要回答两个问题**：
  1. 哪个地点机器故障最多？（In which location did machines break the most?）
  2. 该地点哪些机器故障最频繁？
- **标准做法（Tableau 操作链）**：
  1. 导入 JSON（勾选所有 Schema 层级）；
  2. 建计算字段 `Unhealthy`：`IF [Status] = "Unhealthy" THEN 10 ELSE 0 END`（每条 unhealthy 计 10 分钟停机）；
  3. 建柱状图 "Down Time per Factory"（Factory × SUM(Unhealthy)）；
  4. 建柱状图 "Down Time per Device Type"（Device Type × SUM(Unhealthy)）；
  5. 合成 Dashboard，第一张图 **Use as Filter**，点击故障最多的工厂筛选第二张图；
  6. 选中结果截图上传。
- **方法论要点**：用"记录数 × 单位时长"把状态字段转成可量化指标；用仪表盘联动（filter）实现"先定位地点、再下钻设备"的两层分析。

## 学习要点

1. 原始 JSON 需先理解结构（每 10 分钟一条 → unhealthy 数 × 10 = 停机分钟），单位换算是第一道坑。
2. 计算字段把定性状态（Unhealthy）转成定量指标（分钟），是 Tableau 核心操作。
3. 仪表盘用 "Use as Filter" 实现下钻联动，比单图更有分析叙事感。
4. 结论要有"数量+单位+占比"支撑：Seiko 480 分钟（48 条记录），Laser Welder 占其 100%。
5. 数据要独立重算验证：4 家工厂 unhealthy 总数 48+42+11+2=103 条 × 10 = 1,030 分钟总停机，与设备维度汇总一致（480+430+70+20+10+10+10=1,030）。
6. 成品工作簿（.twb）可复用：官方步骤指南 PDF + 答案截图是"照着做"的完整教材。

## 实践建议

1. 动手在 Tableau 里从零做一遍：导入 JSON → 计算字段 → 两张图 → Dashboard 联动。
2. 用同一份 JSON 试其他维度（如按工厂×设备交叉看），练多维探索。
3. 提交前用"总数勾稽"校验：各工厂停机之和 = 各设备停机之和。

## 关键数据/答案

- 原始数据：160,704 条记录；unhealthy 记录数 × 10 分钟 = 停机时长
- **Q1 答案：Daikibo Factory Seiko（大阪）故障最多——48 条 unhealthy / 480 分钟停机**（深圳 42 条/420 分钟、东京 Meiyo 11 条/110 分钟、柏林 2 条/20 分钟）
- **Q2 答案：Seiko 内故障最频繁的设备是 Laser Welder（激光焊接机）——48 条 / 480 分钟，占该厂 100%**
- 全部工厂合计：LaserWelder 48/480、LaserCutter 43/430、HeavyDutyDrill 7/70、Furnace 2/20、SpotWelder 1/10、CNC 1/10、ConveyorBelt 1/10、MetalPress 0、AirWrench 0
