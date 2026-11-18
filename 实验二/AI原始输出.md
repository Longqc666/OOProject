# DeepSeek 原始输出记录（活动 3）

- 时间：2026-11-18
- 工具：DeepSeek（网页版）
- 提示词：V3（全文见报告 12.3 节）
- 说明：以下保留网页版回复的正文、两套自评、规则落位表与取舍建议；三个 PlantUML 代码块已原样转存为
  `12_AI_arch_A_raw.puml`、`13_AI_arch_B_raw.puml`、`14_AI_class_raw.puml`，未改字符，未做语义修改。

---

## 一、两个备选体系结构方案

### 方案一：模块化单体 + 端口适配器（Modular Monolith / Hexagonal）

**一句话定位**：以“借阅流通”为单一可部署单元，内部按领域模块强边界划分，ILS 通过适配器端口隔离，
规则版本化内置，适合小团队无运维场景。

**自评**

| 维度 | 分数 | 理由 |
|---|---|---|
| 可维护性 | 5 | 单一部署单元、模块内强边界、端口隔离 ILS，3～5 人团队无需分布式调试；规则与审计同库同事务，C4 天然满足 |
| 可扩展性 | 3 | 模块可后续拆分为服务，但横向扩展粒度粗；借阅高峰需整体扩容 |
| 需求匹配度 | 5 | 实时续借（C3）由本地事务保证；规则版本化内置；ILS 故障可在端口层降级；无运维负担 |

**不适用的场景**：借阅流通 QPS 极高需要按模块独立弹性伸缩；团队规模扩大后多团队并行开发同一单体产生合并冲突；
未来需要把预约/规则独立对外提供服务能力时，模块边界会被迫提前拆分。

### 方案二：领域服务拆分（借阅核心 + 规则服务 + ILS 网关服务）

**一句话定位**：把“续借决策”和“ILS 集成”拆成独立服务，借阅核心保持实时强一致，规则与 ILS 网关作为
可独立演进/降级的支撑服务。

**自评**

| 维度 | 分数 | 理由 |
|---|---|---|
| 可维护性 | 3 | 三个服务边界清晰，但 3～5 人无专职运维要承担多服务部署、联调、版本兼容成本；审计同源需跨服务设计（审计写回借阅核心同库） |
| 可扩展性 | 5 | 规则服务可独立发版/灰度；ILS 网关可独立扩缩容与限流；借阅核心可横向扩展 |
| 需求匹配度 | 4 | C3 实时性靠借阅核心本地事务 + 同步调用保证；C2 幂等与对账集中在网关服务更专业；但跨服务同步调用增加 P95 风险，需熔断降级设计 |

**不适用的场景**：团队 3～5 人且无运维，多服务带来的部署/监控/链路追踪负担超过收益；ILS 调用量不大时，
独立网关服务属于过度拆分；规则变更频率低时，规则服务独立部署价值有限。

---

## 二、关键类图

覆盖 Patron、BibliographicRecord、Item、Loan、Reservation、RenewRuleSet、RuleEvaluator、RenewalApplication、
RuleDecision；关系与多重性见 `14_AI_class_raw.puml`。要点：

- Patron 1 — 0..* Loan：一个读者多笔借阅；
- BibliographicRecord 1 — 1..* Item：书目包含副本；
- Item 1 — 0..* Loan：副本有多条历史借阅；
- Item 1 — 0..* Reservation：副本被多人预约；
- Loan 1 — 0..1 RenewalApplication：借阅可有续借申请；
- RenewRuleSet 1 — 0..* RenewalApplication：规则版本被申请引用；
- RenewalApplication 1 *— 1 RuleDecision：申请持有判定结果。

---

## 三、业务规则落位表

| 规则 | 含义 | 落位类 / 方法 / 校验器 | 说明 |
|---|---|---|---|
| BR-01 | 学生最多 2 次、教师 5 次 | `RenewRuleSet.maxRenewCount(patronType)`，由 `RuleEvaluator.evaluate` 调用；`RenewalApplication.ruleVersionId` 记录版本 | 次数上限按读者类型配置化、版本化，判定时绑定规则版本 |
| BR-02 | 每次顺延 30 天 | `RenewRuleSet.renewPeriodDays()` → `RuleEvaluator` 计算 `newDueDate` → `Loan.renew(newDueDate)` | 顺延天数来自规则版本，写入 `Loan.dueDate` 与 `RenewalApplication.newDueDate` |
| BR-03 | 有他人预约不可续借 | `RuleEvaluator.evaluate(hasReservation)` 结合 `Reservation` 查询；`Item 1—0..* Reservation` 提供事实 | 借阅核心查该副本预约队列，作为判定输入；不改变队列（BR-05） |
| BR-04 | 已逾期不可续借 | `Loan.isOverdue(asOf)` + `RenewRuleSet.allowRenewWhenOverdue()` 由 `RuleEvaluator` 判定 | 逾期事实来自借阅记录，策略是否允许由规则版本决定 |
| BR-05 | 续借不改变预约队列位次 | `Reservation.queuePosition` 只读；续借流程不调用任何 `Reservation` 变更方法 | 续借仅更新 `Loan` 与写 `RenewalApplication`，预约聚合保持不动 |

---

## 四、AI 的取舍建议（原文）

- 团队 3～5 人、无运维、ILS 调用量中等、规则变更不频繁 → 选方案一，用端口适配器把 ILS 与规则隔离，
  C2/C3/C4/C5 都能在单体内用本地事务 + 版本化规则 + 同库审计满足。
- 若规则需要频繁灰度、ILS 调用需要独立限流熔断、借阅核心要独立扩缩容 → 选方案二，但需接受多服务运维成本，
  并把审计同源设计为写回借阅核心同库。

---

## 五、处理记录

- 语法修正：0 处（三份代码首轮渲染成功）。
- 语义修改：0 处（未调整组件、类、关系）。
- 代码块：三个，分别对应方案一组件图、方案二组件图、关键类图。
