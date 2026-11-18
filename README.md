# 面向对象分析与设计 · 课程实验仓库

学号：32406300076　姓名：庞华杰　班级：软件工程 243

## 实验一　需求分析

选题：图书馆智能借阅助手（AI 原生应用 · 智能体系统）
建模范围：借阅流通子系统　核心用例：UC-03 续借图书

### 目录结构

```
实验一/
├── 32406300076.pdf                 # 实验报告（提交件）
├── 实验一报告_32406300076.docx      # 报告源文件
├── puml/                            # PlantUML 模型源文件
│   ├── 01_usecase_borrowing.puml    # 图 1 用例图（手工）
│   ├── 02_class_renew.puml          # 图 2 类图（手工）
│   ├── 03_state_loanrecord.puml     # 图 3 状态图（手工）
│   ├── 04_sequence_renew.puml       # 图 4 顺序图（手工）
│   ├── 05_AI_usecase_raw.puml       # 图 5 AI 原始输出（未修改）
│   └── 06_AI_class_raw.puml         # 图 6 AI 原始输出（未修改）
└── img/                             # 渲染图（由 puml 生成，与源文件同源）
```

### 渲染方式

需要 JDK（PlantUML 依赖）与 plantuml.jar：

```
java -jar plantuml.jar -tpng -charset UTF-8 -o img puml/*.puml
```

类图使用 PlantUML 内置 Smetana 布局引擎（`!pragma layout smetana`），无需安装 Graphviz。

### 提交时序

按实验指导书“提交时序边界”的要求，手工模型（commit 1）先于 AI 辅助环节（commit 2）入库，
报告与 PDF 最后提交（commit 3）。

## 实验二　系统设计

选题延续实验一：图书馆智能借阅助手（AI 原生应用 · 智能体系统）
设计范围：借阅流通子系统　核心用例：UC-03 续借图书
设计主线：模块化单体 + 六边形（端口-适配器）；同步核心 + Outbox 异步旁路

### 目录结构

```
实验二/
├── 32406300076.pdf                 # 实验报告（提交件）
├── 实验二报告_32406300076.docx      # 报告源文件
├── AI原始输出.md                    # DeepSeek 原始输出记录（活动 3）
├── build_tpl.py                     # 按学校模板生成报告的脚本
├── puml/                            # PlantUML 模型源文件
│   ├── 01_architecture_layered.puml # 图 1 分层体系结构
│   ├── 02_domain_class.puml         # 图 2 领域类图
│   ├── 03_app_class.puml            # 图 3 应用类图（BCE + 校验链）
│   ├── 04_state_loanrecord_design.puml  # 图 4 借阅记录状态图
│   ├── 05_state_renewal_request.puml    # 图 5 续借申请状态图
│   ├── 06_sequence_renew_design.puml    # 图 6 设计级顺序图
│   ├── 07_ui_prototype.puml         # 图 7 界面原型（Salt）
│   ├── 08_data_model_core.puml      # 图 8 核心数据模型
│   ├── 09_data_model_renewal.puml   # 图 9 续借事务数据模型
│   ├── 10_component_development.puml# 图 10 组件图（开发视图）
│   ├── 11_deployment.puml           # 图 11 部署视图
│   ├── 12_AI_arch_A_raw.puml        # 图 12 AI 方案 A（原始输出）
│   ├── 13_AI_arch_B_raw.puml        # 图 13 AI 方案 B（原始输出）
│   └── 14_AI_class_raw.puml         # 图 14 AI 关键类图（原始输出）
└── img/                             # 渲染图（由 puml 生成，与源文件同源）
```

### 渲染方式

```
java -jar plantuml.jar -tpng -charset UTF-8 -o img puml/*.puml
```

### 提交时序

手工设计（commit 1）先于 AI 对照件（commit 2）入库，报告与 PDF 最后提交（commit 3）。
实验二的 AI 对照件保留原始输出，未做语法或语义修改；评估与最终选择见报告第十四章。

