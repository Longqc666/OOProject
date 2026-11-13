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
