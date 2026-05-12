# CLAUDE.md — 消费红利策略项目注意事项

## 项目概览

消费红利股票多因子选股策略课程作业。数据月频（2013-08 ~ 2022-11），5 列：code/close/divi/cap/date。详细大纲：`任务大纲.md`。

## 数据口径关键结论（经实证验证）

### divi：股息率快照，非分红事件
- `dps_proxy = close * divi / 100` 在相邻月之间几乎不变（98% 变化 < 0.01）
- `close` 变化与 `divi` 变化相关系数 ≈ -0.78，符合 `yield = 固定DPS / 变动close`
- divi 几乎每月有值 → 连续更新的月度快照，非事件标识

### DY 因子构建：必须走 dps_proxy → ffill → /close 路线
- 不能直接 ffill divi：会冻结旧股价下的股息率
- 实证：10.5% 的行两种方法差异 > 0.1%，最大差 7.55%
- 正确做法：`dps_proxy = close * divi / 100` → `groupby(code).ffill()` → `dy = dps_proxy_ffill / close * 100`
- 覆盖率：divi 原始 68.6%，dps ffill 后 84.9%

### 收益口径：用价格收益
- 无法从 divi 重建含息收益（无除息日信息）
- 主收益：`price_ret = close_{t+1} / close_t - 1`
- 报告中声明保守偏差

### 样本池动态变化
- 有效样本从 135 只（2013-08）逐步增到 259 只（2022-11），每月动态筛选 `close.notna() & (cap > 0)`

### 已排除的因子
EP、BM、日度低波、换手率/流动性 — 均无对应数据

## 环境障碍

- conda run 不支持多行 `-c` 内联脚本 → 写入临时 `.py` 文件执行
- 环境：`D:\Anaconda\envs\QuantEnv\python.exe`

## 已决策参数
- 费率 `0.001`（单边千分之一），敏感性 `0` / `0.002` / `0.003`
- 基准：可交易样本等权，价格收益口径，不计手续费
- 月度调仓主策略，季度对照
- 收益不含分红 → 红利税不纳入回测

## 增强实验结论（2026-05-11）

### 温和混合为最优策略
- `composite = 0.7 × dy_neutral_z − 0.3 × rev_1m_z`：年化 19.45%，夏普 0.64，换手 27.4%
- 全面优于 DY(规模中性) 的 17.68%/0.57 和等权合成 17.20%/0.57/40.8%
- 反转因子应作为辅助信号（30% 权重）而非等权因子

### 已弃用方案
- 反转重筛选（DY Top30→反转筛至 Top15）：仅 306.6%，低于纯 DY_neut 351.0%
  原因：在已筛选高质量池内二次硬阈值筛选反而剔除了有效信号
- 趋势过滤（修正未来函数后）：降至 13.39%/0.39，低于无过滤
  原因：简单价格均线择时无法区分策略"有效"和"无效"的市场区间

### 随机基准
- DY(规模中性)：P87.0, Z=1.09（中等偏上但未达极端分位）
- 仅股息率：P62.2, Z=0.16（接近随机分布中心）

### Fama-MacBeth 规模控制
- `ret ~ dy_a + ln_cap + rev_a` 多变量 FM 中 DY t=1.60（边际支持，非显著），应写"控制规模和反转后 DY 仍为正、具备边际支持"
- 规模中性化后的 DY 并非规模效应的衍生品

### 2026-05-11 笔记本重排
- Cell 35 (5.2 绩效分析续)、36-37 (超额/回撤) 曾错放在 Ch6 区域 → 已移回 Ch5
- Cell 39 (调频/TopN/费率) 曾错放在 Ch6 区域 → 已移入 Ch7
- Cell 61 (7.6 多空验证) 曾错放在 Ch8 区域 → 已移入 Ch7
- Cell 49 的 7.1+7.2+7.3 已在同一 markdown cell 中（无需拆分）
- 温和混合已导出到 output/summary_enhanced.json 和 output/table_enhanced_core.csv

## 2026-05-11 报告去 AI 化 + 风格优化

### report.md 修改
- 应用 humanizer-zh skill 规则去 AI 化：删填充短语、打破三段式结构、变化句子节奏、减少破折号
- 每小节开头加入核心判断句（券商金工研报惯例）
- 摘要改为两段直接陈述，去掉"可以概括为三点"等模板句
- 经济逻辑（原第 5 节）收紧为 1-2 段
- 风险提示压缩重复表述
- 数值、表格、图表路径均未改动

### notebook markdown 优化
- 33 个 markdown cell 重写：保持简短，每节说明目的 + 关键发现
- 标题编号连续（1.1→1.3, 2.1→2.3, 3.1, 4.1, 5.1→5.2, 6.1→6.6, 7.1→7.6, 8.1→8.3, 9）
- FM 代码 cell 使用 display() 输出表格（无长 print）
- 验证：0 errors, 0 unexecuted cells, 无 bfill, 无绝对路径, 无含息总收益

### Word 生成
- build_report_docx.py 正常运行，输出 `消费红利策略研究报告.docx`（1.9MB）

### 量化 skill 审查
- quant-analyst skill 实际为加密货币回测框架（openclaw-quant），不适用 A 股因子审查
- humanizer-zh skill 成功加载，对 report.md 实施 24 条规则去 AI 化
- 代理恢复后成功安装 quant-factor-screener 和 grad-fama-french 两个量化 skill
- quant-factor-screener 审查发现：低波因子失效（t=0.09）与A股低波异象普遍实证不符，已在report.md补充讨论；行业集中风险需在风险提示中明确
- grad-fama-french 审查发现：FM多因子中Rev（t=2.40）强于DY（t=1.60），增强策略收益部分来自反转成分，report.md已修正表述——明确0.7/0.3权重的换手+经济逻辑双重考量，不纯以统计预测力为标准

## 笔记本执行注意事项
- `jupyter nbconvert --to notebook --execute --allow-errors --inplace` 是可靠的全量执行方式
- 新增 cell 必须确保 `execution_count` 和 `outputs` 字段存在（nbformat 要求）
- 修改 notebook JSON 后应通过 nbconvert 重新执行确保一致性