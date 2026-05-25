# EDC智能稽查与RBQM风险识别系统（MVP）

这是一个可本地运行的 Streamlit MVP，用于：
- 读取 Excel **全部 sheet**
- 展示 sheet 摘要与预览
- 自动推荐字段映射
- 解析方案 PDF / DOCX 文本
- 生成规则草稿与可编辑规则表
- 执行基础核查，输出 Findings / Query / RBQM
- 生成 Word 稽查报告

## 运行方式

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 主要说明

1. 这是一版 **可运行的 MVP**，重点先解决：
   - Excel 全 sheet 完整读取
   - 字段映射底座
   - 基础规则引擎
   - Findings / Query / RBQM / Word 报告导出

2. 方案规则抽取为 **启发式抽取 + 人工可编辑** 模式：
   - 会先从方案文本中抽取关键章节/阈值线索
   - 再自动生成规则草稿
   - 用户可在页面中继续修订

3. 核查引擎当前实现了 MVP 级规则：
   - 关键字段缺失
   - 访视/时间逻辑冲突
   - 实验室阈值规则（基于可编辑规则表）
   - AE 严重程度 vs SAE 一致性检查
   - 已退出仍给药
   - 筛选/随机/给药时间顺序检查

4. 受限分析逻辑已实现：
   - 若仅读取到少量 sheet 或关键模块缺失，会标记“受限分析”
   - 不会把“未读取到”误写为“不存在”

## 建议目录结构

- `app.py`：主应用
- `requirements.txt`
- `README.md`

## 后续增强建议

- 接入数据库（PostgreSQL）
- 用户权限与项目历史
- 更强的协议规则抽取
- 更完整的给药 / 访视窗口计算
- CAPA 建议引擎
- 报告模板美化
