---
name: wjx-score-analysis
description: Use when user mentions 成绩单, 问卷星, wjx.cn, or wants to extract exam scores / generate score reports for 展鹏公考. Also triggers on "今日成绩", "昨日成绩" etc.
---

# 展鹏问卷星成绩单分析

从问卷星测试活动页面提取所有学员的逐题作答数据，用 `template.html` 模板生成 HTML 成绩分析报告。

## 前提条件

1. Chrome 已安装（如果远程调试未开启，脚本会自动拉起 `google-chrome --remote-debugging-port=9222`）
2. 至少有 1 个标签页（脚本会自动检查登录状态，未登录时自动尝试登录，失败则等待人工协助）

## 使用流程

**一条命令完成全部流程**（登录检查 → 数据提取 → AI 分析 → HTML 生成）：

```bash
# 今日成绩
python3 /home/lujunjie/Project/wjx-score-analysis/.claude/skills/wjx-score-analysis/extract_and_generate.py \
  --day today \
  --output-dir /home/lujunjie/Project/wjx-score-analysis/成绩单 \
  --exam-name "学前测"

# 昨日成绩（把 today 换成 yesterday）
# 前天（把 today 换成 前天）
# 也可用偏移量：-1=今日, -2=昨日, -3=前天
```

**用户说"查询今日/昨日成绩单"时，直接运行上面的命令，不需要额外步骤。**

也可传入完整 URL（带自定义筛选条件）：
```bash
python3 /home/lujunjie/Project/wjx-score-analysis/.claude/skills/wjx-score-analysis/extract_and_generate.py \
  "用户提供的完整URL" \
  --output-dir /home/lujunjie/Project/wjx-score-analysis/成绩单 \
  --exam-name "学前测"
```

参数:
- 第一个参数（可选）: 完整 URL 或 activity_id，使用 `--day` 时可省略
- `--day`: 日期快捷方式，支持 `today`/`今日`、`yesterday`/`昨日`、`前天`，或偏移量 `-1`/`-2`/`-3`...
- `--activity`: activity_id，默认 `331434168`
- `--output-dir`: 输出目录，默认 `./成绩单`
- `--exam-name`: 考试名称
- `--limit N`: 只提取前 N 人（调试用）
- `--template`: 自定义模板路径（默认用同目录 template.html）
- `--username`/`--password`: 问卷星登录凭据（已内置默认值）

脚本自动完成：自动登录 → 提取汇总 → 提取详情 → 生成 学情分析 → 输出 HTML 成绩单。

## 输出

文件名自动带上目标日期前缀：

```
成绩单/
  2026-04-15-学前测成绩单-{姓名}.html   # 基于 template.html 的报告
  2026-04-15-全部成绩详情.json           # 完整数据
  2026-04-15-成绩汇总.csv                # 汇总表
```

## 模板说明

`template.html` 使用 `{{placeholder}}` 占位符，脚本通过字符串替换填充：

| 占位符 | 内容 |
|--------|------|
| `{{name}}` | 学员姓名 |
| `{{exam_name}}` | 考试名称 |
| `{{submit_time}}` | 提交时间 |
| `{{total_score}}` | 总分 |
| `{{total_questions}}` | 总题数 |
| `{{total_correct}}` | 答对数 |
| `{{overall_rate}}` | 总正确率 |
| `{{subject_cards}}` | 科目汇总横表 HTML |
| `{{radar_svg}}` | SVG 雷达图 |
| `{{bar_chart}}` | 水平柱状图 HTML |
| `{{detail_tables}}` | 细分横表 HTML |
| `{{ai_section}}` | AI 分析 HTML |

修改 `template.html` 的 CSS 即可统一调整所有成绩单的风格。
