---
name: wjx-score-analysis
description: Use when user mentions 成绩单, 问卷星, wjx.cn, or wants to extract exam scores / generate score reports for 展鹏教育. Also triggers on "今日成绩", "昨日成绩" etc.
---

# 展鹏问卷星成绩单分析

从问卷星测试活动页面提取所有学员的逐题作答数据，生成 HTML 成绩分析报告。

## 前提条件

已安装: `cd /home/lujunjie/Project/wjx-score-analysis && pip install -e .`

## 使用

```bash
# 启动 Web UI（自动打开 Chrome + 浏览器）
wjx-score

# 已有 Chrome 运行时
wjx-score --no-chrome

# 自定义端口
wjx-score --port 3000 --cdp-port 9222

# 自定义输出目录和考试名称
wjx-score --output-dir ./成绩单 --exam-name "学前测"
```

启动后在浏览器中操作：设置筛选条件 → 查询 → 勾选学员 → 生成成绩单。

## 输出

文件名自动带上目标日期前缀：

```
成绩单/
  学前测成绩单-{姓名}.html   # HTML 分析报告
  全部成绩详情.json           # 完整数据
  成绩汇总.csv                # 汇总表
```
