# gaozhong-yingyu-handout-to-courseware

高中英语「讲义 docx + PPT模版 → 课件 pptx」生成 skill。

- 自动识别讲次类型：主观题·读后续写（第9讲型）/ 客观题·阅读理解（第10讲型）
- 内容来源配方：A 讲义直出 / B 模版直出 / C AI生成（标"待老师审核"）
- **模版拆解 Gate**：`dissect_template.py` 把模版拆成 `template_spec.md/json`，`build_pptx.py` 生成前强制核对字体/字号/版式
- 题型分页：阅读·七选五·语法填空·翻译每页1题，完形每页3-4题；讲解一律"题页→答页"两页揭示
- 标杆PPT 仅用于一次性提炼结构与设计 token，**不克隆**

## 快速开始
```bash
# 0.（新模版必跑）拆解模版 → 生成前 Gate
python3 scripts/dissect_template.py "原始课件模板.pptx" -o references/
# 1. 解析讲义（自动识别类型 + 抽 Knowledge Map 图）
python3 scripts/extract_handout.py "讲义.docx" -o output/
# 2.（可选）编写 output/ai_scaffold.json —— C类AI教学脚手架，schema 见 references/content_schema.md
# 3. 渲染课件
python3 scripts/build_pptx.py --content output/content.json --template "原始课件模板.pptx" -o output/课件.pptx
```

详见 [SKILL.md](SKILL.md)。已验证：第9讲（读后续写，83页）、第10讲（阅读理解，52页·盲跑）。
