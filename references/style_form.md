# 全局文字样式表单（逐run拆解自参考PPT；渲染唯一来源）

| 样式key | 字号 | 加粗 | 颜色 | 西文(latin) | 中文(ea) | 备注 |
|---|---|---|---|---|---|---|
| `title_slot` | 40 | ✗ | BA7AC2 | 阿里巴巴普惠体 B | 阿里巴巴普惠体 B |  |
| `loc_tag` | 24 | 继承 | A076CE | 微软雅黑 | 微软雅黑 |  |
| `section_big_reading` | 199 | ✓ | 默认 | 可口可乐在乎体 楷体 | 可口可乐在乎体 楷体 |  |
| `section_big_continuation` | 160 | ✗ | 默认 | 可口可乐在乎体 楷体 | 可口可乐在乎体 楷体 |  |
| `source_big_reading` | 72 | ✓ | 默认 | 可口可乐在乎体 楷体 | 可口可乐在乎体 楷体 |  |
| `source_big_continuation` | 88 | ✗ | 默认 | 可口可乐在乎体 楷体 | 可口可乐在乎体 楷体 |  |
| `passage` | 20 | ✗ | 默认 | Arial | 等线 |  |
| `passage_mark` | 20 | ✓ | 默认 | Arial | 等线 |  |
| `q_stem` | 24 | ✓ | 默认 | Arial | 继承 |  |
| `q_opt` | 24 | ✗ | 默认 | Arial | 继承 |  |
| `q_opt_zh` | 24 | ✓ | 默认 | 微软雅黑 | 微软雅黑 |  |
| `q_practice` | 24 | ✗ | 默认 | Arial | 继承 |  |
| `ans_head` | 28 | ✗ | FF0000 | Arial | 继承 |  |
| `ans_body` | 20 | ✗ | FF0000 | Arial | 继承 |  |
| `test_stem` | 24 | ✓ | 默认 | Arial | 继承 |  |
| `test_opt` | 24 | ✗ | 默认 | Arial | 继承 |  |
| `test_ans_head` | 68 | ✗ | FF0000 | Arial | 继承 |  |
| `test_ans_body` | 24 | ✗ | FF0000 | Arial | 继承 |  |
| `method_head` | 44 | ✓ | 默认 | 可口可乐在乎体 楷体 | 可口可乐在乎体 楷体 | 标杆方法标题为sz118独立大字页，此处适配为页内标题 |
| `method_line` | 24 | ✗ | 默认 | 微软雅黑 | 微软雅黑 |  |
| `warn_line` | 30 | ✓ | 默认 | 微软雅黑 | 微软雅黑 |  |
| `intro_big_continuation` | 90 | 继承 | 默认 | 可口可乐在乎体 楷体 | 可口可乐在乎体 楷体 |  |
| `guide_prompt` | 96 | 继承 | 默认 | 可口可乐在乎体 楷体 | 可口可乐在乎体 楷体 |  |
| `plot_box` | 53 | 继承 | 默认 | 可口可乐在乎体 楷体 | 可口可乐在乎体 楷体 |  |
| `label_badge` | 46 | 继承 | 默认 | 阿里巴巴普惠体 B | 阿里巴巴普惠体 B |  |
| `plot_name_box` | 60 | 继承 | 默认 | 可口可乐在乎体 楷体 | 可口可乐在乎体 楷体 |  |
| `trans_basic_zh` | 66 | 继承 | 默认 | 可口可乐在乎体 楷体 | 可口可乐在乎体 楷体 |  |
| `trans_basic_en` | 46 | 继承 | 默认 | Arial | 宋体 |  |
| `trans_upgrade_zh` | 32 | 继承 | 默认 | 可口可乐在乎体 楷体 | 可口可乐在乎体 楷体 |  |
| `trans_upgrade_en` | 26 | 继承 | 默认 | Arial | 宋体 |  |
| `detail_label` | 36 | 继承 | 默认 | 可口可乐在乎体 楷体 | 可口可乐在乎体 楷体 |  |
| `detail_group` | 38 | 继承 | 默认 | 可口可乐在乎体 楷体 | 可口可乐在乎体 楷体 |  |
| `detail_zh` | 32 | 继承 | 默认 | 可口可乐在乎体 楷体 | 可口可乐在乎体 楷体 |  |
| `detail_en` | 28 | 继承 | 默认 | Arial | 等线 |  |
| `seg_header` | 26 | ✓ | 默认 | Arial | 阿里巴巴普惠体 B |  |
| `essay_body` | 24 | 继承 | 默认 | Arial | 等线 |  |
| `essay_badge` | 33 | ✓ | 默认 | 阿里巴巴普惠体 B | 阿里巴巴普惠体 B |  |

> 改样式只改本表单(json)，不改代码；缺省项=继承模版。混排规则：题干粗/选项不粗/选项中文粗微软雅黑；【答案】大字+【解析】小字；语篇仅段号①加粗。