# PPT 生成 MCP

把一个结构化的 `deck` JSON 渲染成可编辑的 `.pptx`。它是独立的本地 MCP，不调用在线模型、图片搜索或任何商业服务；主模型负责写内容，这个技能只负责排版。

## 功能

- 自动生成封面和页脚，可选自动结尾页。
- 9 种版式：`section`、`bullets`、`two_column`、`comparison`、`image`、`table`、`quote`、`stats`、`closing`。
- 文本支持 `**粗体**`、`*斜体*` 和 `` `等宽` ``。
- 本地图片按中心裁切填满相框，保持比例，不拉伸。
- 表格、数据卡、引用页和演讲者备注。
- 四种配色：`urban` 暖灰（默认）、`ion` 紫罗兰、`minimal` 极简、`blank` 经典蓝。
- 完全离线，依赖已放在 `vendor/`。

## 工具

### `create_presentation`

| 参数 | 必填 | 说明 |
| --- | --- | --- |
| `deck` | 是 | 演示内容对象，见下方 schema。 |
| `output_path` | 是 | 输出 `.pptx` 的绝对路径，父目录会自动创建。 |
| `template` | 否 | `urban` / `ion` / `minimal` / `blank`，默认 `urban`。 |
| `overwrite` | 否 | 输出已存在时是否覆盖，默认 `false`。 |

返回 JSON：

```json
{
  "ok": true,
  "output_path": "D:\\deck\\季度汇报.pptx",
  "template": "urban",
  "template_label": "暖灰",
  "slide_count": 8,
  "file_size": 123456,
  "warnings": []
}
```

## deck schema

```json
{
  "title": "从数据到决策",
  "subtitle": "2026 年产品策略复盘",
  "author": "战略与产品团队",
  "date": "2026.10",
  "eyebrow": "QUARTERLY REVIEW",
  "footer": "从数据到决策",
  "cover_image": "D:\\assets\\cover.jpg",
  "closing": {
    "title": "谢谢",
    "subtitle": "欢迎交流，也欢迎把下一步交给我。"
  },
  "slides": [
    {
      "layout": "section",
      "eyebrow": "01 / CONTEXT",
      "title": "我们面对的三个变化",
      "subtitle": "增长放缓、渠道碎片化、决策周期变长"
    },
    {
      "layout": "bullets",
      "eyebrow": "02 / FINDINGS",
      "title": "三个关键发现",
      "bullets": [
        "**高价值用户**贡献了 62% 的复购，但只占活跃用户的 18%",
        "新渠道获客成本在 6 个月内上升了 *34%*",
        "决策链从 2 人扩展到 4.2 人，销售周期变长"
      ],
      "key_message": "增长问题不是流量不够，而是高价值人群没有被识别和经营。"
    },
    {
      "layout": "two_column",
      "title": "现状与目标",
      "columns": [
        {
          "heading": "现状",
          "bullets": ["指标看板分散", "复盘依赖人工", "策略落地慢"]
        },
        {
          "heading": "目标",
          "bullets": ["一个统一指标口径", "每周自动复盘", "策略 48 小时内上线"]
        }
      ],
      "key_message": "先把口径统一，再谈自动化。"
    },
    {
      "layout": "comparison",
      "title": "两种路径的取舍",
      "columns": [
        {"heading": "继续堆功能", "bullets": ["短期收入可见", "维护成本持续上升", "用户价值不清晰"]},
        {"heading": "聚焦高价值人群", "bullets": ["短期增长较慢", "复购和客单更高", "组织协同更简单"]}
      ]
    },
    {
      "layout": "image",
      "title": "用户旅程中的关键时刻",
      "bullets": ["首次激活", "第 7 天留存", "第 30 天复购"],
      "image": "D:\\assets\\journey.png",
      "image_caption": "图：用户旅程中的三个关键节点",
      "image_position": "right"
    },
    {
      "layout": "table",
      "title": "三条策略的投入产出",
      "table": {
        "headers": ["策略", "投入", "预计影响", "优先级"],
        "rows": [
          ["高价值人群识别", "中", "+8% 复购", "P0"],
          ["渠道归因统一", "低", "节省 12% 预算", "P1"],
          ["自动复盘", "中", "每周节省 6 小时", "P1"]
        ]
      }
    },
    {
      "layout": "quote",
      "quote": "真正的增长，来自对少数关键行为的持续投入。",
      "attribution": "产品复盘会议，2026.10"
    },
    {
      "layout": "stats",
      "title": "一页看清结果",
      "metrics": [
        {"value": "62%", "label": "高价值用户复购占比", "note": "较上季度 +7pt"},
        {"value": "34%", "label": "新渠道获客成本涨幅", "note": "需要重新评估渠道组合"},
        {"value": "48h", "label": "策略上线时间", "note": "目标从 2 周缩短到 2 天"}
      ],
      "key_message": "把资源从“更多流量”转向“更高价值”。"
    }
  ]
}
```

### 版式字段

| layout | 必填字段 | 可选字段 |
| --- | --- | --- |
| `section` | `title` | `eyebrow`、`subtitle` |
| `bullets` | `title`、`bullets` | `key_message`、`numbered`、`notes` |
| `two_column` | `title`、`columns` | `key_message` |
| `comparison` | `title`、`columns`（2 栏） | `key_message` |
| `image` | `title`、`image` | `bullets`、`image_caption`、`image_position`、`key_message` |
| `table` | `title`、`table.headers`、`table.rows` | `key_message` |
| `quote` | `quote` | `attribution`、`notes` |
| `stats` | `title`、`metrics` | `key_message` |
| `closing` | 无 | `title`、`subtitle` |

`bullets` 的元素可以是字符串、`{"text": "...", "level": 1}`，或嵌套数组。单页最多 12 条；超过 6 条会自动分两栏。

## 安装

### 从技能市场安装

把整个目录打成 zip，`mcp.json` 必须位于压缩包根目录：

```powershell
Compress-Archive -Path .\* -DestinationPath ..\ppt-generator-1.0.0.zip
```

市场安装后会解压到 `data/mcps/ppt-generator`，默认不启用；在技能市场里打开开关即可。

### 本地直接启用

把目录放到 `mcps/ppt-generator`，在 `data/mcp.json`（或应用的 MCP 设置）里启用；也可以用 MCP 客户端直接运行：

```powershell
python C:\path\to\ppt-generator\server.py
```

协议是 stdio JSON-RPC，支持 `initialize`、`tools/list`、`tools/call`、`ping`。

## 设计说明

- 只使用 python-pptx 的原生形状、文本框、表格和图片，生成的文件可以在 PowerPoint / WPS / LibreOffice 里继续编辑。
- 版式参考 [slide-deck-ai](https://github.com/barun-saha/slide-deck-ai)（MIT）的模板驱动思路；渲染层为独立实现，没有引入它的 LiteLLM、Streamlit、图片搜索或 API Key 逻辑。
- 模板文件来自 slide-deck-ai 的 `pptx_templates/`，许可见 `TEMPLATES_LICENSE.md`。
- 第三方 Python 依赖的许可见 `THIRD_PARTY_NOTICES.md`。

## 限制

- 单次最多 40 页。
- 图片只接受本地绝对路径，不下载网络图片。
- 表格建议不超过 8 行；更多内容建议拆页。
- 16:9 模板；不支持自定义母版或任意版式脚本。
