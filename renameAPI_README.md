# renameAPI 使用说明

## 功能概述

提供一个API接口，接收本地文件路径或file:// URL列表，按规则匹配文件名并直接在原路径覆写重命名。

## 启动方式

```bash
python renameAPI.py
```

默认监听 `0.0.0.0:8000`。

## 接口说明

### POST /jr/rename

请求体：

```json
{
  "url_list": [
    "C:\\data\\承租人\\审计报告_2024.pdf",
    "file:///C:/data/担保人/年度审计报告.docx"
  ]
}
```

响应体：

```json
{
  "results": [
    {
      "url": "C:\\data\\承租人\\审计报告_2024.pdf",
      "status": "renamed",
      "old_name": "审计报告_2024.pdf",
      "new_name": "审计报告_2024#承租人审计报告#.pdf",
      "rule": "审计报告1"
    }
  ]
}
```

结果状态：

- renamed：命中规则并完成重命名
- skipped：未命中规则或文件名无需变更
- error：路径不合法或文件不存在

### GET /health

返回服务状态：

```json
{
  "status": "ok"
}
```

## 规则说明

匹配逻辑基于以下规则：

- 文件名关键词匹配
- 路径包含指定的文件夹名称

规则定义在 `renameAPI.py` 的 `dict_t["rules"]` 中。

## 注意事项

- 仅支持本地路径或 file:// URL
- http/https URL不处理
- 直接在原路径覆写重命名，请确保有写入权限
