
# 加密货币网站收集器 - 使用说明

## 📋 功能介绍

这个程序可以自动：
1. 从CoinDesk等加密货币网站获取最新文章
2. 筛选出阅读量最高的3篇文章
3. 将文章翻译成中文
4. 保存为格式化的Word文档

## 🛠️ 安装依赖

在运行程序之前，需要安装以下Python库：

```bash
pip install requests beautifulsoup4 python-docx
```

## 📝 使用方法

### 方法一：直接运行

```bash
python coindesk_collector.py
```

程序会自动生成一个Word文档，文件名格式为：`CoinDesk热门文章_YYYY-MM-DD.docx`

### 方法二：作为模块导入

```python
from coindesk_collector import CoinDeskCollector, WordDocumentCreator

# 获取文章
collector = CoinDeskCollector()
articles = collector.get_articles()

# 创建文档
doc = WordDocumentCreator()
# ... 自定义处理
```

## ⚙️ 配置说明

### 修改文章数量

在`main()`函数中修改：
```python
top_articles = articles_sorted[:3]  # 改为你想要的数量
```

### 修改翻译API

当前使用MyMemory免费翻译API，如需使用其他翻译服务：
1. 百度翻译API
2. 腾讯翻译API
3. Google Cloud Translation API

修改`Translator`类中的`translate_to_chinese()`方法即可。

### 自定义文档格式

在`WordDocumentCreator`类中可以修改：
- 标题字体大小和颜色
- 段落间距
- 文档样式等

## 📁 输出文件

生成的Word文档包含：
- 文档标题（包含日期）
- 生成时间和来源信息
- 每篇文章包含：
  - 中文翻译标题
  - 原文标题
  - 原文链接
  - 中文翻译摘要
  - 中文翻译正文

## ⚠️ 注意事项

1. **网络连接**：需要稳定的网络连接访问CoinDesk网站
2. **请求频率**：程序已添加延迟，避免请求过快被封禁
3. **翻译限制**：免费翻译API有调用次数限制
4. **网站更新**：如果CoinDesk网站结构改变，需要更新HTML选择器

## 🔧 故障排除

### 问题1：无法获取文章

**可能原因**：
- 网络连接问题
- CoinDesk网站结构改变
- 被网站反爬虫机制拦截

**解决方法**：
- 检查网络连接
- 更新HTML选择器
- 添加代理或使用更长的延迟

### 问题2：翻译失败

**可能原因**：
- 翻译API调用限制
- 网络问题

**解决方法**：
- 等待一段时间后重试
- 更换翻译API
- 使用付费翻译服务

### 问题3：Word文档无法打开

**可能原因**：
- 文档生成过程中出错
- python-docx库版本问题

**解决方法**：
- 更新python-docx库
- 检查错误日志

## 📜 许可证

本程序仅供学习和个人使用。

## 🤝 贡献

欢迎提交Issue和Pull Request来改进这个程序！
