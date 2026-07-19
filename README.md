# ppt-conform — PPT 标准模板转换智能体

将源 PPT 按照企业标准模板进行自动转换，确保品牌视觉一致性。支持多种模板风格，保留内容完整性，自动去除水印。

---

## ✨ 特性

- **🎨 多风格模板支持** — 内置 4 种施耐德电气风格（白色简约 / 浅绿色清新 / 深绿色商务 / 渐变科技）
- **📝 内容完整保留** — 文本、表格、图表、图片、形状全部保留，不丢失信息
- **🧠 智能格式映射** — 模板格式优先 + 原格式兜底，内容与结构不割裂
- **🆔 精确去重机制** — 基于 shape_id 的内容映射，避免重复/丢失
- **🚫 水印自动去除** — 文本关键词 + 图片水印双重检测
- **✅ 质量校验** — 结构、样式、占位符、溢出检测多重校验
- **📐 强制 16:9** — 输出统一为宽屏格式
- **🌐 Web 界面** — Streamlit 可视化操作，上传即转换
- **💻 命令行接口** — 支持脚本调用和批量处理

---

## 🚀 快速开始

### 环境要求

- Python 3.10+
- pip

### 安装

```bash
# 克隆仓库
git clone https://github.com/ccjie168/ppt-conform.git
cd ppt-conform

# 安装依赖
pip install -r requirements.txt
```

### 启动 Web 界面

```bash
streamlit run app.py
```

浏览器自动打开，上传源 PPT → 选择风格 → 一键转换 → 下载结果。

### 命令行使用

```bash
python -m cli.main \
  --input source.pptx \
  --style F1 \
  --output output.pptx
```

**参数说明**:

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `--input` | 源 PPT 文件路径 | 必填 |
| `--output` | 输出 PPT 文件路径 | 必填 |
| `--style` | 模板风格 (F1/F2/F3/F4) | `F1` |
| `--template` | 自定义模板文件路径 | 内置模板 |

---

## 🎨 模板风格

| 风格 ID | 名称 | 背景色 | 适用场景 |
|---------|------|--------|----------|
| **F1** | 白色简约 | `#FFFFFF` | 通用文档、汇报材料 |
| **F2** | 浅绿色清新 | `#E7FFD9` | 环保、新能源主题 |
| **F3** | 深绿色商务 | `#0A2F24` | 正式商务、客户提案 |
| **F4** | 渐变科技 | `#0A2F24 → #3DCD58` | 技术分享、产品发布 |

---

## 🏗️ 架构设计

### 核心原理：内容与格式分离

```
┌─────────────┐     ┌─────────────┐     ┌─────────────┐
│  源 PPT     │────▶│  提取器      │────▶│ 内容模型    │
│  (任意格式)  │     │  Extractor  │     │ ContentBlock│
└─────────────┘     └─────────────┘     └──────┬──────┘
                                                │
                                                ▼
┌─────────────┐     ┌─────────────┐     ┌─────────────┐
│  模板 PPT   │────▶│  分析器      │     │  回填器      │
│  (标准模板)  │     │  Analyzer   │────▶│  Replayer   │
└─────────────┘     └─────────────┘     └──────┬──────┘
                                                │
                                                ▼
                                         ┌─────────────┐
                                         │  输出 PPT   │
                                         │  (标准格式)  │
                                         └─────────────┘
```

### 关键设计

1. **内容与格式分离提取** — 提取时同时记录内容和格式，建立 shape_id 映射
2. **模板格式优先，原格式兜底** — 回填时优先应用模板定义，缺失时用原格式
3. **精确去重** — shape_id 精确匹配，替代脆弱的文本匹配
4. **页眉页脚独立处理** — 不提取原PPT页眉页脚，直接使用模板标准格式

### 数据模型

```python
ContentBlock:
  ├── type: paragraph / list / image / table / chart
  ├── text: 文本内容
  ├── semantic_role: title / subtitle / body_main / body_sidebar
  ├── raw_shape_id: 与 raw_shape 对应的唯一标识
  ├── text_format: TextFormat（字体、字号、颜色、对齐...）
  └── shape_format: ShapeFormat（填充、边框、位置、大小...）
```

---

## 📁 项目结构

```
ppt-conform/
├── app.py                          # Streamlit Web 入口
├── cli/main.py                     # 命令行入口
├── config/                         # 配置文件
│   ├── layout_mappings.yaml        # 版式映射
│   ├── master_styles.yaml          # Master 风格配置
│   ├── validation_rules.yaml       # 校验规则
│   └── watermark_blacklist.yaml    # 水印黑名单
├── core/                           # 核心逻辑
│   ├── models.py                   # 数据模型
│   ├── extractor/pptx_extractor.py # PPT 内容提取器
│   ├── analyzer/template_analyzer.py # 模板分析器
│   ├── replayer/content_replayer.py # 内容回填器
│   ├── validator/validator.py      # 质量校验器
│   ├── registry/template_registry.py # 模板注册表
│   └── watermark/detector.py       # 水印检测器
├── templates/                      # 模板文件
├── tests/                          # 测试用例
├── docs/                           # 文档
├── requirements.txt
└── pyproject.toml
```

---

## 🧪 测试

```bash
# 运行全部测试
pytest tests/ -v

# 运行对比测试（验证内容完整性）
python tests/test_conversion_compare.py

# 运行单个测试文件
pytest tests/test_extractor.py -v
```

**当前测试覆盖率**: 21 个测试用例，全部通过
**转换成功率**: 100%（标准对比测试）

---

## ⚙️ 配置说明

### 16:9 宽屏尺寸

| 属性 | 值 |
|------|-----|
| 宽度 | 12192000 EMU (13.33 英寸) |
| 高度 | 6858000 EMU (7.5 英寸) |
| 比例 | 16:9 (约 1.778) |

### 施耐德品牌色

| 颜色 | 色值 | 用途 |
|------|------|------|
| 深绿色 | `#0A2F24` | 深色背景、标题文字 |
| 浅绿色 | `#E7FFD9` | 浅色背景 |
| 亮绿色 | `#3DCD58` | 强调色、Logo |
| 白色 | `#FFFFFF` | 白色背景、深色页文字 |

### 页眉页脚阈值

- **页眉区域**: 幻灯片顶部 8%
- **页脚区域**: 幻灯片底部 12%

---

## 🔧 API 使用

### Python API

```python
from core.extractor.pptx_extractor import PptxExtractor
from core.registry.template_registry import TemplateRegistry
from core.replayer.content_replayer import ContentReplayer
from core.models import UserConfig

# 1. 提取源 PPT 内容
extractor = PptxExtractor()
content_models = extractor.extract("source.pptx")

# 2. 初始化模板注册表和回填器
registry = TemplateRegistry()
replayer = ContentReplayer(
    registry,
    template_path="templates/se_energy_tech_ppt_20260421.pptx"
)

# 3. 配置并执行转换
config = UserConfig(
    input_path="source.pptx",
    output_path="output.pptx",
    master_style="F1",  # F1-F4
)
replayer.replay(content_models, config)

print(f"转换完成: output.pptx")
```

---

## 📋 支持的内容类型

| 类型 | 支持程度 | 说明 |
|------|----------|------|
| 文本（标题/正文） | ✅ 完整支持 | 模板格式优先，原格式兜底 |
| 列表（多级） | ✅ 完整支持 | 保留层级结构 |
| 表格 | ✅ 完整支持 | 保留内容和基础样式 |
| 图片 | ✅ 完整支持 | 按原位置和比例放置 |
| 图表 | ⚠️ 部分支持 | 保留图表对象，样式可能有差异 |
| 自选图形 | ✅ 完整支持 | 保留填充色和边框色 |
| SmartArt | ❌ 暂不支持 | 未来版本计划 |
| 动画/切换 | ❌ 不支持 | 模板转换一般不需要 |
| 音视频 | ❌ 不支持 | |

---

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！

### 开发环境

```bash
git clone https://github.com/ccjie168/ppt-conform.git
cd ppt-conform
pip install -r requirements.txt
pip install pytest
pytest tests/ -v
```

### 提交规范

- feat: 新功能
- fix: 修复
- docs: 文档
- refactor: 重构
- test: 测试
- perf: 性能优化

---

## 📄 许可证

MIT License

---

## 📮 联系方式

- GitHub: [ccjie168/ppt-conform](https://github.com/ccjie168/ppt-conform)
- Issues: 欢迎提交问题和建议

---

*本项目用于企业 PPT 标准化转换，提升品牌视觉一致性*
