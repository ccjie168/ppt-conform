# PPT 标准模板转换智能体 - 项目状态总结

> **项目路径**: `/workspace/ppt-conform/`
> **GitHub 仓库**: https://github.com/ccjie168/ppt-conform
> **最后更新**: 2026-07-19
> **最新 Commit**: `d66b462` (feat: 解决内容与格式割裂问题，转换成功率达100%)

---

## 一、项目概述

将源 PPT（Trae/豆包生成或其他来源）按照公司（施耐德电气）标准模板进行转换，确保品牌一致性。

### 核心功能
- ✅ 模板应用（4种风格：白色、浅绿色、深绿色、渐变）
- ✅ 内容完整保留（文本、表格、图表、形状、图片）
- ✅ 内容与格式智能映射（模板格式优先，原格式兜底）
- ✅ 质量校验（结构、样式、占位符、溢出检测）
- ✅ 去水印功能（文本关键词、图片水印检测）
- ✅ 动态 Master 识别（4种风格自动识别）
- ✅ 模板持久化（上次上传模板自动复用）
- ✅ 强制 16:9 输出
- ✅ Streamlit Web 界面
- ✅ 命令行接口

### 当前转换成功率
- **标准对比测试**: 100% (20/20 通过)
- **测试覆盖**: 封面页、单文本框、多文本框、图片页、表格页 × 4种Master风格 + 默认模板

---

## 二、项目结构

```
ppt-conform/
├── app.py                          # Streamlit 前端入口
├── cli/
│   └── main.py                     # 命令行入口
├── config/
│   ├── layout_mappings.yaml        # 版式映射配置
│   ├── master_styles.yaml          # Master 风格配置
│   ├── validation_rules.yaml       # 校验规则
│   └── watermark_blacklist.yaml    # 水印黑名单
├── core/
│   ├── analyzer/
│   │   ├── template_analyzer.py    # 模板分析器（风格识别核心）
│   │   └── template_format_extractor.py
│   ├── extractor/
│   │   └── pptx_extractor.py       # PPT 内容抽取器（内容+格式分离提取）
│   ├── registry/
│   │   └── template_registry.py    # 模板注册表
│   ├── replayer/
│   │   └── content_replayer.py     # 内容重放器（模板格式优先+原格式兜底）
│   ├── validator/
│   │   ├── validator.py            # 质量校验器
│   │   └── rules.py                # 校验规则
│   ├── watermark/
│   │   └── detector.py             # 水印检测器
│   └── models.py                   # 数据模型（TextFormat/ShapeFormat/ContentBlock）
├── templates/
│   ├── se_energy_tech_ppt_20260421.pptx  # 施耐德模板（4种风格）
│   └── icons/
├── tests/                          # 单元测试
│   ├── test_conversion_compare.py  # 对比测试（内容完整性验证）
│   ├── test_e2e.py
│   ├── test_extractor.py
│   ├── test_models.py
│   ├── test_registry.py
│   ├── test_replayer.py
│   ├── test_validator.py
│   └── test_watermark_detector.py
├── requirements.txt
├── pyproject.toml
└── deploy.sh
```

---

## 三、4种模板风格（Master）

| 风格ID | 风格名称 | 背景色 | 主题名称 | 状态 |
|--------|----------|--------|----------|------|
| F1 | 白色简约 | #FFFFFF | White content slides | ✅ 已识别 |
| F2 | 浅绿色清新 | #E7FFD9 | Light green content slides | ✅ 已识别 |
| F3 | 深绿色商务 | #0A2F24 | Dark green content slides | ✅ 已识别 |
| F4 | 渐变科技 | 渐变 (#0A2F24→#3DCD58) | Gradient content slides | ✅ 已识别 |

**识别策略**: 主题名称优先 + 背景色验证（在 `core/analyzer/template_analyzer.py` 中实现）

---

## 四、核心架构设计（内容与格式分离）

### 架构原则
1. **提取阶段**: 内容与格式分离，建立精确映射关系（shape_id）
2. **回填阶段**: 模板格式优先，原格式兜底
3. **去重策略**: shape_id 精确去重，替代文本匹配
4. **页眉页脚**: 不提取原PPT内容，只保留模板标准格式和施耐德图标

### 数据模型
- **TextFormat**: 字体名称、字号、颜色、粗体、斜体、下划线、对齐、行距
- **ShapeFormat**: 填充色、填充类型、线条色、线条宽度、位置、大小、旋转
- **ContentBlock**: content + text_format + shape_format + raw_shape_id + semantic_role

### 核心方法
- `PptxExtractor._extract_slide()` - 提取内容和格式，分配 shape_id，建立映射
- `PptxExtractor._extract_text_format()` - 递归解析格式继承（run→占位符→主题）
- `ContentReplayer._copy_slide_content()` - 内容回填，shape_id 精确去重
- `ContentReplayer._find_placeholder_by_role()` - 语义角色匹配占位符
- `ContentReplayer._apply_template_format_to_paragraph()` - 模板格式优先应用
- `ContentReplayer._apply_run_format()` - run级格式应用（模板→原格式→默认）

### 回填策略
- **标题**: 填入模板 title 占位符，模板样式优先，原格式兜底
- **副标题**: 填入模板 subtitle 占位符
- **主正文**: 来自同一个 shape → 填入 body 占位符；来自多个 shape → 按原位置回填
- **侧边栏**: 优先匹配右半部分占位符，无则按原位置回填
- **图片/表格**: 按原位置添加，保留内容
- **额外形状**: 保留原填充色和边框色（用于 color pairing 视觉效果）

---

## 五、技术栈

- **语言**: Python 3.10+
- **PPT 操作**: python-pptx 0.6.23+
- **XML 解析**: lxml 4.9.3+
- **数据验证**: pydantic 2.6.1+
- **前端**: Streamlit
- **测试**: pytest 7.4.4+
- **部署**: Streamlit Cloud

---

## 六、已完成的主要工作

### P0 - 核心功能
- ✅ 模板分析与风格识别（4种风格全部正确识别）
- ✅ 内容抽取与重放
- ✅ 表格完整样式保留
- ✅ 图表完整样式保留
- ✅ 字体遵循模板（标题/正文字体大小颜色匹配模板）
- ✅ 16:9 强制输出校验
- ✅ 水印检测与去除

### P1 - 增强功能
- ✅ 文本溢出检测
- ✅ 主题颜色映射
- ✅ 自选图形支持
- ✅ 模板持久化（记住上次上传的模板）
- ✅ 公司图标右下角定位

### P2 - 架构优化（最新）
- ✅ 内容与格式分离架构
- ✅ TextFormat / ShapeFormat 数据模型
- ✅ shape_id 精确映射与去重
- ✅ 格式继承解析（run级 None 值递归获取实际值）
- ✅ 模板格式优先、原格式兜底的回填策略
- ✅ 多文本框按原位置回填，保留原PPT布局
- ✅ 页眉页脚不提取原PPT内容
- ✅ 语义角色智能匹配（body_sidebar 优先右半部分）
- ✅ 对比测试脚本（内容完整性验证）

### 质量保障
- ✅ 结构校验
- ✅ 样式校验
- ✅ 占位符校验
- ✅ 资源校验
- ✅ 溢出校验
- ✅ 转换成功率对比测试

---

## 七、已知问题 & 待优化

### 测试相关
- ✅ 所有21个测试通过（含新增的对比测试）

### 功能待完善
- ⚠️ ShapeFormat 未完全整合到回填决策中（位置/大小仍走 shape_data 字典）
- ⚠️ 图表格式在某些情况下可能不完全匹配
- ⚠️ 深绿色和渐变风格的背景通过主题名称推断，非直接从Master背景解析
- ⚠️ 复杂 SmartArt 未处理

### 可优化方向
- 🔧 格式继承完整解析：从 Master → Layout → Placeholder → Run 全链路
- 🔧 智能布局选择：基于源PPT布局特征自动选择最匹配的模板布局
- 🔧 多语言支持
- 🔧 批量处理能力

---

## 八、快速开始

```bash
# 进入项目目录
cd /workspace/ppt-conform

# 安装依赖
pip install -r requirements.txt

# 启动 Streamlit
streamlit run app.py

# 运行全部测试
pytest tests/ -v

# 运行对比测试（内容完整性验证）
python tests/test_conversion_compare.py

# 命令行使用
python -m cli.main --input source.pptx --style F1 --output output.pptx
```

---

## 九、Git 状态

```
当前分支: master
远程仓库: origin (https://github.com/ccjie168/ppt-conform.git)
最新提交: d66b462 (feat: 解决内容与格式割裂问题，转换成功率达100%)
```

**最近提交记录**:
1. `d66b462` - feat: 解决内容与格式割裂问题，转换成功率达100%
2. `8a4b5a5` - feat: 内容与格式分离架构，模板格式优先+原格式兜底
3. `4febeb7` - fix: 修复4种风格识别问题，主题名称优先策略
4. `db8ae26` - add template file
5. `80f36ad` - fix: AttributeError 'Slide' object has no attribute 'slide'
6. `a027d77` - enhance: per-master theme extraction and layout-level background detection
7. `c377b3b` - fix: recognize all 4 Schneider style variants from slide-level backgrounds

---

## 十、下次继续建议

### 开始前
1. 拉取最新代码: `git pull origin master`
2. 确认模板文件存在: `ls templates/se_energy_tech_ppt_20260421.pptx`
3. 运行测试确认基线: `pytest tests/ -v`

### 优先可做的改进
1. **ShapeFormat 整合**: 将形状位置/大小也纳入"模板优先、原格式兜底"逻辑
2. **格式继承完整化**: 解析 Master→Layout→Placeholder 全链路继承
3. **智能布局匹配**: 基于源PPT布局特征自动选择最匹配的模板布局
4. **真实PPT验证**: 用用户真实PPT做端到端测试，收集失败案例
5. **SmartArt 支持**: 添加对 SmartArt 图形的转换支持

### 启动开发调试
```bash
cd /workspace/ppt-conform
streamlit run app.py
```

### 调试核心转换流程
```python
from core.extractor.pptx_extractor import PptxExtractor
from core.registry.template_registry import TemplateRegistry
from core.replayer.content_replayer import ContentReplayer
from core.models import UserConfig

# 提取
extractor = PptxExtractor()
models = extractor.extract("source.pptx")

# 转换
registry = TemplateRegistry()
replayer = ContentReplayer(registry, template_path="template.pptx")
config = UserConfig(input_path="source.pptx", output_path="output.pptx", master_style="F1")
replayer.replay(models, config)
```

---

## 十一、核心文件说明

| 文件 | 作用 | 关键修改点 |
|------|------|------------|
| `core/models.py` | 数据模型 | TextFormat/ShapeFormat/ContentBlock.raw_shape_id |
| `core/analyzer/template_analyzer.py` | 模板分析与风格识别 | 主题名称优先策略、bg1/bg2映射 |
| `core/extractor/pptx_extractor.py` | PPT内容抽取 | shape_id映射、格式继承解析、页眉页脚跳过 |
| `core/replayer/content_replayer.py` | 内容重放到模板 | shape_id精确去重、模板格式优先+原格式兜底、多shape按原位置回填 |
| `core/validator/validator.py` | 质量校验 | 结构/样式/溢出校验 |
| `core/registry/template_registry.py` | 模板注册表 | Master/Layout管理 |
| `app.py` | Streamlit前端 | 用户交互、模板持久化 |
| `tests/test_conversion_compare.py` | 对比测试 | 内容完整性验证、成功率统计 |

---

## 十二、重要配置说明

### 16:9 宽屏尺寸
- 宽度: 12192000 EMU (13.33 英寸)
- 高度: 6858000 EMU (7.5 英寸)
- 比例: 16:9 (约 1.778)

### 施耐德品牌色
- 深绿色: `#0A2F24`
- 浅绿色: `#E7FFD9`
- 亮绿色: `#3DCD58`
- 白色: `#FFFFFF`

### 页眉页脚区域阈值
- 页眉区域: 顶部 8%
- 页脚区域: 底部 12%

---

*本文档用于新任务快速上手参考，每次重大更新后请同步更新此文档*
