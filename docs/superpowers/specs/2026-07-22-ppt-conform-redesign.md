# PPT 转换工具重构设计文档

> **日期**: 2026-07-22
> **状态**: 待审批
> **作者**: AI 辅助设计

---

## 1. 背景与目标

### 1.1 当前问题

现有 PPT 转换工具存在以下核心问题：

1. **模板上传功能冗余** — 公司专属模板应内置，无需每次上传
2. **深绿色背景显示为白色** — 根因是未解析 Master 的 `clrMap` 颜色映射表
3. **尺寸不一致** — 输出 PPT 保留原 PPT 尺寸，未对齐模板尺寸
4. **幻灯片未使用新母版版式** — 所有页面仍引用原 PPT 的 Layout
5. **缺乏页面类型识别** — 所有页面统一处理，未按类型映射到对应 Layout
6. **缺乏 QA 报告** — 转换后无问题检查报告

### 1.2 新模板分析

内置模板：`2026 se template eng.pptx`

| 属性 | 值 |
|------|-----|
| 尺寸 | 9,144,000 × 5,143,500 EMU (16:9) |
| Master 数量 | 4 |
| 每个 Master 的 Layout 数量 | 25（共 100） |
| Theme 数量 | 4（共享同一色板，名称不同） |
| 字体集 | 1 套（Poppins + 汉仪旗黑） |

**4 个 Master 对应 4 种背景风格：**

| Master | Theme 名称 | 背景 XML | clrMap 映射 | 实际背景色 |
|--------|-----------|---------|------------|-----------|
| 1 | White content slides | `scheme:bg1` | `bg1→lt1` | #FFFFFF（白色） |
| 2 | Light green content slides | `scheme:accent5` | `bg1→lt1` | #E7FFD9（浅绿色） |
| 3 | Dark green content slides | `scheme:bg1` | **`bg1→dk1`** | **#0A2F24（深绿色）** |
| 4 | Gradient content slides | `scheme:bg1` + 全屏渐变图片 | `bg1→dk1` | 渐变图片覆盖 |

**关键发现：** Master 3 的深绿色背景通过 `clrMap` 的 `bg1→dk1` 映射实现（颜色翻转：深色背景 + 浅色文字）。之前代码未解析 `clrMap`，直接将 `bg1` 映射为 `lt1`（白色），导致深绿色背景显示为白色。

### 1.3 设计目标

- 模板内置，取消上传功能
- 正确解析 `clrMap`，4 种背景风格均能正常显示
- 实现 Pre-check 分析、Slide Classification、Layout Mapping
- 采用混合模式：简单页面走 Content Migration，复杂页面走技术适配
- 生成 QA 报告（conversion_report.xlsx）
- 水印、旧 Logo、旧页脚不会带入新 PPT

---

## 2. 总体架构

```
用户上传 PPT
      │
      ▼
┌─────────────────┐
│  Pre-check 分析  │  → 生成源 PPT 质量报告（页面数、母版数、字体清单等）
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  背景风格选择     │  → White / Light Green / Dark Green / Gradient（4 选 1）
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│ Slide Classification │  → 逐页识别类型（Cover/Section/Summary/Table/Chart/Closing 等）
└────────┬────────┘
         │
         ▼
┌──────────┬──────────┐
│ 简单页面  │  复杂页面  │
│ Migration │  Adapt   │
└──────────┴──────────┘
         │
         ▼
┌─────────────────┐
│ Style Normalization │  → 字体替换、颜色统一、页脚/Logo 规范化
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│   QA Report     │  → 输出 conversion_report.xlsx + PPTX
└─────────────────┘
```

---

## 3. 模块详细设计

### 3.1 模板内置与 clrMap 感知

#### 3.1.1 模板内置

- 将 `2026 se template eng.pptx` 作为程序内置资源，路径：`ppt-conform/templates/2026 se template eng.pptx`
- 删除 UI 中的模板上传组件和"使用上次模板"功能
- 4 个背景选项硬编码在 UI 中，直接对应 Master 0-3

#### 3.1.2 clrMap 解析器

新增 `ClrMapResolver` 类，负责正确解析主题颜色引用：

```python
class ClrMapResolver:
    """解析 Master 的 clrMap，将 scheme 颜色名映射为实际 RGB 值"""

    def __init__(self, template_path: str, master_index: int):
        # 1. 读取 Master 的 <p:clrMap> 元素
        # 2. 读取关联 Theme 的 <a:clrScheme>
        # 3. 构建 scheme_name → RGB 映射表

    def resolve_scheme_color(self, scheme_name: str) -> str:
        """将 scheme:bg1 / scheme:dk1 等解析为实际 RGB 值

        示例（Master 3 Dark Green）:
          clrMap = {bg1: dk1, tx1: lt1, ...}
          theme  = {dk1: 0A2F24, lt1: FFFFFF, ...}
          
          resolve("bg1") → clrMap["bg1"] = "dk1" → theme["dk1"] = "0A2F24"
          resolve("tx1") → clrMap["tx1"] = "lt1" → theme["lt1"] = "FFFFFF"
        """
```

**解析流程：**

1. 读取 Master XML 中的 `<p:clrMap>` 属性，获取 `bg1→?`、`tx1→?` 等映射
2. 读取关联 Theme XML 中的 `<a:clrScheme>`，获取每个颜色名的 RGB 值
3. 当遇到 `scheme:bg1` 时：先查 `clrMap` 得到目标色名（如 `dk1`），再查 Theme 得到 RGB（如 `0A2F24`）

**4 个 Master 的解析结果：**

| Master | scheme:bg1 解析 | scheme:tx1 解析 | 背景色 | 文字色 |
|--------|----------------|----------------|--------|--------|
| 1 White | bg1→lt1→#FFFFFF | tx1→dk1→#0A2F24 | 白色 | 深绿 |
| 2 Light Green | bg1→lt1→#FFFFFF | tx1→dk1→#0A2F24 | 浅绿(accent5) | 深绿 |
| 3 Dark Green | **bg1→dk1→#0A2F24** | **tx1→lt1→#FFFFFF** | **深绿** | **白色** |
| 4 Gradient | bg1→dk1→#0A2F24 | tx1→lt1→#FFFFFF | 渐变图片 | 白色 |

### 3.2 Pre-check 分析

#### 3.2.1 分析项

打开源 PPT 后自动分析以下内容，在 UI 上展示报告：

| # | 检查项 | 方法 | 输出 |
|---|--------|------|------|
| 1 | 页面数量 | `len(prs.slides)` | 数字 |
| 2 | 母版数量 | `len(prs.slide_masters)` | 数字 |
| 3 | 使用字体清单 | 遍历所有 slide 的 text_frame | 字体列表 |
| 4 | 是否含旧 SE 模板 | 检查母版名称/主题名称是否含 "Schneider"/"SE" | 是/否 |
| 5 | 是否含外部主题 | 检查 theme 数量 > 1 且名称非标准 | 是/否 |
| 6 | 是否含嵌入 Excel 图表 | 检查 shape 类型是否含 GRAPHIC_FRAME + EMBEDDED | 是/否 |
| 7 | 是否含 SmartArt | 检查是否有 SmartArt XML 部分 | 是/否 |
| 8 | 是否含视频/音频 | 检查 shape 类型 MEDIA | 是/否 |
| 9 | 是否含动画 | 检查 slide XML 中的 `<p:timing>` | 是/否 |
| 10 | 是否有 4:3 页面 | 检查 `prs.slide_width / prs.slide_height` 比例 | 是/否 |
| 11 | 是否存在超出页面边界的对象 | 遍历 shape，检查 `left+width > slide_width` | 是/否 + 数量 |

#### 3.2.2 UI 展示

Pre-check 结果以卡片形式展示在转换配置区上方：
- 绿色 ✅ = 正常
- 黄色 ⚠️ = 警告（可继续转换）
- 红色 ❌ = 严重问题（建议用户修复后再转换）

### 3.3 Slide Classification

#### 3.3.1 页面类型定义

| 类型 | 说明 |
|------|------|
| Cover | 封面页 |
| Section Divider | 章节分隔页 |
| Executive Summary | 执行摘要页 |
| KPI Dashboard | KPI 仪表盘 |
| Timeline | 时间线页 |
| Risk Matrix | 风险矩阵页 |
| Decision Page | 决策页 |
| Table Page | 表格页 |
| Chart Page | 图表页 |
| Image Page | 图片页 |
| Appendix | 附录页 |
| Closing | 结尾页 |
| Content | 通用内容页（默认） |

#### 3.3.2 启发式识别规则

按优先级从高到低匹配：

```python
def classify_slide(slide, slide_index, total_slides):
    # 1. 封面：第1页 + 只有标题/副标题占位符
    if slide_index == 0 and has_only_title_subtitle(slide):
        return "Cover"
    
    # 2. 结尾：最后一页 + 标题/内容含关键词
    if slide_index == total_slides - 1:
        text = get_all_text(slide)
        if any(kw in text for kw in ["se.com", "Thank", "谢谢", "Q&A"]):
            return "Closing"
    
    # 3. 章节分隔：只有大标题，无正文内容
    if has_only_title(slide) and not has_body_content(slide):
        return "Section Divider"
    
    # 4. 摘要：标题含关键词
    title = get_title_text(slide)
    if any(kw in title for kw in ["Executive Summary", "Key Messages", "总结", "摘要", "Overview"]):
        return "Executive Summary"
    
    # 5. 风险页：标题含关键词
    if any(kw in title for kw in ["Risk", "Issue", "Mitigation", "风险", "问题", "应对"]):
        return "Risk Matrix"
    
    # 6. 时间线：标题含关键词
    if any(kw in title for kw in ["Timeline", "Milestone", "Roadmap", "时间线", "里程碑", "路线图"]):
        return "Timeline"
    
    # 7. 表格页：含表格对象
    if has_table(slide):
        return "Table Page"
    
    # 8. 图表页：含图表对象
    if has_chart(slide):
        if count_charts(slide) >= 2:
            return "KPI Dashboard"
        return "Chart Page"
    
    # 9. 图片页：主要是图片
    if is_image_dominant(slide):
        return "Image Page"
    
    # 10. 默认
    return "Content"
```

#### 3.3.3 Migration vs Adaptation 判定

根据页面类型决定走哪条路径：

| 页面类型 | 路径 | 原因 |
|---------|------|------|
| Cover | Migration | 简单：标题+副标题 |
| Section Divider | Migration | 简单：只有标题 |
| Closing | Migration | 简单：固定内容 |
| Executive Summary | Adaptation | 可能含复杂格式 |
| Content | Adaptation | 可能含混合内容 |
| Table Page | Adaptation | 表格对象迁移困难 |
| Chart Page | Adaptation | 图表对象迁移困难 |
| KPI Dashboard | Adaptation | 多图表对象 |
| Risk Matrix | Adaptation | 可能含复杂图形 |
| Timeline | Adaptation | 可能含 SmartArt |
| Image Page | Migration | 图片可复制 |
| Appendix | Adaptation | 内容不确定 |

### 3.4 Layout Mapping

每种页面类型映射到目标 Master 的对应 Layout：

| 页面类型 | Layout 索引 | Layout 名称 |
|---------|------------|------------|
| Cover | 0 | Title slide simple |
| Section Divider | 20 | Section break for longer copy |
| Executive Summary | 1 | One column |
| KPI Dashboard | 9 | One column with object |
| Timeline | 1 | One column |
| Risk Matrix | 1 | One column |
| Table Page | 9 | One column with object |
| Chart Page | 9 | One column with object |
| Image Page | 4 | Title slide with image |
| Appendix | 1 | One column |
| Closing | 24 | Closing slide |
| Content | 1 | One column |

> 注：Layout 索引是每个 Master 内的 0-24 序号，不同 Master 的同名 Layout 索引相同。

### 3.5 Content Migration（简单页面）

**流程：**

1. 基于目标 Master + 目标 Layout 创建新幻灯片
2. 提取旧页面的文本内容（标题、副标题、正文文本）
3. 将文本放入新 Layout 的对应占位符
4. 丢弃旧页面的所有形状、背景、Logo、页脚

**实现方式：**

```python
def migrate_slide(source_slide, target_prs, target_layout, slide_type):
    """基于目标 Layout 创建新幻灯片，迁移文本内容"""
    
    # 1. 创建新幻灯片
    new_slide = target_prs.slides.add_slide(target_layout)
    
    # 2. 提取源页面文本
    title_text = get_title_text(source_slide)
    body_texts = get_body_texts(source_slide)
    
    # 3. 填入占位符
    if new_slide.shapes.title:
        new_slide.shapes.title.text = title_text
    
    for placeholder, text in zip_body_placeholders(new_slide, body_texts):
        placeholder.text = text
    
    # 4. 如是 Image Page，复制图片
    if slide_type == "Image Page":
        copy_images(source_slide, new_slide)
    
    return new_slide
```

**水印防护：** Migration 路径只提取文本内容，旧页面的所有形状（包括水印）被物理丢弃。

### 3.6 Technical Adaptation（复杂页面）

**流程：**

1. 保留原幻灯片
2. 将幻灯片的 Layout 引用切换到目标 Master 的对应 Layout（修改 slide rels）
3. 使用 clrMap 解析正确的背景色，直接设置到幻灯片
4. 删除旧 Logo、旧页脚、旧页码（通过位置和类型识别）
5. 删除水印（通过关键词匹配和图片透明度检测）
6. 应用字体和颜色标准化

**Layout 切换方式：**

通过 zipfile 修改 `ppt/slides/_rels/slide{N}.xml.rels` 中的 slideLayout Target，指向新 Master 的 Layout 文件。

**背景色应用：**

```python
def apply_background(slide, clr_map_resolver, master_index):
    """通过 clrMap 解析正确的背景色并应用"""
    
    # 解析 scheme:bg1 的实际 RGB
    bg_color = clr_map_resolver.resolve_scheme_color("bg1")
    
    # 如果是 Gradient Master (index 3)，复制全屏渐变图片
    if master_index == 3:
        copy_gradient_background(slide)
    else:
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor.from_string(bg_color)
```

**水印防护：** Adaptation 路径主动检测并删除水印（关键词匹配 + 图片透明度检测 + 页脚区域形状清除）。

### 3.7 Style Normalization

#### 3.7.1 字体替换

| 原字体 | 目标字体 | 说明 |
|--------|---------|------|
| 微软雅黑 / Microsoft YaHei | 汉仪旗黑 | 中文字体替换 |
| Calibri / Arial | Poppins | 英文字体替换 |
| 其他中文字体 | 汉仪旗黑 | 统一中文 |
| 其他英文字体 | Poppins | 统一英文 |

字体配置在 `config/master_styles.yaml` 中可调。

#### 3.7.2 颜色统一

| 元素 | 颜色 | 来源 |
|------|------|------|
| SE Green（强调色） | #3DCD58 | 主题 dk2 |
| Dark Green（深色背景） | #0A2F24 | 主题 dk1 |
| Light Green（浅色背景） | #E7FFD9 | 主题 accent5 |
| White | #FFFFFF | 主题 lt1 |

Dark Green 风格下：标题白色、正文白色/浅绿色；其他风格下：标题深绿色、正文深灰色。

#### 3.7.3 页脚和 Logo

- 删除旧 PPT 的所有页脚元素（页脚占位符、页码、底部小文本框、底部小图片）
- 使用新 Master 的页脚和 Logo（通过 Layout 继承，不需要手动添加）

### 3.8 QA Report

#### 3.8.1 输出文件

生成 `conversion_report.xlsx`，与 PPTX 一起提供给用户下载。

#### 3.8.2 报告字段

| 字段 | 类型 | 说明 |
|------|------|------|
| Slide No. | int | 幻灯片序号 |
| Detected Type | string | 检测到的页面类型 |
| Applied Layout | string | 使用的目标 Layout 名称 |
| Migration Mode | string | Migration / Adaptation |
| Font Replaced | string | 替换了哪些字体（如 "微软雅黑→汉仪旗黑"） |
| Objects Moved | int | 重新定位的对象数量 |
| Objects Deleted | int | 删除的对象数量（含水印、旧Logo等） |
| Overflow Risk | string | None / Low / Medium / High |
| Need Manual Review | bool | 是否需要人工检查 |
| Comment | string | 详细说明 |

#### 3.8.3 检查项

| # | 检查项 | 判定方式 | 严重程度 |
|---|--------|---------|---------|
| 1 | 文本溢出 | 文本框高度超出页面边界 | High |
| 2 | 图表压缩 | 图表宽高比变化超过 20% | Medium |
| 3 | 图片变形 | 图片宽高比变化超过 10% | Medium |
| 4 | 旧母版残留 | 检查输出 PPT 是否有非目标 Master | High |
| 5 | 旧字体残留 | 检查是否仍有微软雅黑/Calibri | Medium |
| 6 | 对象超出页面 | shape.left+width > slide_width | Medium |
| 7 | 空白页 | 无任何可见内容的页面 | Low |
| 8 | 页脚重复 | 底部区域有多个页脚元素 | Medium |
| 9 | Logo 重复 | 多个 Logo 图片 | Medium |

---

## 4. UI 变更

### 4.1 删除的功能

- 模板上传区域（file_uploader for template）
- "使用上次" 按钮模板加载逻辑
- 模板分析缓存
- 模板宽高比检查

### 4.2 新增/修改的功能

- **背景风格选择**：4 个固定选项（White / Light Green / Dark Green / Gradient），不再依赖模板分析
- **Pre-check 报告区**：上传 PPT 后自动展示分析结果
- **QA 报告下载**：转换完成后提供 PPTX + xlsx 两个下载按钮

### 4.3 UI 流程

```
1. 用户上传 PPT 文件
2. 自动执行 Pre-check，展示分析报告
3. 用户选择背景风格（4选1）
4. 点击"开始转换"
5. 转换完成，展示结果摘要
6. 下载 PPTX + conversion_report.xlsx
```

---

## 5. 文件结构变更

### 5.1 新增文件

```
ppt-conform/
├── core/
│   ├── precheck/
│   │   ├── __init__.py
│   │   └── analyzer.py          # Pre-check 分析器
│   ├── classifier/
│   │   ├── __init__.py
│   │   └── slide_classifier.py  # Slide Classification
│   ├── clrmap/
│   │   ├── __init__.py
│   │   └── resolver.py          # clrMap 解析器
│   └── qa/
│       ├── __init__.py
│       └── reporter.py          # QA 报告生成器
├── config/
│   └── font_mapping.yaml        # 字体替换配置
└── templates/
    └── 2026 se template eng.pptx  # 内置模板（已存在）
```

### 5.2 修改的文件

| 文件 | 变更内容 |
|------|---------|
| `app.py` | 删除模板上传，增加 Pre-check 展示和 QA 报告下载 |
| `core/replayer/content_replayer.py` | 集成 clrMap 解析、混合模式转换逻辑 |
| `core/models.py` | 增加 PreCheckResult、SlideClassification、QAReport 等数据模型 |
| `config/master_styles.yaml` | 更新为与新模板匹配的配置 |
| `static/schneider_style.css` | UI 样式调整 |

### 5.3 可删除的文件

- `core/adapter/style_adapter.py` — 功能合并到 content_replayer
- 旧的模板文件 `se_energy_tech_ppt_20260421.pptx` — 被新模板替代

---

## 6. 数据模型

```python
class PreCheckResult(BaseModel):
    slide_count: int
    master_count: int
    fonts_used: list[str]
    has_old_se_template: bool
    has_external_theme: bool
    has_embedded_chart: bool
    has_smartart: bool
    has_media: bool
    has_animation: bool
    is_4_3_ratio: bool
    overflow_objects_count: int
    issues: list[PreCheckIssue]

class SlideClassification(BaseModel):
    slide_index: int
    slide_type: str           # Cover / Section / Summary / Table / Chart / ...
    migration_mode: str       # "migration" / "adaptation"
    target_layout_index: int
    confidence: float

class QAReportItem(BaseModel):
    slide_no: int
    detected_type: str
    applied_layout: str
    migration_mode: str
    font_replaced: str
    objects_moved: int
    objects_deleted: int
    overflow_risk: str        # None / Low / Medium / High
    need_manual_review: bool
    comment: str

class ConversionConfig(BaseModel):
    input_path: str
    output_path: str
    report_path: str
    background_style: int     # 0=White, 1=LightGreen, 2=DarkGreen, 3=Gradient
```

---

## 7. 关键技术决策

### 7.1 为什么用混合模式而非纯 Migration

- **图表/表格/SmartArt** 迁移到新 Layout 时容易丢失数据绑定、格式、动画
- **纯文字页面** 用 Migration 能得到最干净的结果（无旧母版、无旧格式残留）
- 混合模式在保证质量的同时降低风险

### 7.2 为什么用 clrMap 而非硬编码

- `clrMap` 是 OOXML 标准的颜色映射机制，PowerPoint 自身也依赖它
- 正确解析 `clrMap` 后，4 种背景风格都能从模板继承正确颜色，无需任何硬编码
- 如果未来模板更新颜色，代码无需修改

### 7.3 为什么修改 slide rels 而非用 python-pptx API

- python-pptx 不支持将已有 slide 重新关联到不同的 Layout
- 通过 zipfile 修改 `slide{N}.xml.rels` 中的 Target 是最直接的方式
- 这是社区公认的做法（参考 Aspose.Slides、python-pptx Issues）

---

## 8. 水印防护策略

### 8.1 Content Migration 路径

- 只提取文本内容，旧页面的所有形状（含水印）被物理丢弃
- 水印**不可能**出现在新页面中

### 8.2 Technical Adaptation 路径

- **文本水印**：通过关键词匹配（17 个关键词 + 3 个正则模式）识别并删除
- **图片水印**：通过 alpha 通道透明度检测识别并删除
- **页脚区域清除**：底部 15% 区域的小形状（文本框/图片）全部删除
- **旧 Logo**：通过位置（右下角）和尺寸识别并删除

### 8.3 QA 兜底

- QA Report 中检查是否有水印关键词残留
- 如果检测到，标记 "Need Manual Review"

---

## 9. 实施计划

### Phase 1: 基础设施

1. 创建 `ClrMapResolver`，正确解析 4 种背景色
2. 更新 `master_styles.yaml` 配置
3. 修改 `content_replayer.py` 集成 clrMap 解析
4. 删除模板上传功能，模板内置

### Phase 2: 分析与分类

5. 实现 `PreCheckAnalyzer`，生成源 PPT 分析报告
6. 实现 `SlideClassifier`，逐页识别类型
7. 实现 Layout Mapping 逻辑

### Phase 3: 混合模式转换

8. 实现 Content Migration 路径（简单页面）
9. 重构 Technical Adaptation 路径（复杂页面）
10. 集成两条路径的统一入口

### Phase 4: 样式标准化

11. 实现字体替换（汉仪旗黑 + Poppins）
12. 实现颜色统一
13. 实现页脚/Logo 清除与规范化

### Phase 5: QA 报告

14. 实现 `QAReporter`，生成 conversion_report.xlsx
15. 实现 QA 检查项（9 项检查）

### Phase 6: UI 更新

16. 更新 `app.py`，删除模板上传，增加 Pre-check 展示
17. 增加 QA 报告下载
18. 样式调整

---

## 10. 验收标准

| # | 验收项 | 验收标准 |
|---|--------|---------|
| 1 | 模板内置 | UI 无模板上传入口，4 个背景选项直接可用 |
| 2 | 深绿色背景 | 选择 Dark Green 后，所有幻灯片背景为 #0A2F24 |
| 3 | 尺寸一致 | 输出 PPT 尺寸为 9144000×5143500 EMU |
| 4 | 版式引用 | 所有幻灯片引用新 Master 的 Layout |
| 5 | 水印清除 | 输出 PPT 中无任何水印文本或图片 |
| 6 | 字体替换 | 输出 PPT 中无微软雅黑/Calibri |
| 7 | Pre-check | 上传 PPT 后展示 11 项分析结果 |
| 8 | 页面分类 | 能识别 Cover/Section/Table/Chart/Closing 等类型 |
| 9 | QA 报告 | 生成 conversion_report.xlsx，含 10 个字段 |
| 10 | 旧 Logo 清除 | 输出 PPT 中无旧 SE Logo |
