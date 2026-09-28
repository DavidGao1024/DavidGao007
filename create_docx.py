# -*- coding: utf-8 -*-
import os
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

OUT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), '高佳顺简历.docx')

doc = Document()

# 设置默认字体
doc.styles['Normal'].font.name = '微软雅黑'
doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
doc.styles['Normal'].font.size = Pt(10.5)


def bullet(text, bold_prefix=None):
    """添加一条项目符号段落，bold_prefix 会加粗"""
    p = doc.add_paragraph(style='List Bullet')
    if bold_prefix:
        p.add_run(bold_prefix).bold = True
    p.add_run(text)
    return p


def labeled(label, text):
    """添加一条 '• 标签：内容' 段落"""
    p = doc.add_paragraph()
    p.add_run(f'• {label}：').bold = True
    p.add_run(text)
    return p


def make_table(headers, rows, widths=None):
    """创建 Table Grid 表格，首行加粗"""
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = 'Table Grid'
    for j, h in enumerate(headers):
        cell = table.rows[0].cells[j]
        cell.text = h
        cell.paragraphs[0].runs[0].bold = True
    for i, row_data in enumerate(rows, start=1):
        for j, cell_text in enumerate(row_data):
            table.rows[i].cells[j].text = cell_text
    return table


# ==================== 标题 ====================
title = doc.add_heading('高佳顺', 0)
title.alignment = WD_ALIGN_PARAGRAPH.CENTER

subtitle = doc.add_paragraph('高级前端工程师（AI 应用方向） | 20 年经验')
subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER

contact = doc.add_paragraph()
contact.alignment = WD_ALIGN_PARAGRAPH.CENTER
contact.add_run('电话：13752099426 | 邮箱：670646790@qq.com | 期望薪资：15-25K | 期望城市：天津')

# ==================== 核心优势 ====================
doc.add_heading('核心优势', level=1)
advantages = [
    ('AI 应用落地', '在医疗教育产品线完成 AI 问诊、AI 伴学、AI 生成病例、AI 语音问诊等能力的前端实现；主导 AI 老年学生端数字人交互（uni-app + hybrid），设计数字人加载时序与作答前交互编排'),
    ('AI 研发效能体系（自建）', '自主研发 3 个 Agent Skill（产品经理 / Bug 根因分析 / 批量导出）与端到端回归 Agent，坚持「先定文档再实施」的规格驱动开发；以浏览器自动化实测取证，形成可追溯验证链，已在 4 条产品线落地'),
    ('产品线主导力', '4 条医疗教育产品线的头号开发者——中医临床思维（333 次提交）、3D 医疗模拟（170 次）、用户中心（142 次）、慢病诊疗思维（98 次），均为全仓第一'),
    ('3D Web 开发', '6 年+ Vue / Three.js 开发经验，主导多个医疗 3D 模拟项目，累计销售额超 2000 万；掌握 glTF 骨骼动画、三维射线拾取、Electron 与 Unity WebGL 双向通信'),
    ('跨端协作与数据模型', 'uni-app 一套代码多端交付（H5 / 微信小程序 / 触屏一体机）；参与前后端共享数据模型层（TypeScript）开发，熟悉接口契约设计与前后端协作'),
    ('团队管理', '具备项目计划制定、外包模型品质把控、新人培训与团队激励制度建设经验'),
]
for label_text, desc in advantages:
    labeled(label_text, desc)

# ==================== 技术栈 ====================
doc.add_heading('技术栈', level=1)
make_table(
    ['类别', '技术'],
    [
        ('核心技术', 'Vue3 / Vue2、TypeScript、JavaScript、Three.js'),
        ('跨端开发', 'uni-app（H5 / 微信小程序 / 触屏一体机）、Electron、Android & iOS SDK 集成'),
        ('构建与状态', 'Vite、Webpack、Pinia、VueX、UnoCSS'),
        ('UI 与组件', 'Element Plus、uni-ui、自定义组件库'),
        ('数据可视化', 'ECharts、Three.js、Cornerstone.js（DICOM 医学影像）'),
        ('AI 集成', 'OpenAI 兼容接口、Qwen 系列模型、提示词工程、SSE 流式输出、敏感词合规过滤'),
        ('AI 研发工具', 'Claude Code、MCP（chrome-devtools / playwright）、Agent Skill、SubAgent、自定义命令、规范驱动开发（Spec-Driven）'),
        ('数据模型与协作', 'Node.js、Express 架构理解、MongoDB / 达梦 DMDB、Redis、接口契约设计'),
        ('游戏引擎（历史）', 'CocosCreator、Cocos2dx、Lua、Egret'),
    ]
)

# ==================== AI 能力与实践 ====================
doc.add_heading('AI 能力与实践', level=1)

doc.add_heading('A. AI 功能落地（产品级）', level=2)

doc.add_heading('中医临床思维平台（主导）', level=3)
for item in [
    ('界面融合', '主导教师端 AI 助手入口的界面翻新与布局避让处理，保证 AI 入口与原有教学流程兼容'),
]:
    labeled(item[0], item[1])

doc.add_heading('临床思维平台（西医线）', level=3)
for item in [
    ('AI 问诊', '实现 AI 问诊与传统问诊双模式分流，支持病例级「支持 AI 问诊」筛选与作答'),
    ('AI 伴学', '拼装患者信息、病历摘要、标准答案与学生作答，实时生成分阶段引导（问诊 / 体格检查 / 辅助检查三类）'),
    ('AI 生成病例', '教师端简略 / 详细双模式生成，并支持章节级（病史采集、体格检查、辅助检查、治疗计划）生成'),
    ('AI 语音问诊', '基于 getUserMedia 录音并导出 WAV，对接语义识别接口完成语音问诊'),
]:
    labeled(item[0], item[1])

doc.add_heading('AI 老年学生端 · 数字人（uni-app）', level=3)
bullet('主导 hybrid 数字人集成，设计「数字人加载 → 进度条反馈 → 对话框呈现 → 开始作答」的加载时序与交互编排')
bullet('重构接口调用层，将原生 uni.request 统一封装为 Promise 链式请求模块，显著减少重复逻辑')
bullet('补齐数据加解密（crypto-js / js-md5）与多环境配置切换能力')

doc.add_heading('AI 助手组件集成', level=3)
bullet('集成 AI 助手组件：OpenAI 兼容自建通道支持，支持后端动态下发模型与系统提示词')
bullet('接入敏感词合规过滤，满足业务内容安全要求')

doc.add_heading('B. AI 研发效能体系（自建）', level=2)
doc.add_heading('自研 Agent Skill（3 个）', level=3)
make_table(
    ['Skill', '定位与亮点'],
    [
        ('product-manager', '产品经理 Agent：六阶段流程（需求访谈 → 现状诊断 → 方向提案 → 深化方案 → 原型构建 → 存档交付）+ 两道强制检查点（方向选定、方案批准），未获批不推进；产出《后端需求清单》（接口路径 / 改动类型 / 字段定义 / 出入参示例，可直接转交后端）与可点击 HTML 原型（读项目主题变量注入色板，换肤自动跟随）；职责止于方案与原型，不改业务代码'),
        ('bug-repairer', 'Bug 根因分析 Agent：按 URL 自动识别禅道 / 云效双 Bug 系统；代码优先于复现，按确信度三级分流（100% 直接改 / 80-90% 拦截接口确认 / 低于 80% 才完整复现）；内置 AES-128-ECB 接口解密；前后端归因四条件判定（数据直达 / 同源对照 / 模板无误 / 无遗漏处理）；修复后强制浏览器验证 + 同组件影响面反查'),
        ('exporting-user-data', '批量导出 Agent：一条命令完成登录态复用 → 526 家单位全量用户拉取 → Excel（总索引 + 每单位一页签）→ 回读自校验；5 并发约 2 分钟；校验不通过禁止交付'),
    ]
)

doc.add_heading('SubAgent 与端到端回归', level=3)
bullet('考试流程测试 Agent：自动完成登录 → 考试码入场 → 6 个页签作答（基本信息 / 中医诊断 / 辨病辨证 / 病症鉴别 / 中医治法 / 中药处方）→ 提交 → 成绩核验，逐项比对「实际作答」与「成绩页学生答案」以验证判分逻辑，自动生成含截图与验证清单的 HTML 报告')
bullet('多子代理协作流水线：按「实施 → 规格审查 → 代码质量审查」分工，21 个子代理单轮运行 4 小时产出 641 行有效变更')

doc.add_heading('规格驱动工作流：先与 AI 确认文档，再动手实施', level=3)
bullet('每个功能按「设计规范 → 实施计划 → 分步实施 → 逐项验收」四段推进，每段独立留痕')
bullet('设计文档交付前先做占位符、内部一致性、范围、歧义四项自审；阶段检查点不可逾越')

doc.add_heading('工程化沉淀', level=3)
bullet('建设项目上下文记忆库（六类文档：项目概要 / 产品背景 / 技术上下文 / 系统模式 / 活动上下文 / 进度）与七份团队规范库')
bullet('开发 Vue2 → Vue3 迁移自动化三件套（质量分析 / 批量修复 / 稳定修复），支持 dry-run 与自动备份')
bullet('沉淀自定义命令集，覆盖病例编辑、练习配置、考试流程验证、多角色登录')

doc.add_heading('C. AI 中台协作（参与）', level=2)
bullet('参与 AI 中台接口对接与联调：统一模型网关（OpenAI 兼容协议）、多业务场景配置化、按机构的额度校验与 Token 用量统计')
bullet('参与前后端共享数据模型层（TypeScript）建设，覆盖病例思维模型与接口参数 / 结果契约')

doc.add_heading('D. AI 驱动的全流程研发（一人产品线）', level=2)
p = doc.add_paragraph()
p.add_run('以 AI 为协作主体，单人承担一条产品线从需求到上线的完整研发闭环：')
make_table(
    ['环节', 'AI 协作方式'],
    [
        ('产品功能设计', 'product-manager Skill：需求访谈 → 现状诊断 → 方向提案 → 深化方案，产出《后端需求清单》'),
        ('UI 设计', '同 Skill 的原型构建：读取项目主题变量注入色板，产出可点击高保真原型，换肤自动跟随'),
        ('全链路编码', 'Claude Code + MCP：按「设计规范 → 实施计划」分步实施，规格驱动、逐段留痕'),
        ('测试与改 Bug', 'bug-repairer Skill + 考试回归 Agent：根因定位、四条件归因、修复后浏览器验证、端到端回归自动出报告'),
    ]
)
p = doc.add_paragraph()
p.add_run('单人闭环完成从需求到上线的完整研发链路，大幅压缩跨角色沟通成本。').italic = True

# ==================== 工作经历 ====================
doc.add_heading('工作经历', level=1)

# --- 天堰科技 ---
doc.add_heading('天津天堰科技股份有限公司 | 前端开发工程师 | 2021.10 - 至今', level=2)
p = doc.add_paragraph()
p.add_run('核心业绩：').bold = True
p.add_run('中医临床思维平台全年 333 次提交，为公司该产品线头号开发者；3D 医疗模拟项目累计销售额超 2000 万；中西医临床思维年销售额 1000 万+')

doc.add_heading('一、AI 应用方向', level=3)
bullet('主导临床思维 4.2 改版 ✅已上线：完成全平台界面翻新与 AI 功能集成（AI 语音问诊、AI 伴学、AI 评价、AI 生成病例、AI 小助手随时问答）')
bullet('主导 AI 老年学生端数字人（uni-app）开发：hybrid 数字人集成、加载时序设计、请求层重构')
bullet('自建 AI 研发效能体系并推广至 4 条产品线，覆盖 Bug 自动归因、原型生成、端到端回归与自动化验证')
bullet('沉淀 3 个自研 Agent Skill（产品经理 / Bug 根因分析 / 批量导出）与考试回归 Agent，形成「设计规范 → 实施计划 → 分步实施 → 逐项验收」的规格驱动开发流程')

doc.add_heading('二、产品线主力开发', level=3)
bullet('中医临床思维平台（主导）：全仓 333 次提交（第一），该产品线头号开发者')
for sub_item in [
    '主导教师端病例编辑、考试与练习管理、成绩统计等核心模块',
    '主导冗余代码治理：4 天清理 9,700+ 行、58 个文件',
    '主导构建性能优化：42.17s → 39.23s（提升 7.1%）',
]:
    doc.add_paragraph(sub_item, style='List Bullet 2')
bullet('用户中心 / 订单中台：全仓 142 次提交（第一）。负责组织机构、单位管理、产品订单、权限分配、病例包配置等模块；开发通用 Excel 导出能力并落地订单与病例设置页；支撑 526 家真实单位使用')
bullet('老年慢病诊疗思维训练系统：全仓 98 次提交（第一）。集成 DICOM 医学影像查看器、ECharts 数据可视化与 AI 助手组件；主导健康评估增强、客观资料页改版、评分与 Excel 导出优化')
bullet('参与 lcsw3.4 平台数据模型层（TypeScript）：负责病例思维模型相关模块；参与接口参数与结果契约设计')

doc.add_heading('三、3D 与可视化方向', level=3)
bullet('3D 医疗模拟项目（全仓 170 次提交，第一）：基于 Vue3 + Three.js + Electron 开发桌面端 3D 模拟训练系统，完成 677 个 glTF 模型加载与骨骼动画、三维射线拾取点位交互')
bullet('实现 Electron 与 Unity WebGL 的双向通信集成，打通 3D 场景与 Vue 业务层的交互链路')
bullet('负责 VR 项目 3D 场景搭建、模型管理、灯光控制、动画播放与交互设计全流程；对接模型外包公司，制定输出标准并把控模型品质')

# --- 古阳广告 ---
doc.add_heading('天津古阳广告设计有限公司 | CocosCreator 开发工程师 | 2021.01 - 2021.09', level=2)
p = doc.add_paragraph()
p.add_run('核心业绩：').bold = True
p.add_run('完成暑期小学语文 7 个讲次课件开发，已上线推广使用')
for item in [
    '使用 CocosCreator + TypeScript 开发在线教育网课课件',
    '使用 Vue + VueX + Router 二次开发迈格森教师端项目',
    '负责部门新人培训、管理制度整理、周会主持',
]:
    bullet(item)

# --- 集智创研 ---
doc.add_heading('天津集智创研科技有限公司 | Cocos2dx 开发 | 2019.02 - 2020.12', level=2)
p = doc.add_paragraph()
p.add_run('核心业绩：').bold = True
p.add_run('2019 年 8 省棋牌项目流水 6000 万')
for item in [
    '使用 Cocos2d-X 维护 8 个地区棋牌项目（麻将、斗地主、字牌等）',
    '负责 iOS / Android 双端 SDK 集成（登录、支付、分享、语音、广告）',
    '客户端打包、热更新、签名更换等运维工作',
    '开发环境升级（Xcode 9→11、Eclipse→Android Studio）',
]:
    bullet(item)

# --- 宏诚盛世 ---
doc.add_heading('天津宏诚盛世（天津）网络科技有限公司 | Cocos2dx | 2017.11 - 2018.11', level=2)
p = doc.add_paragraph()
p.add_run('职责：').bold = True
p.add_run('项目管理 + 技术开发')
for item in [
    '承担设计开发阶段项目管理，制定开发计划并保证按期交付',
    '负责手游脚本开发，与 UI、策划协作确定功能方案',
    '团队建设：制定员工奖惩制度、月最佳评选机制',
]:
    bullet(item)

# --- 雅讯时空 ---
doc.add_heading('天津雅讯时空科技发展有限公司 | Cocos2dx | 2015.10 - 2016.12', level=2)
bullet('足球经理手游：完全自主研发，包含 PVP、PVE、联赛、商城、好友、排行榜、转会市场等系统')
bullet('魔力乐消消：三消类游戏，参考同类产品进行玩法创新')

# --- 早期经历 ---
doc.add_heading('早期工作经历（2009 - 2015）', level=2)
make_table(
    ['时间', '公司', '职位', '主要成果'],
    [
        ('2014.11-2015.10', '天津虫象科技', 'Flash 开发', '研发部日常管理、培训计划制定'),
        ('2012.03-2014.05', '摩卡软件', 'Flash 开发', '流程控制软件、监控管理软件'),
        ('2011.11-2012.03', '天津象形科技', 'Flash 开发', 'BS 结构网络游戏开发'),
        ('2010.06-2011.11', '天津企商科技', 'Flash 开发', '真人试衣系统（1 项发明专利 + 6 项实用新型 + 1 项外观设计 + 1 项软著）'),
        ('2009.04-2010.05', '天津安信天诚', 'Flash 开发', 'SNS 网站 Flash 游戏开发'),
    ]
)

# ==================== 代表项目 ====================
doc.add_heading('代表项目', level=1)

projects = [
    ('1. 中医临床思维平台 | 2021.10 - 至今 ✅已上线',
     '产品线头号开发者（全仓 333 次提交，第一）',
     'Vue3 + TypeScript + Vite + Pinia + Element Plus + Cornerstone.js',
     '面向中医教学的一体化平台，覆盖教师端病例编辑、考试与练习管理、成绩统计，学生端临床思维作答与医案训练',
     ['主导教师端病例编辑与考试管理模块重建，单文件从老版 9,800 行重构为 4,978 行',
      '主导冗余代码治理，4 天清理 9,700+ 行、58 个文件，构建耗时提升 7.1%']),
    ('2. AI 研发效能体系（自研） | 2026.01 - 至今',
     '体系设计者与唯一开发者',
     'Claude Code + MCP（Playwright / Chrome DevTools）+ Agent Skill + SubAgent + 自定义命令',
     '针对 Bug 排查耗时、产品方案零散、回归依赖人工三类问题，自建一套 AI 研发效能体系，覆盖 4 条产品线。资产清单：3 个 Agent Skill（产品经理 / Bug 根因分析 / 批量数据导出）+ 1 个考试流程回归 Agent + 5 条自定义命令',
     ['Bug 分析内置双 Bug 系统识别（禅道 / 云效）与前后端归因四条件判定，修复后强制浏览器验证并反查同组件影响面',
      '产品经理 Skill 产出《后端需求清单》与可点击 HTML 原型，方案可直接作为开发任务输入',
      '导出 Skill 实现 526 家单位全量拉取与回读自校验，校验不通过禁止交付',
      '确立「设计规范 → 实施计划 → 分步实施 → 逐项验收」的规格驱动流程，四段独立留痕']),
    ('3. AI 老年学生端 · 数字人（uni-app） | 2025.07 - 2025.08 ✅已交付',
     '主要开发者（31 次提交）',
     'uni-app + Vue3 + Three.js + crypto-js',
     '老年医学思维训练的学生端多端应用，集成 3D 数字人导学',
     ['主导 hybrid 数字人集成，设计「加载 → 进度条 → 对话框 → 开始作答」时序',
      '重构请求层为 Promise 链式封装，大幅减少重复逻辑',
      '一套代码多端交付（H5 / 微信小程序 / 触屏一体机）']),
    ('4. 用户中心 / 订单中台 | 2025.03 - 至今 ✅已上线',
     '头号开发者（全仓 142 次提交，第一）',
     'Vue3 + TypeScript + Vite + Pinia + Element Plus + UnoCSS',
     '公司医学教育产品线的统一用户与权限底座，支撑 526 家真实单位',
     ['负责组织机构、单位管理、产品订单、权限分配、病例包配置',
      '开发通用 Excel 导出能力（订单页 + 病例设置页）',
      '建立 commitlint + lint-staged + 体积分析等工程规范']),
    ('5. 老年慢病诊疗思维训练系统 | 2025.05 - 2025.12 ✅已上线',
     '头号开发者（全仓 98 次提交，第一）',
     'Vue3 + TypeScript + Vite + Pinia + Element Plus + Cornerstone.js + ECharts',
     '医学教育平台，专注慢性病诊疗思维训练，含病史采集、体格检查、诊断推理、治疗计划、医患沟通等模块',
     ['集成 DICOM 医学影像查看器、ECharts 数据可视化、AI 助手组件',
      '主导健康评估增强、客观资料页改版与评分导出优化']),
    ('6. 临床思维 4.2 改版 | 2025.09 - 2026.01 ✅已上线',
     '核心开发者（全仓 261 次提交，第二）',
     'Vue3 + TypeScript + AI 集成',
     '临床思维平台全面升级改版',
     ['界面翻新：登录页、首页、病例列表（整体 / 专项 / 接诊 / 医患沟通）、创建考试与练习、学生端作答界面',
      'AI 功能：AI 语音问诊、AI 伴学、AI 评价、AI 小助手随时问答、AI 快速生成病例',
      '新功能：扫码登录、病例二维码分享、开放病例直接练习、开放病例成绩查询']),
    ('7. 3D 医疗模拟 / 数字孪生 | 2023.06 - 2024.06 ✅已交付',
     '头号开发者（全仓 170 次提交，第一）',
     'Vue3 + Three.js + Electron + Unity WebGL',
     '桌面端 3D 医疗模拟训练系统，模拟多环境下不同类型伤员的救治全流程；支持 60 种伤情自由组合与学习 / 考试双模式',
     ['完成 677 个 glTF 模型（约 2.2GB）加载与骨骼动画、三维射线拾取点位交互',
      '打通 Electron 与 Unity WebGL 双向通信，实现 3D 场景与 Vue 业务层交互链路']),
    ('8. 棋牌项目矩阵 | 2019.02 - 2020.05',
     '核心开发',
     'Cocos2dx + Lua',
     '覆盖 8 省的棋牌游戏矩阵，包含麻将、斗地主、跑得快等',
     ['2019 年流水 6000 万，月流水稳定 500 万+']),
    ('9. 真人试衣系统 | 2010.06 - 2011.11',
     '开发者',
     'Flash + FLARToolKit + 人脸识别',
     '基于摄像头的 AR 试衣应用，支持 PC / MAC',
     ['已获 1 项发明专利、6 项实用新型专利、1 项外观设计专利',
      '已获 1 项软件著作权']),
]

for proj_title, role, tech, desc, features in projects:
    p = doc.add_paragraph()
    p.add_run(proj_title).bold = True
    p = doc.add_paragraph()
    p.add_run('角色：').bold = True
    p.add_run(role)
    p = doc.add_paragraph()
    p.add_run('技术栈：').bold = True
    p.add_run(tech)
    p = doc.add_paragraph()
    p.add_run('项目描述：').bold = True
    p.add_run(desc)
    for f in features:
        doc.add_paragraph(f'• {f}')
    doc.add_paragraph()

# ==================== 教育背景 ====================
doc.add_heading('教育背景', level=1)
p = doc.add_paragraph()
p.add_run('民办天狮职业技术学院').bold = True
p.add_run(' | 大专 | 工商管理 | 2003 - 2006')

# ==================== 保存 ====================
doc.save(OUT_PATH)
print(f'Word 文档已生成：{OUT_PATH}')