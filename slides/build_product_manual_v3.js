const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        Header, Footer, AlignmentType, LevelFormat,
        HeadingLevel, BorderStyle, WidthType, ShadingType,
        PageNumber, PageBreak, TabStopType, TabStopPosition } = require('docx');
const fs = require('fs');

const C = { gov: '2563EB', mgmt: '059669', ent: 'D97706', pub: '7C3AED',
  acc: 'F59E0B', text: '1E293B', gray: '64748B', line: 'E2E8F0' };
const border = { style: BorderStyle.SINGLE, size: 1, color: C.line };

function bodyT(text, opts) {
  return new Paragraph({ spacing: { before: 60, after: 60 },
    children: [new TextRun({ text, size: 20, font: 'Microsoft YaHei', color: C.text, ...(opts||{}) })] });
}
function secH(text, color) {
  return new Paragraph({ spacing: { before: 360, after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: color, space: 4 } },
    children: [new TextRun({ text, bold: true, size: 28, font: 'Microsoft YaHei', color })] });
}
function featItem(title, color) {
  return new Paragraph({ spacing: { before: 80, after: 0 },
    indent: { left: 360 },
    children: [new TextRun({ text: '\u25b8 ' + title, bold: true, size: 22, font: 'Microsoft YaHei', color })] });
}
function featDesc(desc) {
  return new Paragraph({ spacing: { before: 0, after: 80 },
    indent: { left: 720 },
    children: [new TextRun({ text: desc, size: 20, font: 'Microsoft YaHei', color: C.text })] });
}

const coverSection = {
  properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
  children: [
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 4800, after: 0 },
      children: [new TextRun({ text: '青川县庄子上工业园区', bold: true, size: 44, font: 'Microsoft YaHei', color: C.gov })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200, after: 0 },
      children: [new TextRun({ text: '智慧园区平台', bold: true, size: 52, font: 'Microsoft YaHei', color: C.text })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 400, after: 0 },
      children: [new TextRun({ text: '产品手册', bold: true, size: 36, font: 'Microsoft YaHei', color: C.gov })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 800, after: 0 },
      children: [new TextRun({ text: 'V2.0', size: 24, font: 'Microsoft YaHei', color: C.gray })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 1200, after: 0 },
      children: [new TextRun({ text: '青川县庄子上工业园区管委会', size: 22, font: 'Microsoft YaHei', color: C.gray })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80, after: 0 },
      children: [new TextRun({ text: new Date().getFullYear() + '年', size: 22, font: 'Microsoft YaHei', color: C.gray })] }),
  ]
};


const bodySection = {
  properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1080, bottom: 1440, left: 1080 } } },
  headers: { default: new Header({ children: [new Paragraph({
    alignment: AlignmentType.RIGHT, border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.gov, space: 4 } },
    children: [new TextRun({ text: '青川县庄子上智慧平台产品手册', size: 16, font: 'Microsoft YaHei', color: C.gray })]
  })] }) },
  footers: { default: new Footer({ children: [new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [
      new TextRun({ text: '第 ', size: 16, font: 'Microsoft YaHei', color: C.gray }),
      new TextRun({ children: [PageNumber.CURRENT], size: 16, font: 'Microsoft YaHei', color: C.gray }),
      new TextRun({ text: ' 页', size: 16, font: 'Microsoft YaHei', color: C.gray }),
    ]
  })] }) },
  children: [
    // 目录页
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 1200, after: 400 },
      children: [new TextRun({ text: '目 录', bold: true, size: 36, font: 'Microsoft YaHei', color: C.text })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 600 },
      children: [new TextRun({ text: 'TABLE OF CONTENTS', size: 18, font: 'Microsoft YaHei', color: C.gray })] }),
    new Paragraph({ spacing: { before: 80, after: 80 }, indent: { left: 360 },
      children: [new TextRun({ text: '一、产品概述', size: 20, font: 'Microsoft YaHei', color: '1E293B' })] }),
    new Paragraph({ spacing: { before: 80, after: 80 }, indent: { left: 360 },
      children: [new TextRun({ text: '二、政府领导层视角', size: 20, font: 'Microsoft YaHei', color: '2563EB' })] }),
    new Paragraph({ spacing: { before: 80, after: 80 }, indent: { left: 720 },
      children: [new TextRun({ text: 'AI招商服务', size: 20, font: 'Microsoft YaHei', color: '2563EB' })] }),
    new Paragraph({ spacing: { before: 80, after: 80 }, indent: { left: 720 },
      children: [new TextRun({ text: '全域安全监控', size: 20, font: 'Microsoft YaHei', color: '2563EB' })] }),
    new Paragraph({ spacing: { before: 80, after: 80 }, indent: { left: 720 },
      children: [new TextRun({ text: '告警与监视统计', size: 20, font: 'Microsoft YaHei', color: '2563EB' })] }),
    new Paragraph({ spacing: { before: 80, after: 80 }, indent: { left: 360 },
      children: [new TextRun({ text: '三、园区管理层视角', size: 20, font: 'Microsoft YaHei', color: '059669' })] }),
    new Paragraph({ spacing: { before: 80, after: 80 }, indent: { left: 720 },
      children: [new TextRun({ text: '智慧运营与监控', size: 20, font: 'Microsoft YaHei', color: '059669' })] }),
    new Paragraph({ spacing: { before: 80, after: 80 }, indent: { left: 720 },
      children: [new TextRun({ text: '能效监测与亩均论英雄', size: 20, font: 'Microsoft YaHei', color: '059669' })] }),
    new Paragraph({ spacing: { before: 80, after: 80 }, indent: { left: 720 },
      children: [new TextRun({ text: '运营管理与数字化办公', size: 20, font: 'Microsoft YaHei', color: '059669' })] }),
    new Paragraph({ spacing: { before: 80, after: 80 }, indent: { left: 360 },
      children: [new TextRun({ text: '四、入驻企业视角', size: 20, font: 'Microsoft YaHei', color: 'D97706' })] }),
    new Paragraph({ spacing: { before: 80, after: 80 }, indent: { left: 720 },
      children: [new TextRun({ text: '企业服务中心', size: 20, font: 'Microsoft YaHei', color: 'D97706' })] }),
    new Paragraph({ spacing: { before: 80, after: 80 }, indent: { left: 720 },
      children: [new TextRun({ text: '供需对接与商业服务', size: 20, font: 'Microsoft YaHei', color: 'D97706' })] }),
    new Paragraph({ spacing: { before: 80, after: 80 }, indent: { left: 720 },
      children: [new TextRun({ text: '场馆、活动与餐饮', size: 20, font: 'Microsoft YaHei', color: 'D97706' })] }),
    new Paragraph({ spacing: { before: 80, after: 80 }, indent: { left: 360 },
      children: [new TextRun({ text: '五、公众用户视角', size: 20, font: 'Microsoft YaHei', color: '7C3AED' })] }),

    new Paragraph({ children: [new PageBreak()] }),

    
    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: '一、产品概述', bold: true, font: 'Microsoft YaHei' })] }),
    bodyT('青川县庄子上工业园区智慧平台是一个集数据采集、智能分析、可视化展示、业务管理于一体的综合性智慧园区解决方案。平台为园区政府、管理方、入驻企业和公众用户四类角色提供定制化的功能与界面，涵盖89+功能模块，对接12大第三方数据源。'),
    bodyT('平台采用五层架构设计：C端展示层(大屏数据驾驶舱、园区官网、H5移动端、微信小程序)、应用服务层(AI招商、安全监控、企业服务、运营管理、能效监测)、数据服务层(数据湖、分析引擎、AI算法、GIS平台)、设施层(视频监控、传感器、智能计量表)和数据对接层(天然气、电力、水、综合执法、环保、应急、交通、公安、自然资源、气象、人社、运营商)。'),
    bodyT('园区规划面积2.1km²，已入驻47家企业，从业人数3,000+人。平台覆盖从招商引资到企业服务、从安全监控到能效管理的全生命周期管理。'),
    new Paragraph({ children: [new PageBreak()] }),

    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: '二、政府领导层视角', color: C.gov, bold: true, font: 'Microsoft YaHei' })] }),
    bodyT('为园区政府管理部门提供数据驱动的决策支持，涵盖招商引资、产业规划、安全监管、环保监控等核心能力，为园区发展提供数据基础与决策依据。本视角聚焦宏观决策与全局监控，支持数据可视化大屏展示。'),

    secH('AI招商服务', C.gov),
    bodyT('运用大数据绘制区域产业图谱，精准识别主导产业的「强链、补链、延链、扩链」环节。基于深度学习模型，从海量企业数据中智能挖掘与园区产业定位匹配的目标企业，并生成360度企业画像报告。'),

    featItem('AI智能招商', C.gov),
    featDesc('将园区资源禀赋、政策优势与目标企业需求进行AI智能匹配，推荐成功率最高的招商目标。支持数据采集、需求分析、智能算法、结果输出全流程管理。操作按钮：【智能匹配】【导出报告】【查看画像】【筛选条件】。'),
    featItem('数字沙盘', C.gov),
    featDesc('基于实景三维建模，打造园区的三维数字沙盘，立体展示区位交通、产业布局、地块信息、配套规划，为招商引资提供沉浸式展示体验。操作按钮：【三维漫游】【地块查询】【配套查看】【一键分享】。支持VR模式切换。'),
    featItem('产业链图谱', C.gov),
    featDesc('运用大数据绘制区域产业图谱，精准识别农副产品加工、新能源等主导产业的强链、补链、延链环节，可可视化展示产业链上下游关系。操作按钮：【图谱查看】【节点展开】【产业分析】【导出图谱】。'),
    featItem('目标企业挖掘', C.gov),
    featDesc('基于深度学习模型，从海量企业数据中智能挖掘与园区产业定位匹配、有投资扩张意向的优质企业。提供TOP5智能推荐，包含企业资本、行业、扩张意向等多维度信息。操作按钮：【智能挖掘】【筛选条件】【查看详情】【导出列表】。'),
    featItem('招商项目管理', C.gov),
    featDesc('实现从项目线索、洽谈跟进、协议签订到落地建设的数字化、流程化管理，支持全生命周期跟踪。操作按钮：【新增项目】【进度更新】【文档管理】【统计分析】。支持甘特图展示项目进度。'),
    featItem('招商入口', C.gov),
    featDesc('为园区提供线上宣传平台，展示园区优势、招商政策、入驻企业等信息，引流潜在投资者。操作按钮：【预览】【编辑内容】【发布】【查看访问量】。'),
    featItem('云展馆/电子名片', C.gov),
    featDesc('为企业提供线上宣传平台，展示产品与技术，替代传统纸质材料，实现招商信息「随时可看、一键分享」。支持生成二维码、创建企业名片、访问量统计。操作按钮：【创建名片】【生成二维码】【分享】【访问统计】。'),

    secH('全域安全监控', C.gov),
    bodyT('构建「天眼+哨兵+慧眼」三位一体的全域安全监控体系，实现对道路交通安全、地质灾害、环境保护的全天候实时监控与智能预警。'),

    featItem('天眼 - 道路交通安全监控', C.gov),
    featDesc('通过视频监控设备与AI视频分析技术，实现车辆识别、违章捕捉、交通状态实时监控。支持非名录车闯入、驾驶违规等异常事件的自动识别与告警。操作按钮：【实时视频】【违章查询】【车辆追踪】【告警设置】。支持16宫格视频监控布局。'),
    featItem('哨兵 - 地灾安全监测', C.gov),
    featDesc('通过部署在地质灾害易发区域的多种传感器，实现对滑坡、泥石流、地表移动等灾害的实时监测与提前预警，确保园区地质安全。操作按钮：【实时数据】【预警设置】【历史查询】【导出报告】。'),
    featItem('慧眼 - 环保监控检测', C.gov),
    featDesc('对园区环境指标实时监测，包括扬尘、噪声、水质、排放等数据的追踪与分析，支持环保合规管理与绿色园区创建。操作按钮：【实时监测】【超标告警】【数据导出】【报表生成】。'),
    featItem('视频监控总览', C.gov),
    featDesc('统一视频监控平台，支持16宫格智能排列，覆盖园区关键区域，实现全园区视频监控的统一管控与快速切换。操作按钮：【16宫格视图】【单屏放大】【轮播切换】【设备状态查看】。'),
    featItem('应急指挥联动', C.gov),
    featDesc('提供一键调度、预案管理、实时通讯等功能，确保异常事件发生时能够快速响应、精准处置，响应效率提升60%。操作按钮：【一键调度】【预案启动】【通讯录】【事件记录】。'),

    secH('告警与监视统计', C.gov),
    featItem('告警中心', C.gov),
    featDesc('统一告警管理中心，支持高/中/低风险告警分级。自动识别非名录车闯入、驾驶违规、设备离线、库容应急等多种异常事件，实现告警处置全流程管理。操作按钮：【告警筛选】【处置分配】【闭环跟踪】【统计分析】。支持按级别、类型、状态筛选。'),
    featItem('审计一本账', C.gov),
    featDesc('提供全域监管视图，实时监控排放总量、运输车次、违规事件、库容利用率等数据。支持生成审计报告，提供排放监控、运输监控、场站营运等多维度统计与导出。操作按钮：【数据导出】【报告生成】【多维度筛选】【图表查看】。'),

    new Paragraph({ children: [new PageBreak()] }),

    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: '三、园区管理层视角', color: C.mgmt, bold: true, font: 'Microsoft YaHei' })] }),
    bodyT('为园区日常运营管理人员提供全面的智慧运营工具，涵盖安全监控、能效管控、设备管理、数据分析、日常办公等日常运维能力。本视角聚焦运营效率提升与精细化管控。'),

    secH('智慧运营与监控', C.mgmt),
    featItem('全域安全监控(管理视角)', C.mgmt),
    featDesc('统一管理天眼、哨兵、慧眼三大监控体系。支持实时视频查看、历史回放、车辆追踪、边界告警、环境指标监测、应急指挥调度等功能。操作按钮：【实时监控】【历史回放】【告警处理】【调度指挥】。'),
    featItem('视频监控总览(管理视角)', C.mgmt),
    featDesc('支持16宫格智能排列视图，可快速切换园区各关键点位视频流。支持设备在线状态监控、视频流加载状态显示。操作按钮：【16宫格】【单屏全屏】【设备状态】【视频录制】。'),
    featItem('告警中心(管理视角)', C.mgmt),
    featDesc('可可视化展示各类告警事件，支持按级别、类型、状态筛选。提供告警处置跟踪、处置率统计、告警趋势分析等功能。操作按钮：【筛选查询】【处置分配】【趋势图表】【导出数据】。'),
    featItem('应急指挥联动(管理视角)', C.mgmt),
    featDesc('提供一键调度、预案管理、实时通讯等功能，确保异常事件发生时能够快速响应、精准处置。操作按钮：【应急预案】【调度指令】【事件记录】【效能评估】。'),
    featItem('设备管理', C.mgmt),
    featDesc('统一管理园区所有智能设备，包括摄像头、传感器、GPS设备、雷达、称重设备等。支持设备在线状态监控、远程维护、异常告警。操作按钮：【设备列表】【状态监控】【远程配置】【告警设置】。'),

    secH('能效监测与亩均论英雄', C.mgmt),
    featItem('企业用能监测', C.mgmt),
    featDesc('通过智能计量表实时采集能耗数据，支持水、电、气多维度监控。可可视化展示能耗趋势，支持异常告警与能耗分析。操作按钮：【实时数据】【趋势图表】【异常告警】【数据导出】。'),
    featItem('亩均论英雄', C.mgmt),
    featDesc('大数据评价体系，通过投入产出、能耗排放等多维度指标对企业综合评价，实现资源要素差异化配置，倒逼企业提质增效。操作按钮：【企业排名】【评价详情】【资源配置】【报告导出】。支持按年度、产业筛选。'),
    featItem('能效分析', C.mgmt),
    featDesc('企业能耗效率多维度分析，提供同比环比趋势可视化。帮助企业识别能耗浪费点，优化生产过程能耗。操作按钮：【能效对比】【趋势分析】【浪费点识别】【优化建议】。'),
    featItem('异常预警', C.mgmt),
    featDesc('基于历史数据与算法模型，实时监测用能异常并自动触发预警，支持用电突增、用水异常、气压偏低、设备效率下降等多种预警类型。操作按钮：【预警列表】【原因分析】【处理跟踪】【预警设置】。'),

    secH('运营管理与数字化办公', C.mgmt),
    featItem('运营服务中心', C.mgmt),
    featDesc('大屏指挥+园小二团队+长效服务模式，提供园区运营的统一服务窗口与指挥调度能力。操作按钮：【大屏视图】【团队管理】【服务记录】【调度指令】。'),
    featItem('数字化办公', C.mgmt),
    featDesc('包括审批中心(审批发起/待办/已办/流程跟踪)、日常巡检(巡检计划/记录/问题上报)、资产管理(资产登记/领用/归还/盘点)、合同管理(合同起草/审批/归档/到期提醒)、档案管理(档案分类/借阅/归还)、会议管理(会议室预定/会议通知/纪要)、用车申请、加班申请、日常报销、入职/离职管理等全流程在线化办公模块。操作按钮：【发起审批】【待办事项】【已办查询】【流程跟踪】。支持自定义审批流程。'),
    featItem('数据中心', C.mgmt),
    featDesc('汇聚园区全量数据，构建统一数据湖。提供数据质量监控(完整性/准确性/及时性评估)、自定义报表(拖拽式报表设计器)、多维度分析(钻取/联动/环比)、数据导出等能力。支持数据备份与灾难恢复，确保数据安全。操作按钮：【数据查询】【报表生成】【数据导出】【备份恢复】。'),
    featItem('系统管理', C.mgmt),
    featDesc('包括用户管理(账号创建/禁用/信息修改)、角色权限(RBAC模型/菜单级/按钮级权限)、组织架构(部门/岗位/职级)、系统参数(全局配置/字典管理)、日志审计(操作日志/登录日志/异常日志)、消息通知配置(短信/邮件/站内信)等基础管理功能。操作按钮：【用户管理】【角色配置】【权限分配】【日志查询】。'),

    new Paragraph({ children: [new PageBreak()] }),

    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: '四、入驻企业视角', color: C.ent, bold: true, font: 'Microsoft YaHei' })] }),
    bodyT('为园区入驻企业提供全流程数字化服务，涵盖诉求处理、缴费服务、供需对接、场馆餐饮、用工信息等多个维度。本视角聚焦企业便利性与服务质量提升。'),

    secH('企业服务中心', C.ent),
    featItem('诉求直达(随手拍)', C.ent),
    featDesc('支持文字+图片+定位快速提交诉求，涵盖设施报修、安全隐患、环境环卫、排水堵塞、照明故障、绿化养护、噪音投诉等多种诉求类型。系统2小时内响应处理，支持诉求数据导出为Excel/PDF/Word格式。操作按钮：【提交诉求】【查看进度】【历史诉求】【导出数据】。支持图片上传和定位自动获取。'),
    featItem('服务直达', C.ent),
    featDesc('一站式服务中心，支持天然气费、自来水费、电费、通信费等在线查询与缴费。提供用户信息管理、历史记录查询、支付状态统计等功能。操作按钮：【账单查询】【在线缴费】【历史记录】【发票管理】。'),
    featItem('园企交互', C.ent),
    featDesc('提供在线沟通(即时消息/文件传输)、通告公告(园区通知/政策发布)、政策推送(智能匹配/主动推送)、双向互动沟通(留言反馈/满意度评价)等功能，建立园区与企业之间的高效沟通桥梁。操作按钮：【发送消息】【查看公告】【政策查看】【互动回复】。'),
    featItem('企业入驻', C.ent),
    featDesc('支持线上申请(提交入驻意向/上传资质材料)、证照管理(数字证照/证照到期提醒)、入驻流程全线上化，包括数字证照、电子签名等。操作按钮：【入驻申请】【证照上传】【进度查询】【电子签约】。'),
    featItem('信用评价', C.ent),
    featDesc('企业信用多维度评价体系(经营合规/合同履约/环保达标/安全生产/纳税信用)，帮助企业建立良好信用记录，助力融资与合作。操作按钮：【信用查询】【评价详情】【信用报告】【修复申请】。'),

    secH('供需对接与商业服务', C.ent),
    featItem('物流供需', C.ent),
    featDesc('园区货源+返程空车+车辆资源智能匹配，每张卡片带联系方式，支持多条件筛选(货物类型/吨位/出发地/目的地)。包括生产供应商、运输供应商管理。操作按钮：【发布货源】【查找车辆】【智能匹配】【联系供应商】。'),
    featItem('特色供需', C.ent),
    featDesc('青川山珍(木耳/香菇/竹荪/核桃/蜂蜜/茶叶/腊肉/天麻/杜仲)等9种特色产品展示与供需对接，直接联系供应商，支持产品分类浏览。操作按钮：【产品浏览】【发布供需】【联系供应商】【订单管理】。'),
    featItem('政策推送', C.ent),
    featDesc('惠企政策精准匹配与主动推送，涵盖补贴(技术改造/研发/就业)、减免(税收/租金/能耗)、人才(引进/培训/住房)等多类政策，支持政策查询与筛选。操作按钮：【政策匹配】【查看详情】【申报入口】【历史政策】。'),
    featItem('数字化活动', C.ent),
    featDesc('支持招标(发布标书/投标准入/评标)、竞价(在线出价/实时排名/成交确认)、审批全流程线上化，提供数字化活动管理平台。操作按钮：【活动发布】【在线参与】【结果公示】【记录查询】。'),
    featItem('客户管理', C.ent),
    featDesc('客户关系维护(客户档案/联系人/跟进记录)、商机追踪(商机阶段/预计金额/成交概率)、合同执行(合同台账/履行进度/收款管理)，包括合同管理、客户资料管理等功能。操作按钮：【客户列表】【商机跟进】【合同管理】【数据分析】。'),

    secH('场馆、活动与餐饮', C.ent),
    featItem('场馆服务', C.ent),
    featDesc('提供多功能会议室、展览厅、体育馆、培训教室、会议室、接待厅、多媒体厅、园区广场等8大场馆的在线预约、查看开放时间、状态管理等功能。排序规则为待审批>被取消>已通过。操作按钮：【场馆浏览】【在线预约】【查看预约】【取消预约】。支持3D波纹动态卡片展示。'),
    featItem('园区活动', C.ent),
    featDesc('支持文体活动(运动会/文艺汇演)、培训讲座(技能培训/企业管理/安全生产)、公益活动(环保/扶贫/志愿)、商务交流(企业沙龙/行业论坛/项目路演)、庆典活动、展览展示等多类活动的参与、活动状态跟踪。排序规则为待审批>被取消>已通过。操作按钮：【活动浏览】【在线报名】【我的活动】【活动评价】。'),
    featItem('园区餐饮', C.ent),
    featDesc('流动卡片展示6家餐厅，20+道菜品一览。支持分类筛选(中餐/快餐/面食/饮品甜点)、今日菜谱查看、菜品价格与位置信息。操作按钮：【餐厅浏览】【今日菜谱】【菜品搜索】【评价查看】。'),
    featItem('用工信息', C.ent),
    featDesc('招聘信息发布+岗位智能匹配，支持分类筛选(行业/岗位/薪资/学历)，帮助企业快速解决用工问题。包括招聘信息发布、简历投递、岗位匹配等。操作按钮：【浏览岗位】【投递简历】【发布招聘】【匹配推荐】。'),
    featItem('停车服务', C.ent),
    featDesc('高德地图导航+实时余位查看+在线缴费，全园区停车位统一管理，支持多车牌绑定与管理。操作按钮：【地图导航】【余位查询】【在线缴费】【车牌管理】。'),
    featItem('访客录入', C.ent),
    featDesc('预约(填写访客信息/选择拜访对象/预约时间)、登记(扫码签到/身份验证)、拜访记录全线上化，提供车牌管理、门禁管理、访客趋势统计(ECharts双曲线趋势图：人员/车辆)等功能。操作按钮：【访客预约】【扫码登记】【访客记录】【统计分析】。'),

    new Paragraph({ children: [new PageBreak()] }),

    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: '五、公众用户视角', color: C.pub, bold: true, font: 'Microsoft YaHei' })] }),
    bodyT('为园区公众提供便民信息服务，包括园区公告、活动信息、商业服务、用工信息等。本视角聚焦信息获取便捷性与公众参与体验。'),

    secH('信息服务与园区参与', C.pub),
    featItem('园区公告', C.pub),
    featDesc('实时查看园区通告，包括施工(道路施工/管网改造)、检查(安全检查/设备检修)、维护(绿化养护/设施维护)等各类公告信息。操作按钮：【公告列表】【详情查看】【搜索】【分类筛选】。'),
    featItem('园区活动', C.pub),
    featDesc('查看并参与文体活动、培训讲座、公益活动、商务交流等园区活动，跟踪活动状态(报名中/进行中/已结束)。操作按钮：【活动列表】【在线报名】【我的报名】【活动评价】。'),
    featItem('园区政策', C.pub),
    featDesc('查看公开政策信息，了解园区发展动态、优惠政策、服务指南等。操作按钮：【政策列表】【详情查看】【搜索】【收藏】。'),
    featItem('园区餐饮', C.pub),
    featDesc('6家餐厅信息+今日菜谱+价格信息+分类筛选，包括全部菜品、饮品甜点等分类浏览。操作按钮：【餐厅查看】【菜谱浏览】【价格查看】【分类筛选】。'),
    featItem('停车服务', C.pub),
    featDesc('地图查看停车位实时余位，在线缴费，支持多车牌管理。操作按钮：【地图查看】【余位查询】【在线缴费】【车牌绑定】。'),
    featItem('场馆服务', C.pub),
    featDesc('浏览园区8大场馆信息(多功能会议厅/展览厅/体育馆/培训教室/会议室/接待厅/多媒体厅/园区广场)，了解开放时间与在线预约。操作按钮：【场馆浏览】【开放时间】【在线预约】【我的预约】。'),
    featItem('特色供需', C.pub),
    featDesc('浏览青川山珍(木耳/香菇/竹荪/核桃/蜂蜜/茶叶/腊肉/天麻/杜仲)等9种特色产品，直接联系供应商。操作按钮：【产品浏览】【供应商联系】【分类查看】【搜索】。'),
    featItem('用工信息', C.pub),
    featDesc('查看园区企业招聘信息，岗位智能匹配与筛选。操作按钮：【岗位浏览】【投递简历】【岗位筛选】【企业查看】。'),
    featItem('微信小程序', C.pub),
    featDesc('快速接入智慧平台，提供便捷的移动端服务入口，支持扫码即用、功能菜单导航、个人中心管理、消息通知推送。操作按钮：【扫码进入】【功能菜单】【个人中心】【消息通知】。'),
    featItem('园区宣传', C.pub),
    featDesc('园区概况(历史沿革/区位优势)、发展历程(阶梯式时间轴展示)、优势产业(农副产品加工/新能源/新材料)、合作案例多维度展示。操作按钮：【园区介绍】【发展历程】【优势产业】【成功案例】。'),

  ]
};


const doc = new Document({
  styles: { default: { document: { run: { font: 'Microsoft YaHei', size: 20 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 32, bold: true, font: 'Microsoft YaHei', color: C.text },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 } },
    ] },
  sections: [coverSection, bodySection],
});

const out = __dirname + '/output/product_manual_v3.docx';
Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(out, buffer);
  console.log('V3 OK: ' + buffer.length + ' bytes');
}).catch(err => {
  console.error('Error:', err.message);
});
