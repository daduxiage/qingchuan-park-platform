# 长期记忆

## 青川县庄子上工业园区智慧园区平台

### 项目信息
- **工作空间**: 20260427142154
- **平台名称**: 青川县庄子上工业园区智慧园区平台
- **测试账号**: admin / admin123
- **平台Logo**: 🏭

### 功能模块（按设计方案分类）

#### 核心业务模块
1. **AI招商服务** - AI智能招商、数字沙盘、企业档案、产业链图谱、目标企业挖掘、招商项目管理、招商门户、云展馆/电子名片
2. **全域安全监控** - 道路交通安全监控(天眼)、地灾安全监测(哨兵)、环保监控检测(慧眼)、视频监控总览、应急指挥联动
3. **企业服务** - 诉求直达(随手拍)、服务直达、园企交互、供需对接、企业入驻、证照管理、信用评价
4. **运营管理** - 物流供需、特色供需、政策推送、数字化活动、餐饮服务、园区宣传
5. **能效与运营监测** - 企业用能监测(计量)、亩均论英雄、能效分析、异常预警

#### 基础功能模块
- 管理支撑、数字化办公、基础支撑、数据中心、数据接入与基础管理

### 技术规范
- 样式风格: 深蓝科技风(#101838, #162969)，主色#40c4ff
- 字体: Microsoft YaHei/微软雅黑
- 布局: 全屏式首页，无左侧导航栏
- 主题切换: 支持深色/浅色模式

## 3D 建筑绘制方案

### Three.js + AMap.CustomLayer（industrial_economy.html）
- 7200 栋建筑，120×60 网格，BoxGeometry + MeshPhongMaterial
- 手动墨卡托转换公式，matrixAutoUpdate=false
- Three.js r0.152.2，CDN: cdn.jsdelivr.net

### Loca.PolygonLayer + GeoJSON（deyang_mining_amap_dashboard.html）
- Polygon 特征数据 + extrusionHeight
- Loca v2.0.0 动态加载: webapi.amap.com/loca?v=2.0.0
- 支持函数式样式（每栋建筑独立颜色/高度）
- 新文件 qingchuan_jiaxing_buildings.html 采用此方案

## park_*.html 园区子页面架构

### 页面列表（7页）
| 页面 | 标题 | 源文件 | 说明 |
|------|------|--------|------|
| park_sandbox.html | 数字化沙盘 | digital_sand.html | 3D数字沙盘 |
| park_economy.html | 园区经济概览 | park_overview.html | 地图+悬浮图表 |
| park_land.html | 用地规划 | land_planning.html | 用地规划 |
| park_enterprise.html | 企业监测 | park_overview.html | 企业列表(entPanel) |
| park_service.html | 园区服务 | park_overview.html | 服务仪表盘(servicePanel) |
| park_energy.html | 用能监测 | park_energy_source.html | ECharts地图+能耗图表 |
| park_safety.html | 全域安全监控 | safety_monitor.html | 安全监控 |

### 统一设计规范
- **标题栏**: screen-container > title-box（时间左 + 标题中 + 刷新/主页右）
- **底部导航**: 圆形 bottom-tabs（80px圆形按钮，虚线边框旋转动画，7个tab-btn）
- **导航方式**: `<a>` 链接跳转（非 in-page JS panel 切换）
- **样式**: 深蓝科技风，--primary-color: #40c4ff
- **禁止注入**: 不得使用 park-hdr/park-tabs/ph-bar/ph-wrap 等注入式代码
