/**
 * park_charts.js — 园区平台统一图表样式
 * 基于 Apache ECharts，所有可视化页面共享调用
 *
 * ============================================================
 *  图表 #1：趋势曲线图 trendLine — 样式标准
 * ============================================================
 *
 * 【生成方法】
 *   ParkCharts.trendLine(domId, opts)
 *     domId  - 容器 DOM id（字符串）
 *     opts   - { data:number[], name:string, color?:hex, unit?:string }
 *     返回值  - ECharts 实例（可调用 .resize() 等）
 *
 * 【视觉规范】
 *   主题色    T.primary  = #40c4ff （默认曲线色，可 opts.color 覆盖）
 *   坐标轴色  T.axis     = #3a5a80
 *   轴文字色  T.text     = #8ab4d8
 *   高亮文字  T.textBright = #cfe4ff
 *   背景色    T.bg       = rgba(6,14,40,.95)
 *
 * 【曲线样式】
 *   type: line, smooth: true          — 平滑曲线
 *   symbol: circle, symbolSize: 8     — 圆形数据点，直径 8px
 *   lineStyle: { width: 2 }           — 线宽 2px
 *   areaStyle: {}                     — 默认半透明面积填充
 *
 * 【数据标签】
 *   label.position: 'top'             — 标签置于数据点上方
 *   label.color: T.textBright         — 高亮文字色
 *   label.fontSize: 9                 — 字号 9px
 *   label.distance: 6                 — 标签距数据点 6px，避免与相邻点标签重叠
 *   label.formatter: '{c}'           — 仅显示数值，单位由 Y 轴/tooltip 体现
 *
 * 【Grid 布局 — 自适应间距】
 *   grid.top: 38                      — 上边距 38px，为 position:'top' 标签+纵向间距留空间
 *                                      （28→38 解决相邻标签重叠，2026-05-17）
 *   grid.bottom: 24                   — 下边距 24px，容纳底部图例
 *   grid.left: 4 / right: 32          — 左右边距，right:32 防止右侧标签 & "12月" 被裁切
 *   containLabel: true                — 自适应补偿标签 & 轴文字占位
 *   ★ 不改变图表容器宽高占比，仅调整 ECharts 内部绘图区边距
 *
 * 【Grid 间距设计原则】
 *   所有 grid 间距调整遵循同一原则：
 *     ◆ 不动图表容器 width/height 占比
 *     ◆ 只扩展 ECharts 内部 grid.{top,right,bottom,left} 边距
 *     ◆ 配合 containLabel:true 实现自适应补偿
 *
 *   修订记录（2026-05-17）：
 *     #1 顶点标签被裁切 → grid.top: 10→28
 *     #2 右侧"12月"被裁切 → grid.right: 6→16
 *     #3 相邻点标签重叠   → grid.top: 28→38 + label.distance:6
 *     #4 右侧标签仍被裁切 → grid.right: 16→32
 *     #5 标签单位冗余     → label.formatter: '{c}单位' → '{c}'
 *     #6 补入场动画       → + animationDuration/Easing/Delay（与柱状图对齐）
 *
 * 【网格线】
 *   yAxis.splitLine.color: T.text     — 与轴文字同色（#8ab4d8）
 *   opacity: .15                      — 15% 透明度，既可见又不抢眼
 *
 * 【图例】
 *   legend.bottom: 0                  — 底部贴边
 *   legend.icon: 'roundRect'          — 统一圆角矩形图标（16×6）
 *   legend.textStyle.color: T.text    — 与轴文字同色
 *   legend.textStyle.fontSize: 9      — 字号 9px
 *
 * 【交互】
 *   tooltip.trigger: 'axis'          — 轴线触发
 *   tooltip.axisPointer: 'cross'     — 十字准星
 *   animationDuration: 1200           — 入场动画 1.2s（与柱状图一致）
 *   animationEasing: 'elasticOut'     — 弹性缓出，逐条延迟 100ms
 *   emphasis.focus: 'series'         — hover 高亮整条线
 *   emphasis.scale: 1.5              — 数据点放大 1.5 倍
 *   emphasis.shadowBlur: 16          — 光晕阴影
 *
 * 【X 轴】
 *   type: category, boundaryGap: false
 *   data: ['1月'..'12月'] （默认月份，可按需覆盖）
 *
 *   ★ 与柱状图交互维度对齐项：入场动画（duration/easing/delay）一致
 *
 * ============================================================
 *  图表 #2：排名柱状图 barRank — 样式标准
 * ============================================================
 *
 * 【生成方法】
 *   ParkCharts.barRank(domId, opts)
 *     domId  - 容器 DOM id（字符串）
 *     opts   - { items:[{name:string,value:number}], name:string, unit:string, colors?:string[] }
 *     返回值  - ECharts 实例（可调用 .resize() 等）
 *
 * 【视觉规范】
 *   主题色    T.primary  = #40c4ff （默认柱色，可 opts.colors 覆盖）
 *   坐标轴色  T.axis     = #3a5a80
 *   轴文字色  T.text     = #8ab4d8
 *   高亮文字  T.textBright = #cfe4ff
 *   排名配色  ['#40c4ff','#4cd964','#ff9500','#7b61ff','#ff6b81',...] 9 色循环
 *
 * 【排序规则】
 *   内部自动 opts.items.sort(desc) + yAxis.inverse:true → 最高值在最上面
 *
 * 【柱体样式】
 *   type: bar, barWidth: 16           — 柱宽 16px
 *   barCategoryGap: '30%'             — 分类间距 30%
 *   borderRadius: [0,5,5,0]           — 右侧圆角 5px
 *   itemStyle.shadowBlur: 8           — 柱体阴影
 *   ★ 逐条独立颜色：data[i].itemStyle.color = rankColors[i%9]
 *   ★ series.color: rankColors[0]     — 图例 Icon 取色基准
 *
 * 【数据标签】
 *   label.position: 'right'           — 标签置于柱体右侧
 *   label.color: T.textBright         — 高亮文字色
 *   label.fontSize: 9                 — 字号 9px
 *   label.distance: 4                 — 标签距柱体 4px
 *   label.formatter: '{c}'            — 仅显示数值，单位由 X 轴体现
 *
 * 【Grid 布局 — 自适应间距】
 *   grid.left: 4                      — 左边距 4px
 *   grid.right: 45                    — 右边距 45px，为 position:'right' 标签留空间
 *   grid.top: 8 / bottom: 28          — 上下边距，bottom 容纳图例
 *   containLabel: true                — 自适应补偿标签 & 轴文字占位
 *   ★ 与曲线图同原则：不动容器宽高，只调 grid 边距
 *
 * 【网格线】
 *   xAxis.splitLine.color: T.text     — 与轴文字同色（#8ab4d8）
 *   opacity: .15                      — 15% 透明度
 *
 * 【坐标轴】
 *   xAxis.type: 'value'               — X 轴为数值轴（横向）
 *   yAxis.type: 'category'            — Y 轴为类目轴（企业名）
 *   yAxis.inverse: true               — 反转使最高值在顶部
 *   yAxis.axisLine.color: T.axis      — 轴线色 #3a5a80
 *   yAxis.axisLabel.color: 函数       — 逐条与柱体同色
 *   XY 轴 fontSize: 未显式设置        — 走 ECharts 默认，与曲线图一致
 *
 * 【图例】
 *   legend.bottom: 0                  — 底部贴边
 *   legend.icon: 'roundRect'          — 统一圆角矩形图标（16×6）
 *   legend.textStyle.color: T.text    — 与轴文字同色 #8ab4d8
 *   legend.textStyle.fontSize: 9      — 字号 9px
 *   ★ 与曲线图图例配置完全一致
 *
 * 【交互 & 动画】
 *   tooltip.trigger: 'axis'           — 轴线触发
 *   tooltip.axisPointer: 'shadow'     — 阴影指示器
 *   animationDuration: 1200           — 入场动画 1.2s
 *   animationEasing: 'elasticOut'     — 弹性缓出
 *   animationDelay: idx*100           — 逐条错开 100ms
 *   emphasis.scale: 1.08              — hover 放大
 *   emphasis.shadowBlur: 20           — hover 光晕
 *
 * ============================================================
 *  图表 #3：南丁格尔玫瑰图 pieRose — 样式标准
 * ============================================================
 *
 * 【生成方法】
 *   ParkCharts.pieRose(domId, opts)
 *     domId  - 容器 DOM id（字符串）
 *     opts   - { data:[{name:string,value:number}], name?:string, colors?:string[] }
 *     返回值  - ECharts 实例（可调用 .resize() 等）
 *
 * 【视觉规范】
 *   玫瑰配色  ['#40c4ff','#4cd964','#ff9500','#7b61ff','#ff6b81','#ffbb33',...] 9 色循环
 *   坐标轴色  T.axis     = #3a5a80
 *   轴文字色  T.text     = #8ab4d8
 *   高亮文字  T.textBright = #cfe4ff
 *   背景色    T.bg       = rgba(6,14,40,.95)
 *
 * 【扇区样式】
 *   type: pie, roseType: 'area'       — 南丁格尔玫瑰图（半径映射数值）
 *   radius: ['22%', '62%']            — 内径 22% 空心，外径 62% 为标签留空间
 *   center: ['50%', '48%']            — 居中偏上，均匀分配标题-饼图-图例间距
 *   borderRadius: 6                   — 扇区圆角 6px
 *   borderColor: T.bg                 — 扇区间隔色（与背景融合）
 *   borderWidth: 2                    — 扇区间隔 2px
 *   shadowBlur: 10                    — 扇区阴影
 *   ★ 逐条独立颜色：data[i].itemStyle.color = roseColors[i%9]
 *   ★ series.color: roseColors        — 图例 Icon 取色数组
 *
 * 【数据标签】
 *   label.show: true                  — 显示标签
 *   label.color: T.textBright         — 高亮文字色
 *   label.fontSize: 9                 — 字号 9px
 *   label.formatter: '{b}\n{d}%'      — 名称 + 百分比（两行）
 *   label.lineHeight: 14              — 行高 14px
 *   ★ 外径 62% 确保最外层标签不被容器裁切
 *
 * 【布局 — 自适应】
 *   center 48% → 标题-饼图-图例间距均匀（同曲线/柱状图间距原则）
 *   radius 外径 62% → 标签空间充足
 *
 * 【图例】
 *   legend.bottom: 0                  — 底部贴边
 *   legend.icon: 'roundRect'          — 统一圆角矩形图标（16×6）
 *   legend.textStyle.color: T.text    — 与轴文字同色 #8ab4d8
 *   legend.textStyle.fontSize: 9      — 字号 9px
 *   legend.data: 自动从 opts.data 提取 name
 *   ★ 与曲线图/柱状图图例配置完全一致
 *
 * 【交互 & 动画】
 *   tooltip: item 触发                — 显示 {名称}: {值} ({百分比}%)
 *   emphasis.scaleSize: 10            — hover 扇区外扩 10px
 *   emphasis.label.fontSize: 13       — hover 标签放大
 *
 * 【入场动画】
 *   animationType: 'scale'            — 扇区从中心缩放入场
 *   animationDuration: 1200           — 入场动画 1.2s（与曲线/柱状图一致）
 *   animationEasing: 'elasticOut'     — 弹性缓出
 *   animationDelay: idx × 80          — 逐条延迟 80ms
 *
 * 【自动旋转】
 *   驱动方式: requestAnimationFrame（60fps 平滑）
 *   旋转速度: 0.005°/ms ≈ 每圈 72 秒
 *   启动时机: setTimeout 延迟 2s → 等入场动画播完再启动
 *             （避免旋转帧的 animation:false 打断入场效果）
 *   标签同步: ECharts setOption 自动重算标签位置，天然跟随扇区旋转
 *   资源清理: 覆写 ch.dispose() → cancelAnimationFrame 停止旋转帧
 *
 *   修订记录（2026-05-17）：
 *     #1 基础饼图升级为南丁格尔玫瑰图
 *     #2 center 55%→48% 均匀标题-饼图-图例间距
 *     #3 补入场动画（scale + elasticOut + 逐条延迟）
 *     #4 加入自动旋转 + 标签同步跟随
 *     #5 旋转降速 0.015→0.005°/ms
 *     #6 setTimeout 2s 延迟避免打断入场动画
 *
 * ============================================================
 *  图表 #4：环形统计图 pieRing — 样式标准
 * ============================================================
 *
 *   ★ 与 #3 玫瑰图的核心区别：非玫瑰（无roseType）、中心显示数值总和
 *
 * 【生成方法】
 *   ParkCharts.pieRing(domId, opts)
 *     domId  - 容器 DOM id（字符串）
 *     opts   - { data:[{name:string,value:number}], name?:string, unit?:string, colors?:string[] }
 *     返回值  - ECharts 实例（可调用 .resize() 等）
 *
 * 【视觉规范】
 *   配色      ['#40c4ff','#4cd964','#ff9500','#7b61ff','#ff6b81','#ffbb33',...] 9 色循环
 *   轴文字色  T.text     = #8ab4d8
 *   高亮文字  T.textBright = #cfe4ff
 *   背景色    T.bg       = rgba(6,14,40,.95)
 *
 * 【扇区样式】
 *   type: pie（普通环形，非玫瑰）
 *   radius: ['42%', '62%']            — 内径42%为中心数字留空间，外径62%为标签留空间
 *   center: ['50%', '48%']            — 居中偏上，均匀分配标题-饼图-图例间距
 *   borderRadius: 6                   — 扇区圆角 6px
 *   borderColor: T.bg                 — 扇区间隔色（与背景融合）
 *   borderWidth: 2                    — 扇区间隔 2px
 *   shadowBlur: 10                    — 扇区阴影
 *   ★ 逐条独立颜色：data[i].itemStyle.color
 *   ★ series.color: 数组              — 图例 Icon 取色
 *
 * 【中心数字】
 *   graphic 主数字                   — 居中显示所有 data.value 的总和
 *   字号: cw × 0.09（自适应容器宽度），字重: 900, 颜色: #e8f4ff, 字体: DIN Alternate
 *   graphic 单位文字                  — 数字下方居中，字号: cw × 0.03, 颜色 #7a9fc4
 *   ★ 防止不同分辨率溢出：字号按容器宽度比例计算
 *
 * 【数据标签】
 *   label.show: true                  — 外部标签（与 #3 玫瑰图一致）
 *   label.color: T.textBright         — 高亮文字色
 *   label.fontSize: 9                 — 字号 9px
 *   label.formatter: '{b}\n{d}%'      — 名称 + 百分比（两行）
 *   label.lineHeight: 14              — 行高 14px
 *   ★ 外径 62% 确保最外层标签不被容器裁切
 *
 * 【布局 — 自适应】
 *   center 48% → 标题-饼图-图例间距均匀
 *   radius 外径 62% → 标签+中心数字空间充足
 *
 * 【图例】
 *   legend.bottom: 0                  — 底部贴边
 *   legend.icon: 'roundRect'          — 统一圆角矩形图标（16×6）
 *   legend.textStyle.color: T.text    — 与轴文字同色 #8ab4d8
 *   legend.textStyle.fontSize: 9      — 字号 9px
 *   ★ 与 #1/#2/#3 图例配置完全一致
 *
 * 【入场动画】
 *   animationType: 'scale'            — 扇区从中心缩放入场
 *   animationDuration: 1200           — 1.2s（与曲线/柱状/玫瑰图一致）
 *   animationEasing: 'elasticOut'     — 弹性缓出
 *   animationDelay: idx × 80          — 逐条延迟 80ms
 *
 * 【自动旋转】
 *   驱动方式: requestAnimationFrame（60fps）
 *   旋转速度: 0.005°/ms ≈ 每圈 72 秒
 *   启动时机: setTimeout 延迟 2s → 等入场动画播完
 *   资源清理: 覆写 ch.dispose() → cancelAnimationFrame
 *   ★ 与 #3 玫瑰图旋转机制完全一致
 *
 * 【交互】
 *   tooltip: item 触发                — {名称}: {值} ({百分比}%)
 *   emphasis.scaleSize: 10            — hover 扇区外扩 10px
 *   emphasis.label.fontSize: 13       — hover 标签放大
 *
 * ============================================================
 *  图表 #5：柱状标注图 barMark — 样式标准
 * ============================================================
 *
 *   ★ 与 #2 barRank 的核心区别：纵向柱状图、带 markPoint/markLine 标注
 *
 * 【生成方法】
 *   ParkCharts.barMark(domId, opts)
 *     domId  - 容器 DOM id（字符串）
 *     opts   - { data:number[], name:string, unit:string, color?:hex, xData?:string[] }
 *     返回值  - ECharts 实例（可调用 .resize() 等）
 *
 * 【视觉规范】
 *   主色      #5470c6（ECharts 默认柱色，可 opts.color 覆盖）
 *   轴文字色  T.text     = #8ab4d8
 *   网格线色  T.text, opacity:.15
 *
 * 【柱体样式】
 *   type: bar（纵向）
 *   barWidth: 自适应，barCategoryGap: '30%'
 *   borderRadius: [2,2,0,0]            — 顶部圆角
 *   itemStyle.color: 渐变或纯色
 *
 * 【标注点 markPoint】
 *   data: [{type:'max', name:'最大'}, {type:'min', name:'最小'}]
 *   ★ 自动标记最大值和最小值
 *
 * 【标注线 markLine】
 *   data: [{type:'average', name:'均值'}]
 *   线色: 与柱体同色，虚线 dashed, 宽 1.5px
 *   标签色: T.textBright, fontSize:9
 *   ★ 自动显示平均值横线，颜色跟随 opts.color
 *
 * 【数据标签】
 *   无外部标签（由 markPoint 标注替代）
 *
 * 【Grid 布局】
 *   left:4, right:48, top:28, bottom:24
 *   containLabel: true
 *   ★ right:48 为 markLine 右侧均值标签留空间（同 #1 曲线图间距原则）
 *
 * 【网格线】
 *   yAxis.splitLine: T.text, opacity:.15
 *
 * 【坐标轴】
 *   xAxis: category, 默认 12 个月（可 opts.xData 覆盖）
 *   yAxis: value, name: unit
 *
 * 【图例】
 *   legend.bottom: 0
 *   legend.icon: 'roundRect', itemWidth:16, itemHeight:6
 *   legend.textStyle: color T.text, fontSize:9
 *   ★ 与 #1/#2/#3/#4 完全一致
 *
 * 【入场动画】
 *   animationDuration: 1200
 *   animationEasing: 'elasticOut'
 *   animationDelay: idx × 100
 *   ★ 与 #2 barRank 一致
 *
 * 【交互】
 *   tooltip: axis 触发
 *   emphasis: shadowBlur:20, shadowColor rgba(64,196,255,.6)
 *
 * ============================================================
 *  图表 #6：堆叠柱状图 barStack — 样式标准
 * ============================================================
 *
 *   ★ 与 #2/#5 的核心区别：多系列堆叠 + 百分比标签
 *
 * 【生成方法】
 *   ParkCharts.barStack(domId, opts)
 *     domId  - 容器 DOM id（字符串）
 *     opts   - {
 *         series: [{name:string, data:number[]}],
 *         xData?: string[],
 *         colors?: string[],
 *         unit?: string
 *       }
 *     返回值  - ECharts 实例（可调用 .resize() 等）
 *
 * 【视觉规范】
 *   堆叠配色  ['#5070dd','#b6d634','#505372','#ff994d','#0ca8df'] 5色
 *   轴文字色  T.text     = #8ab4d8
 *   网格线色  T.text, opacity:.15
 *
 * 【柱体样式】
 *   type: bar, stack: 'total'         — 堆叠模式
 *   barWidth: '60%'                   — 柱宽 60%
 *   ★ 每个系列按传入 colors[] 独立配色
 *   ★ 数据自动转为百分比（值/该列总和），Y轴显示 0-100%
 *
 * 【数据标签】
 *   label.show: true                  — 每段显示原始数值（非百分比）
 *   tooltip 自定义 formatter          — 同样显示原始数值，非百分比
 *
 * 【Y 轴】
 *   max: 1                             — 百分比堆叠上限
 *   axisLabel: v*100+'%'               — Y 轴显示 0-100%
 *
 * 【Grid 布局】
 *   left:4, right:24, top:28, bottom:24, containLabel:true
 *
 * 【图例】
 *   legend.selectedMode: false        — 不可点击切换
 *   legend.icon: 'roundRect', itemWidth:16, itemHeight:6
 *   legend.textStyle: color T.text, fontSize:9
 *
 * 【入场动画】
 *   animationDuration: 1200
 *   animationEasing: 'elasticOut'
 */
(function(global){
  'use strict';

  // ===== 统一主题 =====
  var T = {
    bg: 'rgba(6,14,40,.95)',
    text: '#8ab4d8',
    textBright: '#cfe4ff',
    axis: '#3a5a80',
    gridLine: 'rgba(64,196,255,.08)',
    primary: '#40c4ff',
    colors: ['#40c4ff','#4cd964','#ff9500','#7b61ff','#ff6b81','#ffbb33','#40c4ff','#4cd964','#ff9500']
  };

  var PC = global.ParkCharts = {};

  // =====================
  // 1. 趋势曲线图 ★图表标准#1
  // =====================
  PC.trendLine = function(domId, opts){
    var el = document.getElementById(domId);
    if (!el) return null;
    var ch = echarts.init(el);
    var data = opts.data || [];
    var name = opts.name || '';
    var color = opts.color || T.primary;
    var unit = opts.unit || '';
    var months = ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月'];

    ch.setOption({
      tooltip: {
        trigger: 'axis',
        axisPointer: { type: 'cross', label: { backgroundColor: T.bg } }
      },
      xAxis: {
        type: 'category', boundaryGap: false,
        data: months
      },
      yAxis: {
        type: 'value', name: unit,
        splitLine: { lineStyle: { color: T.text, opacity: .15 } }
      },
      legend: {
        data: [name],
        bottom: 0,
        icon: 'roundRect',
        itemWidth: 16,
        itemHeight: 6,
        textStyle: { color: T.text, fontSize: 9 }
      },
      grid: { left: 4, right: 32, top: 38, bottom: 24, containLabel: true },
      series: [{
        name: name,
        type: 'line',
        color: color,
        data: data,
        smooth: true,
        symbol: 'circle',
        symbolSize: 8,
        lineStyle: { width: 2 },
        areaStyle: {},
        animationDuration: 1200,
        animationEasing: 'elasticOut',
        animationDelay: function(idx){ return idx * 100; },
        label: { show: true, position: 'top', color: T.textBright, fontSize: 9, distance: 6, formatter: '{c}' },
        emphasis: {
          focus: 'series',
          itemStyle: { shadowBlur: 16, shadowColor: color, shadowOffsetY: 2 },
          scale: 1.5
        }
      }]
    });

    return ch;
  };

  // =====================
  // 2. 排名柱状图 ★图表标准#2
  // =====================
  PC.barRank = function(domId, opts){
    var el = document.getElementById(domId);
    if (!el) return null;
    var ch = echarts.init(el);
    var items = (opts.items || []).slice();
    // 降序排列 → inverse:true 让最高值显示在最上面
    items.sort(function(a,b){return b.value-a.value;});
    var name = opts.name || '';
    var unit = opts.unit || '';
    var rankColors = opts.colors || ['#40c4ff','#4cd964','#ff9500','#7b61ff','#ff6b81','#40c4ff','#4cd964','#ff9500','#7b61ff'];

    ch.setOption({
      tooltip: { trigger:'axis', axisPointer:{ type:'shadow' } },
      legend: {
        data: [name],
        bottom: 0,
        icon: 'roundRect',
        itemWidth: 16,
        itemHeight: 6,
        textStyle: { color: T.text, fontSize: 9 }
      },
      // grid.right:45 为右侧标签留空间（同曲线图间距设计原则）
      grid: { left: 4, right: 45, top: 8, bottom: 28, containLabel: true },
      xAxis: {
        type: 'value',
        splitLine: { lineStyle: { color: T.text, opacity: .15 } },
        name: unit
      },
      yAxis: {
        type: 'category',
        data: items.map(function(i){return i.name;}),
        inverse: true,
        axisLabel: {
          color: function(value, index){
            return rankColors[index % rankColors.length];
          }
        },
        axisLine: { lineStyle: { color: T.axis } },
        axisTick: { show: false }
      },
      series: [{
        name: name,
        type: 'bar',
        // series 级 color 供图例 Icon 取色
        color: rankColors[0],
        data: items.map(function(i, idx){
          return {
            value: i.value,
            itemStyle: {
              color: rankColors[idx % rankColors.length],
              borderRadius: [0, 5, 5, 0]
            }
          };
        }),
        barWidth: 16,
        barCategoryGap: '30%',
        animationDuration: 1200,
        animationEasing: 'elasticOut',
        animationDelay: function(idx){ return idx * 100; },
        itemStyle: {
          shadowBlur: 8,
          shadowColor: 'rgba(0,0,0,.5)'
        },
        label: {
          show: true,
          position: 'right',
          color: T.textBright,
          fontSize: 9,
          formatter: '{c}',
          distance: 4
        },
        emphasis: {
          itemStyle: {
            shadowBlur: 20,
            shadowColor: 'rgba(64,196,255,.6)',
            shadowOffsetX: 4
          },
          scale: 1.08
        }
      }]
    });

    return ch;
  };

  // =====================
  // 3. 南丁格尔玫瑰图 ★图表标准#3
  // =====================
  PC.pieRose = function(domId, opts){
    var el = document.getElementById(domId);
    if (!el) return null;
    var ch = echarts.init(el);
    var data = opts.data || [];
    var name = opts.name || '';
    var roseColors = opts.colors || ['#40c4ff','#4cd964','#ff9500','#7b61ff','#ff6b81','#ffbb33','#36d7b7','#a78bfa','#fb7185'];

    ch.setOption({
      tooltip: { trigger:'item', formatter:'{b}: {c} ({d}%)' },
      legend: {
        data: data.map(function(d){return d.name;}),
        bottom: 0,
        icon: 'roundRect',
        itemWidth: 16,
        itemHeight: 6,
        textStyle: { color: T.text, fontSize: 9 }
      },
      // center 48% 均匀分配标题-饼图-图例间距；radius 外径 62% 为标签留空间
      series: [{
        name: name,
        type: 'pie',
        radius: ['22%', '62%'],
        center: ['50%', '48%'],
        roseType: 'area',
        animationType: 'scale',
        animationDuration: 1200,
        animationEasing: 'elasticOut',
        animationDelay: function(idx){ return idx * 80; },
        startAngle: 0,
        // series 级 color 供图例 Icon 取色
        color: roseColors,
        data: data.map(function(d, idx){
          return {
            name: d.name,
            value: d.value,
            itemStyle: {
              color: roseColors[idx % roseColors.length],
              borderRadius: 6,
              borderColor: T.bg,
              borderWidth: 2,
              shadowBlur: 10,
              shadowColor: 'rgba(0,0,0,.4)'
            }
          };
        }),
        label: {
          show: true,
          color: T.textBright,
          fontSize: 9,
          formatter: '{b}\n{d}%',
          lineHeight: 14
        },
        emphasis: {
          scaleSize: 10,
          label: { fontSize: 13, fontWeight: 'bold' }
        }
      }]
    });

    // 自动旋转：等入场动画播完（~2s）再启动，避免 animation:false 打断入场
    var startAngle = 0;
    var lastTime = Date.now();
    function rotatePie(){
      if (!ch || ch.isDisposed()) return;
      var now = Date.now();
      var dt = now - lastTime;
      lastTime = now;
      startAngle = (startAngle + dt * 0.005) % 360;
      ch.setOption({
        series: [{
          startAngle: startAngle,
          animation: false
        }]
      });
      ch._rotateRAF = requestAnimationFrame(rotatePie);
    }
    setTimeout(function(){
      if (!ch || ch.isDisposed()) return;
      lastTime = Date.now();
      ch._rotateRAF = requestAnimationFrame(rotatePie);
    }, 2000);

    // 清理：dispose 时停止旋转
    var origDispose = ch.dispose;
    ch.dispose = function(){
      if (ch._rotateRAF) cancelAnimationFrame(ch._rotateRAF);
      origDispose.call(ch);
    };

    return ch;
  };

  // =====================
  // 4. 环形统计图 ★图表标准#4
  // =====================
  PC.pieRing = function(domId, opts){
    var el = document.getElementById(domId);
    if (!el) return null;
    var ch = echarts.init(el);
    var data = opts.data || [];
    var name = opts.name || '';
    var unit = opts.unit || '';
    var ringColors = opts.colors || ['#40c4ff','#4cd964','#ff9500','#7b61ff','#ff6b81','#ffbb33','#36d7b7','#a78bfa','#fb7185'];

    // 计算数值总和
    var total = 0;
    data.forEach(function(d){ total += d.value; });

    // 中心数字字号按容器宽度自适应
    var cw = el.clientWidth || 300;
    var numSize = Math.round(cw * 0.09);
    var unitSize = Math.round(cw * 0.03);

    ch.setOption({
      tooltip: { trigger:'item', formatter:'{b}: {c} ({d}%)' },
      legend: {
        data: data.map(function(d){return d.name;}),
        bottom: 0,
        icon: 'roundRect',
        itemWidth: 16,
        itemHeight: 6,
        textStyle: { color: T.text, fontSize: 9 }
      },
      // center 48% + radius 外径 62% 为图例、中心数字和外部标签留空间
      series: [{
        name: name,
        type: 'pie',
        radius: ['42%', '62%'],
        center: ['50%', '48%'],
        // series 级 color 供图例 Icon 取色
        color: ringColors,
        data: data.map(function(d, idx){
          return {
            name: d.name,
            value: d.value,
            itemStyle: {
              color: ringColors[idx % ringColors.length],
              borderRadius: 6,
              borderColor: T.bg,
              borderWidth: 2,
              shadowBlur: 10,
              shadowColor: 'rgba(0,0,0,.4)'
            }
          };
        }),
        label: {
          show: true,
          color: T.textBright,
          fontSize: 9,
          formatter: '{b}\n{d}%',
          lineHeight: 14
        },
        animationType: 'scale',
        animationDuration: 1200,
        animationEasing: 'elasticOut',
        animationDelay: function(idx){ return idx * 80; },
        emphasis: {
          scaleSize: 10,
          label: { show: true, fontSize: 13, fontWeight: 'bold', color: T.textBright }
        }
      }],
      // 中心数字：数值总和 + 单位，水平居中
      graphic: [{
        type: 'text',
        left: 'center',
        top: '41%',
        style: {
          text: String(total),
          fontSize: numSize,
          fontWeight: 900,
          fill: '#e8f4ff',
          fontFamily: 'DIN Alternate, Helvetica Neue, sans-serif',
          textShadow: '0 0 10px rgba(200,230,255,0.3)',
          textAlign: 'center'
        }
      },
      unit ? {
        type: 'text',
        left: 'center',
        top: '55%',
        style: {
          text: unit,
          fontSize: unitSize,
          fill: '#7a9fc4',
          textAlign: 'center'
        }
      } : null].filter(Boolean)
    });

    // 自动旋转：与 pieRose 相同机制
    var startAngle = 0;
    var lastTime = Date.now();
    function rotatePie(){
      if (!ch || ch.isDisposed()) return;
      var now = Date.now();
      var dt = now - lastTime;
      lastTime = now;
      startAngle = (startAngle + dt * 0.005) % 360;
      ch.setOption({
        series: [{ startAngle: startAngle, animation: false }]
      });
      ch._rotateRAF = requestAnimationFrame(rotatePie);
    }
    setTimeout(function(){
      if (!ch || ch.isDisposed()) return;
      lastTime = Date.now();
      ch._rotateRAF = requestAnimationFrame(rotatePie);
    }, 2000);

    // 清理
    var origDispose = ch.dispose;
    ch.dispose = function(){
      if (ch._rotateRAF) cancelAnimationFrame(ch._rotateRAF);
      origDispose.call(ch);
    };

    return ch;
  };

  // =====================
  // 5. 柱状标注图 ★图表标准#5
  // =====================
  PC.barMark = function(domId, opts){
    var el = document.getElementById(domId);
    if (!el) return null;
    var ch = echarts.init(el);
    var data = opts.data || [];
    var name = opts.name || '';
    var unit = opts.unit || '';
    var color = opts.color || '#5470c6';
    var xData = opts.xData || ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月'];

    ch.setOption({
      tooltip: { trigger:'axis' },
      legend: {
        data: [name],
        bottom: 0,
        icon: 'roundRect',
        itemWidth: 16,
        itemHeight: 6,
        textStyle: { color: T.text, fontSize: 9 }
      },
      grid: { left:4, right:48, top:28, bottom:24, containLabel: true },
      xAxis: {
        type: 'category',
        data: xData,
        axisTick: { show: false }
      },
      yAxis: {
        type: 'value',
        name: unit,
        splitLine: { lineStyle: { color: T.text, opacity: .15 } }
      },
      series: [{
        name: name,
        type: 'bar',
        color: color,
        data: data,
        barCategoryGap: '30%',
        itemStyle: {
          borderRadius: [2, 2, 0, 0],
          color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
            { offset: 0, color: color },
            { offset: 1, color: 'rgba(43,104,224,0.55)' }
          ])
        },
        animationDuration: 1200,
        animationEasing: 'elasticOut',
        animationDelay: function(idx){ return idx * 100; },
        markPoint: {
          data: [
            { type: 'max', name: '最大' },
            { type: 'min', name: '最小' }
          ],
          symbol: 'pin',
          symbolSize: 32,
          label: { color: T.textBright, fontSize: 9 }
        },
        markLine: {
          data: [{ type: 'average', name: '均值' }],
          lineStyle: { color: color, type: 'dashed', width: 1.5 },
          label: { color: T.textBright, fontSize: 9 }
        },
        emphasis: {
          itemStyle: {
            shadowBlur: 20,
            shadowColor: 'rgba(64,196,255,.6)',
            shadowOffsetX: 4
          }
        }
      }]
    });

    return ch;
  };

  // =====================
  // 6. 堆叠柱状图 ★图表标准#6
  // =====================
  PC.barStack = function(domId, opts){
    var el = document.getElementById(domId);
    if (!el) return null;
    var ch = echarts.init(el);
    var seriesData = opts.series || [];
    var xData = opts.xData || ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月'];
    var colors = opts.colors || ['#5070dd','#b6d634','#505372','#ff994d','#0ca8df'];

    // 计算每列总和
    var n = xData.length;
    var totalData = [];
    for (var i = 0; i < n; i++) {
      var sum = 0;
      for (var j = 0; j < seriesData.length; j++) {
        sum += (seriesData[j].data[i] || 0);
      }
      totalData.push(sum);
    }

    var series = seriesData.map(function(s, sid){
      return {
        name: s.name,
        type: 'bar',
        stack: 'total',
        barWidth: '60%',
        color: colors[sid % colors.length],
        itemStyle: { borderRadius: 0 },
        label: {
          show: true,
          color: '#fff',
          fontSize: 8,
          formatter: function(p){
            return String(seriesData[p.seriesIndex].data[p.dataIndex]);
          }
        },
        data: s.data.map(function(d, did){
          return totalData[did] <= 0 ? 0 : d / totalData[did];
        }),
        animationDuration: 1200,
        animationEasing: 'elasticOut'
      };
    });

    ch.setOption({
      tooltip: {
        trigger:'axis',
        axisPointer:{ type:'shadow' },
        formatter: function(params){
          var tip = params[0].axisValue;
          for (var i = 0; i < params.length; i++) {
            var p = params[i];
            tip += '<br/>' + p.marker + p.seriesName + ': ' + seriesData[p.seriesIndex].data[p.dataIndex];
          }
          return tip;
        }
      },
      legend: {
        data: seriesData.map(function(s){return s.name;}),
        top: 2,
        icon: 'roundRect',
        itemWidth: 16,
        itemHeight: 6,
        selectedMode: false,
        textStyle: { color: T.text, fontSize: 9 }
      },
      grid: { left:4, right:24, top:28, bottom:24, containLabel: true },
      xAxis: {
        type: 'category',
        data: xData,
        axisTick: { show: false }
      },
      yAxis: {
        type: 'value',
        max: 1,
        axisLabel: { formatter: function(v){ return Math.round(v*100)+'%'; } },
        splitLine: { lineStyle: { color: T.text, opacity: .15 } }
      },
      series: series
    });

    return ch;
  };

})(window);
