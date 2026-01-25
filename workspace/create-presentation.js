const PptxGenJS = require('pptxgenjs');
const fs = require('fs');
const path = require('path');

async function createPresentation() {
    // 创建新的演示文稿
    const pptx = new PptxGenJS();
    
    // 设置演示文稿属性
    pptx.defineLayout({ name: 'A4', width: 10, height: 5.625 });
    pptx.layout = 'A4';
    
    // 添加幻灯片1: 封面页
    const slide1 = pptx.addSlide();
    slide1.addText('vCubeVLA运营项目\n2024年度总结', {
        x: 1,
        y: 1.5,
        w: 8,
        h: 2,
        fontSize: 48,
        bold: true,
        color: 'FFFFFF',
        align: 'center',
        fontFace: 'Arial'
    });
    
    slide1.addText('业务指标达成情况汇报', {
        x: 1,
        y: 3,
        w: 8,
        h: 0.5,
        fontSize: 24,
        color: 'AAB7B8',
        align: 'center',
        fontFace: 'Arial'
    });
    
    slide1.addText('汇报人: [您的姓名]', {
        x: 1,
        y: 4,
        w: 8,
        h: 0.4,
        fontSize: 20,
        color: 'FFFFFF',
        align: 'center',
        fontFace: 'Arial'
    });
    
    slide1.addText('2024年12月', {
        x: 1,
        y: 4.4,
        w: 8,
        h: 0.3,
        fontSize: 16,
        color: 'AAB7B8',
        align: 'center',
        fontFace: 'Arial'
    });
    
    // 设置背景色
    slide1.background = { color: '1C2833' };
    
    // 添加幻灯片2: 年度业绩概览
    const slide2 = pptx.addSlide();
    slide2.addText('年度业绩概览', {
        x: 0.5,
        y: 0.3,
        w: 9,
        h: 0.8,
        fontSize: 32,
        bold: true,
        color: 'FFFFFF',
        fontFace: 'Arial'
    });
    
    slide2.addText('vCubeVLA运营项目关键业务指标达成情况', {
        x: 0.5,
        y: 0.9,
        w: 9,
        h: 0.4,
        fontSize: 16,
        color: 'AAB7B8',
        fontFace: 'Arial'
    });
    
    // 添加四个指标卡片
    const metrics = [
        { value: '115%', label: '收入完成率', desc: '超出目标15个百分点', x: 0.5, y: 1.5 },
        { value: '23%', label: '成本节约', desc: '运营效率显著提升', x: 5, y: 1.5 },
        { value: '18%', label: 'ROI提升', desc: '投资回报率稳步增长', x: 0.5, y: 3 },
        { value: '92%', label: '客户满意度', desc: '服务质量获得高度认可', x: 5, y: 3 }
    ];
    
    metrics.forEach(metric => {
        // 指标卡片背景
        slide2.addShape(pptx.ShapeType.rect, {
            x: metric.x,
            y: metric.y,
            w: 4,
            h: 1.2,
            fill: { color: '2E4053', transparency: 50 },
            line: { color: 'E74C3C', width: 4 }
        });
        
        slide2.addText(metric.value, {
            x: metric.x + 0.1,
            y: metric.y + 0.1,
            w: 3.8,
            h: 0.5,
            fontSize: 48,
            bold: true,
            color: 'E74C3C',
            fontFace: 'Arial'
        });
        
        slide2.addText(metric.label, {
            x: metric.x + 0.1,
            y: metric.y + 0.6,
            w: 3.8,
            h: 0.3,
            fontSize: 16,
            color: 'AAB7B8',
            fontFace: 'Arial'
        });
        
        slide2.addText(metric.desc, {
            x: metric.x + 0.1,
            y: metric.y + 0.8,
            w: 3.8,
            h: 0.3,
            fontSize: 14,
            color: 'FFFFFF',
            fontFace: 'Arial'
        });
    });
    
    slide2.background = { color: '1C2833' };
    
    // 添加幻灯片3: 收入增长分析
    const slide3 = pptx.addSlide();
    slide3.addText('收入增长分析', {
        x: 0.5,
        y: 0.3,
        w: 9,
        h: 0.8,
        fontSize: 32,
        bold: true,
        color: 'FFFFFF',
        fontFace: 'Arial'
    });
    
    slide3.addText('vCubeVLA运营项目月度收入趋势与季度对比', {
        x: 0.5,
        y: 0.9,
        w: 9,
        h: 0.4,
        fontSize: 16,
        color: 'AAB7B8',
        fontFace: 'Arial'
    });
    
    // 添加月度收入趋势图表
    slide3.addChart(pptx.ChartType.line, [
        { name: '收入', labels: ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月'], values: [850, 920, 980, 1050, 1120, 1180, 1250, 1320, 1380, 1450, 1520, 1600] }
    ], {
        x: 0.5,
        y: 1.5,
        w: 6,
        h: 2,
        title: '月度收入趋势',
        showLegend: true,
        legendPos: 'b',
        catAxisTitle: '月份',
        valAxisTitle: '收入(万元)',
        dataLabelFormatCode: '#,##0',
        lineDataSymbol: 'circle',
        lineSize: 3,
        chartColors: ['E74C3C']
    });
    
    // 添加关键指标
    slide3.addText('28%', {
        x: 7,
        y: 1.5,
        w: 2,
        h: 0.8,
        fontSize: 36,
        bold: true,
        color: 'E74C3C',
        fontFace: 'Arial'
    });
    
    slide3.addText('同比增长率', {
        x: 7,
        y: 2.1,
        w: 2,
        h: 0.3,
        fontSize: 14,
        color: 'AAB7B8',
        fontFace: 'Arial'
    });
    
    slide3.addText('115%', {
        x: 7,
        y: 2.8,
        w: 2,
        h: 0.8,
        fontSize: 36,
        bold: true,
        color: 'E74C3C',
        fontFace: 'Arial'
    });
    
    slide3.addText('目标达成率', {
        x: 7,
        y: 3.4,
        w: 2,
        h: 0.3,
        fontSize: 14,
        color: 'AAB7B8',
        fontFace: 'Arial'
    });
    
    slide3.background = { color: '1C2833' };
    
    // 添加幻灯片4: 成本优化成果
    const slide4 = pptx.addSlide();
    slide4.addText('成本优化成果', {
        x: 0.5,
        y: 0.3,
        w: 9,
        h: 0.8,
        fontSize: 32,
        bold: true,
        color: 'FFFFFF',
        fontFace: 'Arial'
    });
    
    slide4.addText('vCubeVLA运营项目成本结构优化与节约分析', {
        x: 0.5,
        y: 0.9,
        w: 9,
        h: 0.4,
        fontSize: 16,
        color: 'AAB7B8',
        fontFace: 'Arial'
    });
    
    // 添加成本节约指标
    slide4.addText('23%', {
        x: 1,
        y: 1.5,
        w: 2,
        h: 0.8,
        fontSize: 42,
        bold: true,
        color: 'E74C3C',
        fontFace: 'Arial'
    });
    
    slide4.addText('总体成本节约', {
        x: 1,
        y: 2.1,
        w: 2,
        h: 0.3,
        fontSize: 14,
        color: 'AAB7B8',
        fontFace: 'Arial'
    });
    
    slide4.addText('18%', {
        x: 3.5,
        y: 1.5,
        w: 2,
        h: 0.8,
        fontSize: 42,
        bold: true,
        color: 'E74C3C',
        fontFace: 'Arial'
    });
    
    slide4.addText('运营效率提升', {
        x: 3.5,
        y: 2.1,
        w: 2,
        h: 0.3,
        fontSize: 14,
        color: 'AAB7B8',
        fontFace: 'Arial'
    });
    
    slide4.addText('¥2.3M', {
        x: 6,
        y: 1.5,
        w: 2,
        h: 0.8,
        fontSize: 42,
        bold: true,
        color: 'E74C3C',
        fontFace: 'Arial'
    });
    
    slide4.addText('年度节约金额', {
        x: 6,
        y: 2.1,
        w: 2,
        h: 0.3,
        fontSize: 14,
        color: 'AAB7B8',
        fontFace: 'Arial'
    });
    
    // 添加成本结构饼图
    slide4.addChart(pptx.ChartType.pie, [
        { name: '成本结构', labels: ['人力成本', '技术成本', '运营成本', '其他'], values: [45, 30, 20, 5] }
    ], {
        x: 0.5,
        y: 2.8,
        w: 4,
        h: 2,
        title: '优化后成本结构',
        showLegend: true,
        legendPos: 'r',
        dataLabelFormatCode: '#,##0%',
        chartColors: ['E74C3C', '2E4053', 'AAB7B8', 'FFFFFF']
    });
    
    // 添加成本节约明细
    slide4.addText('成本节约明细', {
        x: 5,
        y: 2.8,
        w: 4,
        h: 0.4,
        fontSize: 16,
        bold: true,
        color: 'E74C3C',
        fontFace: 'Arial'
    });
    
    const savings = [
        { item: '人力成本优化', value: '-12%' },
        { item: '技术架构优化', value: '-8%' },
        { item: '运营流程改进', value: '-3%' },
        { item: '资源利用率提升', value: '+15%' }
    ];
    
    savings.forEach((saving, index) => {
        const yPos = 3.3 + index * 0.35;
        slide4.addText(saving.item, {
            x: 5,
            y: yPos,
            w: 2.5,
            h: 0.3,
            fontSize: 14,
            color: 'AAB7B8',
            fontFace: 'Arial'
        });
        
        slide4.addText(saving.value, {
            x: 7.5,
            y: yPos,
            w: 1.5,
            h: 0.3,
            fontSize: 14,
            bold: true,
            color: 'FFFFFF',
            fontFace: 'Arial'
        });
    });
    
    slide4.background = { color: '1C2833' };
    
    // 添加幻灯片5: 项目亮点总结
    const slide5 = pptx.addSlide();
    slide5.addText('项目亮点总结', {
        x: 0.5,
        y: 0.3,
        w: 9,
        h: 0.8,
        fontSize: 32,
        bold: true,
        color: 'FFFFFF',
        fontFace: 'Arial'
    });
    
    slide5.addText('vCubeVLA运营项目关键成就与团队贡献', {
        x: 0.5,
        y: 0.9,
        w: 9,
        h: 0.4,
        fontSize: 16,
        color: 'AAB7B8',
        fontFace: 'Arial'
    });
    
    // 添加三个关键成就
    const achievements = [
        {
            title: '智能化运营升级',
            details: ['AI算法优化效率40%', '自动化处理率提升至85%', '运维成本降低35%'],
            x: 0.5,
            y: 1.5
        },
        {
            title: '客户体验优化',
            details: ['客户满意度达92%', '响应时间缩短60%', '客户留存率提升25%'],
            x: 3,
            y: 1.5
        },
        {
            title: '技术架构创新',
            details: ['微服务架构升级完成', '系统可用性达99.9%', '并发处理能力提升3倍'],
            x: 5.5,
            y: 1.5
        }
    ];
    
    achievements.forEach((achievement, index) => {
        // 成就卡片背景
        slide5.addShape(pptx.ShapeType.rect, {
            x: achievement.x,
            y: achievement.y,
            w: 2.2,
            h: 1.8,
            fill: { color: '2E4053', transparency: 50 },
            line: { color: 'E74C3C', width: 4 }
        });
        
        slide5.addText(`0${index + 1}`, {
            x: achievement.x + 0.1,
            y: achievement.y + 0.1,
            w: 2,
            h: 0.5,
            fontSize: 48,
            bold: true,
            color: 'E74C3C',
            fontFace: 'Arial'
        });
        
        slide5.addText(achievement.title, {
            x: achievement.x + 0.1,
            y: achievement.y + 0.6,
            w: 2,
            h: 0.3,
            fontSize: 16,
            bold: true,
            color: 'FFFFFF',
            fontFace: 'Arial'
        });
        
        achievement.details.forEach((detail, detailIndex) => {
            slide5.addText(`• ${detail}`, {
                x: achievement.x + 0.1,
                y: achievement.y + 0.9 + detailIndex * 0.25,
                w: 2,
                h: 0.2,
                fontSize: 12,
                color: 'AAB7B8',
                fontFace: 'Arial'
            });
        });
    });
    
    // 添加团队贡献量化
    slide5.addShape(pptx.ShapeType.rect, {
        x: 0.5,
        y: 3.5,
        w: 9,
        h: 1.5,
        fill: { color: 'E74C3C', transparency: 90 },
        line: { color: 'E74C3C', width: 1 }
    });
    
    slide5.addText('团队贡献量化', {
        x: 0.7,
        y: 3.6,
        w: 8.6,
        h: 0.4,
        fontSize: 16,
        bold: true,
        color: 'E74C3C',
        fontFace: 'Arial'
    });
    
    const teamMetrics = [
        { label: '15+', desc: '团队成员' },
        { label: '3,000+', desc: '工作小时' },
        { label: '12', desc: '创新专利' },
        { label: '98%', desc: '项目完成率' }
    ];
    
    teamMetrics.forEach((metric, index) => {
        const xPos = 0.7 + index * 2.2;
        slide5.addText(metric.label, {
            x: xPos,
            y: 4.1,
            w: 2,
            h: 0.4,
            fontSize: 24,
            bold: true,
            color: 'E74C3C',
            fontFace: 'Arial'
        });
        
        slide5.addText(metric.desc, {
            x: xPos,
            y: 4.4,
            w: 2,
            h: 0.3,
            fontSize: 12,
            color: 'AAB7B8',
            fontFace: 'Arial'
        });
    });
    
    slide5.background = { color: '1C2833' };
    
    // 添加幻灯片6: 2025年规划
    const slide6 = pptx.addSlide();
    slide6.addText('2025年规划', {
        x: 0.5,
        y: 0.3,
        w: 9,
        h: 0.8,
        fontSize: 32,
        bold: true,
        color: 'FFFFFF',
        fontFace: 'Arial'
    });
    
    slide6.addText('vCubeVLA运营项目下一年度目标与关键举措', {
        x: 0.5,
        y: 0.9,
        w: 9,
        h: 0.4,
        fontSize: 16,
        color: 'AAB7B8',
        fontFace: 'Arial'
    });
    
    // 添加2025年目标
    const goals = [
        { value: '130%', label: '收入增长目标', desc: '基于2024年基础，实现收入同比增长30%' },
        { value: '30%', label: '效率提升目标', desc: '通过技术创新和流程优化，实现运营效率再提升30%' },
        { value: '95%', label: '客户满意度目标', desc: '持续优化服务质量，客户满意度提升至95%以上' }
    ];
    
    goals.forEach((goal, index) => {
        const yPos = 1.5 + index * 0.7;
        
        slide6.addShape(pptx.ShapeType.rect, {
            x: 0.5,
            y: yPos,
            w: 5.5,
            h: 0.6,
            fill: { color: '2E4053', transparency: 50 },
            line: { color: 'E74C3C', width: 4 }
        });
        
        slide6.addText(goal.value, {
            x: 0.6,
            y: yPos + 0.05,
            w: 1.5,
            h: 0.5,
            fontSize: 36,
            bold: true,
            color: 'E74C3C',
            fontFace: 'Arial'
        });
        
        slide6.addText(goal.label, {
            x: 2.2,
            y: yPos + 0.1,
            w: 2,
            h: 0.3,
            fontSize: 16,
            bold: true,
            color: 'FFFFFF',
            fontFace: 'Arial'
        });
        
        slide6.addText(goal.desc, {
            x: 2.2,
            y: yPos + 0.35,
            w: 3.6,
            h: 0.2,
            fontSize: 12,
            color: 'AAB7B8',
            fontFace: 'Arial'
        });
    });
    
    // 添加关键举措
    slide6.addText('关键举措', {
        x: 6.5,
        y: 1.5,
        w: 3,
        h: 0.4,
        fontSize: 18,
        bold: true,
        color: 'E74C3C',
        fontFace: 'Arial'
    });
    
    const initiatives = [
        { title: '🚀 智能化升级', items: ['AI算法2.0版本', '自动化覆盖率90%'] },
        { title: '🌐 市场扩展', items: ['新增3个区域市场', '客户基数翻倍'] },
        { title: '⚡ 技术创新', items: ['云原生架构升级', '性能提升5倍'] }
    ];
    
    initiatives.forEach((initiative, index) => {
        const yPos = 1.9 + index * 0.7;
        
        slide6.addShape(pptx.ShapeType.rect, {
            x: 6.5,
            y: yPos,
            w: 3,
            h: 0.6,
            fill: { color: '2E4053', transparency: 30 },
            line: { color: 'E74C3C', width: 1 }
        });
        
        slide6.addText(initiative.title, {
            x: 6.6,
            y: yPos + 0.05,
            w: 2.8,
            h: 0.3,
            fontSize: 14,
            bold: true,
            color: 'E74C3C',
            fontFace: 'Arial'
        });
        
        initiative.items.forEach((item, itemIndex) => {
            slide6.addText(`• ${item}`, {
                x: 6.6,
                y: yPos + 0.3 + itemIndex * 0.15,
                w: 2.8,
                h: 0.15,
                fontSize: 11,
                color: 'AAB7B8',
                fontFace: 'Arial'
            });
        });
    });
    
    // 添加季度里程碑
    slide6.addShape(pptx.ShapeType.rect, {
        x: 0.5,
        y: 3.6,
        w: 9,
        h: 1.2,
        fill: { color: 'E74C3C', transparency: 90 },
        line: { color: 'E74C3C', width: 1 }
    });
    
    slide6.addText('季度里程碑', {
        x: 0.7,
        y: 3.7,
        w: 8.6,
        h: 0.3,
        fontSize: 16,
        bold: true,
        color: 'E74C3C',
        fontFace: 'Arial'
    });
    
    const quarters = [
        { label: 'Q1', goal: '基础建设' },
        { label: 'Q2', goal: '试点上线' },
        { label: 'Q3', goal: '全面推广' },
        { label: 'Q4', goal: '优化迭代' }
    ];
    
    quarters.forEach((quarter, index) => {
        const xPos = 0.7 + index * 2.2;
        
        slide6.addShape(pptx.ShapeType.rect, {
            x: xPos,
            y: 4.0,
            w: 2.1,
            h: 0.7,
            fill: { color: '2E4053', transparency: 50 },
            line: { color: 'E74C3C', width: 1 }
        });
        
        slide6.addText(quarter.label, {
            x: xPos,
            y: 4.1,
            w: 2.1,
            h: 0.25,
            fontSize: 14,
            bold: true,
            color: 'E74C3C',
            align: 'center',
            fontFace: 'Arial'
        });
        
        slide6.addText(quarter.goal, {
            x: xPos,
            y: 4.35,
            w: 2.1,
            h: 0.25,
            fontSize: 12,
            color: 'AAB7B8',
            align: 'center',
            fontFace: 'Arial'
        });
    });
    
    slide6.background = { color: '1C2833' };
    
    // 保存演示文稿
    await pptx.writeFile({ fileName: 'vCubeVLA年度总结.pptx' });
    console.log('演示文稿已生成: vCubeVLA年度总结.pptx');
}

// 执行创建函数
createPresentation().catch(console.error);