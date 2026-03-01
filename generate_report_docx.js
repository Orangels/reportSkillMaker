const { Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle } = require('docx');
const fs = require('fs');

// 读取数据
const data = JSON.parse(fs.readFileSync('/home/orangels/xm_dev/ls_dev/reportSkillMaker/output/extracted_data.json', 'utf8'));

// 计算环比
const totalChangeRate = ((data.overall_stats.valid_cases_dec - data.overall_stats.valid_cases_nov) / data.overall_stats.valid_cases_nov * 100).toFixed(1);

// 创建文档
const doc = new Document({
  styles: {
    default: {
      document: {
        run: {
          font: "仿宋",
          size: 32  // 16磅 = 32半磅
        }
      }
    },
    paragraphStyles: [
      {
        id: "Normal",
        name: "Normal",
        run: {
          font: "仿宋",
          size: 32
        },
        paragraph: {
          spacing: {
            line: 560,  // 28磅行距
            lineRule: "exact"
          },
          indent: {
            firstLine: 640  // 2字符首行缩进
          }
        }
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: {
          width: 12240,   // US Letter width
          height: 15840   // US Letter height
        },
        margin: {
          top: 1440,
          right: 1800,
          bottom: 1440,
          left: 1800
        }
      }
    },
    children: [
      // 发文单位(红色大标题)
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 },
        children: [
          new TextRun({
            text: "临高县公安局情报指挥中心",
            font: "方正小标宋简体",
            size: 110,  // 55磅
            color: "FF0000",
            bold: true
          })
        ]
      }),

      // 红色横线
      new Paragraph({
        border: {
          bottom: {
            style: BorderStyle.SINGLE,
            size: 12,
            color: "FF0000",
            space: 1
          }
        },
        spacing: { after: 400 }
      }),

      // 主标题
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 400 },
        children: [
          new TextRun({
            text: "关于12月份警情分析的报告",
            font: "方正小标宋简体",
            size: 44,  // 22磅
            bold: true
          })
        ]
      }),

      // 一、整体情况
      new Paragraph({
        spacing: { before: 200, after: 200 },
        children: [
          new TextRun({
            text: "一、整体情况",
            font: "黑体",
            size: 32,
            bold: true
          })
        ]
      }),

      // 整体情况内容
      new Paragraph({
        indent: { firstLine: 640 },
        children: [
          new TextRun({
            text: `${data.report_info.period}我局共接报`,
            font: "仿宋",
            size: 32
          }),
          new TextRun({
            text: "有效警情",
            font: "仿宋",
            size: 32,
            bold: true
          }),
          new TextRun({
            text: `${data.overall_stats.valid_cases_dec}起（`,
            font: "仿宋",
            size: 32
          }),
          new TextRun({
            text: "不含骚扰警情",
            font: "仿宋",
            size: 32,
            bold: true
          }),
          new TextRun({
            text: `${data.overall_stats.harassment_cases_dec}起），环比上升${totalChangeRate}%。其中`,
            font: "仿宋",
            size: 32
          }),
          new TextRun({
            text: "刑事警情",
            font: "仿宋",
            size: 32,
            bold: true
          }),
          new TextRun({
            text: `${data.six_categories['刑事警情'].dec_count}起，环比下降${Math.abs(data.six_categories['刑事警情'].change_rate)}%；`,
            font: "仿宋",
            size: 32
          }),
          new TextRun({
            text: "治安警情",
            font: "仿宋",
            size: 32,
            bold: true
          }),
          new TextRun({
            text: `${data.six_categories['治安警情'].dec_count}起，环比下降${Math.abs(data.six_categories['治安警情'].change_rate)}%；`,
            font: "仿宋",
            size: 32
          }),
          new TextRun({
            text: "交通警情",
            font: "仿宋",
            size: 32,
            bold: true
          }),
          new TextRun({
            text: `${data.six_categories['交通警情'].dec_count}起，环比上升${data.six_categories['交通警情'].change_rate}%；`,
            font: "仿宋",
            size: 32
          }),
          new TextRun({
            text: "纠纷警情",
            font: "仿宋",
            size: 32,
            bold: true
          }),
          new TextRun({
            text: `${data.six_categories['纠纷警情'].dec_count}起，环比上升${data.six_categories['纠纷警情'].change_rate}%；`,
            font: "仿宋",
            size: 32
          }),
          new TextRun({
            text: "群众紧急求助",
            font: "仿宋",
            size: 32,
            bold: true
          }),
          new TextRun({
            text: `${data.six_categories['群众紧急求助'].dec_count}起，环比上升${data.six_categories['群众紧急求助'].change_rate}%；`,
            font: "仿宋",
            size: 32
          }),
          new TextRun({
            text: "其他警情",
            font: "仿宋",
            size: 32,
            bold: true
          }),
          new TextRun({
            text: `${data.six_categories['其他警情'].dec_count}起，环比上升${data.six_categories['其他警情'].change_rate}%。`,
            font: "仿宋",
            size: 32
          })
        ]
      }),

      // 二、上升警情类别分布
      new Paragraph({
        spacing: { before: 200, after: 200 },
        children: [
          new TextRun({
            text: "二、上升警情类别分布",
            font: "黑体",
            size: 32,
            bold: true
          })
        ]
      }),

      // (一)交通警情分析
      new Paragraph({
        spacing: { before: 100, after: 100 },
        children: [
          new TextRun({
            text: "(一)交通警情分析",
            font: "黑体",
            size: 32
          })
        ]
      }),

      new Paragraph({
        indent: { firstLine: 640 },
        children: [
          new TextRun({
            text: "我局共接报",
            font: "仿宋",
            size: 32
          }),
          new TextRun({
            text: "交通警情",
            font: "仿宋",
            size: 32,
            bold: true
          }),
          new TextRun({
            text: `${data.six_categories['交通警情'].dec_count}起，环比大幅上升${data.six_categories['交通警情'].change_rate}%。`,
            font: "仿宋",
            size: 32
          })
        ]
      }),

      // 1.从警情类别分析
      new Paragraph({
        indent: { firstLine: 640 },
        children: [
          new TextRun({
            text: "1.从警情类别分析。",
            font: "仿宋",
            size: 32,
            bold: true
          }),
          new TextRun({
            text: `主要集中在机动车与机动车事故${data.traffic_detail.subtypes['机动车与机动车事故'].dec_count}起，其次机动车与非机动车事故${data.traffic_detail.subtypes['机动车与非机动车事故'].dec_count}起，单方事故${data.traffic_detail.subtypes['单方事故'].dec_count}起。`,
            font: "仿宋",
            size: 32
          })
        ]
      }),

      // 2.从发案时间分析
      new Paragraph({
        indent: { firstLine: 640 },
        children: [
          new TextRun({
            text: "2.从发案时间分析。",
            font: "仿宋",
            size: 32,
            bold: true
          }),
          new TextRun({
            text: `交通警情时段分布特征明显，${data.time_distribution.peak_hours[0].hour_range}最为集中，累计发生${data.time_distribution.peak_hours[0].count}起；其次${data.time_distribution.peak_hours[1].hour_range}，累计发生${data.time_distribution.peak_hours[1].count}起，两个时段均对应上午出行、下午通勤高峰。`,
            font: "仿宋",
            size: 32
          })
        ]
      }),

      // 小结
      new Paragraph({
        indent: { firstLine: 640 },
        children: [
          new TextRun({
            text: "小结：",
            font: "仿宋",
            size: 32,
            bold: true
          }),
          new TextRun({
            text: "本月交通警情环比大幅上升，增幅显著。此类警情特征突出：",
            font: "仿宋",
            size: 32
          }),
          new TextRun({
            text: "一是",
            font: "仿宋",
            size: 32,
            bold: true
          }),
          new TextRun({
            text: "机动车事故类型集中，机动车与机动车、机动车与非机动车事故占比较高，反映出路面车流量增大带来的安全风险；",
            font: "仿宋",
            size: 32
          }),
          new TextRun({
            text: "二是",
            font: "仿宋",
            size: 32,
            bold: true
          }),
          new TextRun({
            text: "时段分布特征明显，集中在上午和下午通勤高峰时段，需针对性加强重点时段路面管控；",
            font: "仿宋",
            size: 32
          }),
          new TextRun({
            text: "三是",
            font: "仿宋",
            size: 32,
            bold: true
          }),
          new TextRun({
            text: "年末岁尾人流车流密集，交通安全形势严峻，需持续强化交通安全宣传和路面执法力度。",
            font: "仿宋",
            size: 32
          })
        ]
      }),

      // (二)涉刀警情分析
      new Paragraph({
        spacing: { before: 100, after: 100 },
        children: [
          new TextRun({
            text: "(二)涉刀警情分析",
            font: "黑体",
            size: 32
          })
        ]
      }),

      new Paragraph({
        indent: { firstLine: 640 },
        children: [
          new TextRun({
            text: "我局共接报",
            font: "仿宋",
            size: 32
          }),
          new TextRun({
            text: "涉刀警情",
            font: "仿宋",
            size: 32,
            bold: true
          }),
          new TextRun({
            text: `${data.knife_cases.total}起，环比大幅上升。`,
            font: "仿宋",
            size: 32
          })
        ]
      }),

      // 1.从警情类型分析
      new Paragraph({
        indent: { firstLine: 640 },
        children: [
          new TextRun({
            text: "1.从警情类型分析。",
            font: "仿宋",
            size: 32,
            bold: true
          }),
          new TextRun({
            text: `治安警情${data.knife_cases.by_category['治安警情']}起，群众紧急求助${data.knife_cases.by_category['群众紧急求助']}起，其他警情${data.knife_cases.by_category['其他警情']}起。其中治安警情占比最高，主要为殴打他人、故意伤害案件。`,
            font: "仿宋",
            size: 32
          })
        ]
      }),

      // 2.从辖区分布分析
      new Paragraph({
        indent: { firstLine: 640 },
        children: [
          new TextRun({
            text: "2.从辖区分布分析。",
            font: "仿宋",
            size: 32,
            bold: true
          }),
          new TextRun({
            text: `主要发生在临高临城西门派出所${data.knife_cases.by_jurisdiction['临高临城西门派出所']}起，临高临城东门派出所${data.knife_cases.by_jurisdiction['临高临城东门派出所']}起，临高多文派出所${data.knife_cases.by_jurisdiction['临高多文派出所']}起。`,
            font: "仿宋",
            size: 32
          })
        ]
      }),

      // 小结
      new Paragraph({
        indent: { firstLine: 640 },
        children: [
          new TextRun({
            text: "小结：",
            font: "仿宋",
            size: 32,
            bold: true
          }),
          new TextRun({
            text: "本月涉刀警情环比大幅上升，安全风险突出。此类警情多发生在西门所、东门所等城区派出所辖区，且多与治安殴打、故意伤害案件相关，造成当事人不同程度人身伤害。需严格落实持刀人住所清查部署，加强重点人员管控，依法从严惩处涉刀违法犯罪，有效遏制此类警情上升势头。",
            font: "仿宋",
            size: 32
          })
        ]
      }),

      // (三)涉未成人警情分析
      new Paragraph({
        spacing: { before: 100, after: 100 },
        children: [
          new TextRun({
            text: "(三)涉未成人警情分析",
            font: "黑体",
            size: 32
          })
        ]
      }),

      new Paragraph({
        indent: { firstLine: 640 },
        children: [
          new TextRun({
            text: "我局共接报",
            font: "仿宋",
            size: 32
          }),
          new TextRun({
            text: "涉未成人警情",
            font: "仿宋",
            size: 32,
            bold: true
          }),
          new TextRun({
            text: `${data.minor_cases.total}起，环比显著上升。`,
            font: "仿宋",
            size: 32
          })
        ]
      }),

      // 1.从警情类型分析
      new Paragraph({
        indent: { firstLine: 640 },
        children: [
          new TextRun({
            text: "1.从警情类型分析。",
            font: "仿宋",
            size: 32,
            bold: true
          }),
          new TextRun({
            text: `主要集中在其他警情${data.minor_cases.by_category['其他警情']}起，其次群众紧急求助${data.minor_cases.by_category['群众紧急求助']}起，交通警情${data.minor_cases.by_category['交通警情']}起。`,
            font: "仿宋",
            size: 32
          })
        ]
      }),

      // 小结
      new Paragraph({
        indent: { firstLine: 640 },
        children: [
          new TextRun({
            text: "小结：",
            font: "仿宋",
            size: 32,
            bold: true
          }),
          new TextRun({
            text: "本月涉未成人警情环比显著上升，其他警情和群众紧急求助占比较高。年末岁尾未成年人安全风险增加，需持续强化护苗行动、法治宣讲校园行，加强晚安守护巡逻，精准防范处置，全力守护未成年人健康成长。",
            font: "仿宋",
            size: 32
          })
        ]
      }),

      // (四)金牌港重点园区警情分析
      new Paragraph({
        spacing: { before: 100, after: 100 },
        children: [
          new TextRun({
            text: "(四)金牌港重点园区警情分析",
            font: "黑体",
            size: 32
          })
        ]
      }),

      new Paragraph({
        indent: { firstLine: 640 },
        children: [
          new TextRun({
            text: `我局共接报涉金牌港重点园区警情${data.jinpaigang_cases.total}起，环比上升。`,
            font: "仿宋",
            size: 32
          })
        ]
      }),

      // 从警情类型分析
      new Paragraph({
        indent: { firstLine: 640 },
        children: [
          new TextRun({
            text: "从警情类型分析，",
            font: "仿宋",
            size: 32,
            bold: true
          }),
          new TextRun({
            text: `主要集中在交通警情${data.jinpaigang_cases.by_category['交通警情']}起，其次其他警情${data.jinpaigang_cases.by_category['其他警情']}起，群众紧急求助${data.jinpaigang_cases.by_category['群众紧急求助']}起。`,
            font: "仿宋",
            size: 32
          })
        ]
      }),

      // 小结
      new Paragraph({
        indent: { firstLine: 640 },
        children: [
          new TextRun({
            text: "小结：",
            font: "仿宋",
            size: 32,
            bold: true
          }),
          new TextRun({
            text: "本月金牌港重点园区警情环比上升，交通警情占比最高，反映出园区车流量增大带来的管理压力。需针对性强化园区交通管理和安全防范，优化接处警流程，全力维护园区治安秩序稳定。",
            font: "仿宋",
            size: 32
          })
        ]
      }),

      // 三、下降警情类型分布
      new Paragraph({
        spacing: { before: 200, after: 200 },
        children: [
          new TextRun({
            text: "三、下降警情类型分布",
            font: "黑体",
            size: 32,
            bold: true
          })
        ]
      }),

      // (一)刑事警情分析
      new Paragraph({
        spacing: { before: 100, after: 100 },
        children: [
          new TextRun({
            text: "(一)刑事警情分析",
            font: "黑体",
            size: 32
          })
        ]
      }),

      new Paragraph({
        indent: { firstLine: 640 },
        children: [
          new TextRun({
            text: "我局共接报",
            font: "仿宋",
            size: 32
          }),
          new TextRun({
            text: "刑事警情",
            font: "仿宋",
            size: 32,
            bold: true
          }),
          new TextRun({
            text: `${data.six_categories['刑事警情'].dec_count}起，环比下降${Math.abs(data.six_categories['刑事警情'].change_rate)}%。`,
            font: "仿宋",
            size: 32
          })
        ]
      }),

      // 小结
      new Paragraph({
        indent: { firstLine: 640 },
        children: [
          new TextRun({
            text: "小结：",
            font: "仿宋",
            size: 32,
            bold: true
          }),
          new TextRun({
            text: "本月刑事警情环比大幅下降，管控成效显著，需持续保持严打高压态势，巩固防控成果。",
            font: "仿宋",
            size: 32
          })
        ]
      }),

      // (二)治安警情分析
      new Paragraph({
        spacing: { before: 100, after: 100 },
        children: [
          new TextRun({
            text: "(二)治安警情分析",
            font: "黑体",
            size: 32
          })
        ]
      }),

      new Paragraph({
        indent: { firstLine: 640 },
        children: [
          new TextRun({
            text: "我局共接报",
            font: "仿宋",
            size: 32
          }),
          new TextRun({
            text: "治安警情",
            font: "仿宋",
            size: 32,
            bold: true
          }),
          new TextRun({
            text: `${data.six_categories['治安警情'].dec_count}起，环比下降${Math.abs(data.six_categories['治安警情'].change_rate)}%。`,
            font: "仿宋",
            size: 32
          })
        ]
      }),

      // 小结
      new Paragraph({
        indent: { firstLine: 640 },
        children: [
          new TextRun({
            text: "小结：",
            font: "仿宋",
            size: 32,
            bold: true
          }),
          new TextRun({
            text: "本月治安警情环比呈下降态势，治安防控工作成效明显，需持续强化重点区域巡逻防控，巩固良好态势。",
            font: "仿宋",
            size: 32
          })
        ]
      }),

      // 四、工作建议
      new Paragraph({
        spacing: { before: 200, after: 200 },
        children: [
          new TextRun({
            text: "四、工作建议",
            font: "黑体",
            size: 32,
            bold: true
          })
        ]
      }),

      // (一)强化交通安全管控
      new Paragraph({
        spacing: { before: 100, after: 100 },
        children: [
          new TextRun({
            text: "(一)强化交通安全管控",
            font: "黑体",
            size: 32
          })
        ]
      }),

      new Paragraph({
        indent: { firstLine: 640 },
        children: [
          new TextRun({
            text: "交警部门要针对本月交通警情大幅上升态势，在每日10时至12时、14时至15时重点时段，对主干道、学校周边、商圈路口增派警力，加强路面巡逻管控；强化交通安全宣传引导，提升驾驶员安全意识；严查交通违法行为，依法从严处罚，全力压降交通警情及交通事故。",
            font: "仿宋",
            size: 32
          })
        ]
      }),

      // (二)严打涉刀违法犯罪
      new Paragraph({
        spacing: { before: 100, after: 100 },
        children: [
          new TextRun({
            text: "(二)严打涉刀违法犯罪",
            font: "黑体",
            size: 32
          })
        ]
      }),

      new Paragraph({
        indent: { firstLine: 640 },
        children: [
          new TextRun({
            text: "西门所、东门所等涉刀警情高发单位要严格落实持刀人住所清查部署工作要求，建立重点人员台账，加强日常管控；对持刀伤人案件快侦快办、依法从严惩处，在辖区形成'带刀必查、涉刀必罚'的有力震慑，切实压降涉刀警情，有效遏制此类警情上升势头。",
            font: "仿宋",
            size: 32
          })
        ]
      }),

      // (三)多举措守护未成年人安全
      new Paragraph({
        spacing: { before: 100, after: 100 },
        children: [
          new TextRun({
            text: "(三)多举措守护未成年人安全",
            font: "黑体",
            size: 32
          })
        ]
      }),

      new Paragraph({
        indent: { firstLine: 640 },
        children: [
          new TextRun({
            text: "持续强化涉未成人法制宣传教育，全覆盖普及法律知识与自我防护技能；各派出所要严格落实'晚安行动'要求，组织民辅警对网吧、河边、公园等未成年人易聚集区域开展巡查；严厉打击各类涉未成人违法犯罪，形成有力震慑，切实守护未成年人身心健康与安全成长。",
            font: "仿宋",
            size: 32
          })
        ]
      }),

      // (四)优化园区治安管理
      new Paragraph({
        spacing: { before: 100, after: 100 },
        children: [
          new TextRun({
            text: "(四)优化园区治安管理",
            font: "黑体",
            size: 32
          })
        ]
      }),

      new Paragraph({
        indent: { firstLine: 640 },
        children: [
          new TextRun({
            text: "马袅海岸派出所要紧盯金牌重点园区警情变化，聚焦交通警情占比较高问题，深化警情研判预警，精准把握防控重点；加强园区交通秩序管理，优化接处警流程、提升处置效能。全力压降警情总量、防范风险隐患，切实维护金牌重点园区治安秩序稳定。",
            font: "仿宋",
            size: 32
          })
        ]
      }),

      // 落款
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        spacing: { before: 400 },
        children: [
          new TextRun({
            text: "临高县公安局情报指挥中心",
            font: "仿宋",
            size: 32
          })
        ]
      }),

      new Paragraph({
        alignment: AlignmentType.RIGHT,
        children: [
          new TextRun({
            text: data.report_info.report_date,
            font: "仿宋",
            size: 32
          })
        ]
      }),

      // 抄送抄报
      new Paragraph({
        spacing: { before: 400 },
        children: [
          new TextRun({
            text: "抄送：各所、队、室(中心)",
            font: "仿宋",
            size: 32
          })
        ]
      }),

      new Paragraph({
        children: [
          new TextRun({
            text: "抄报：严树勋副县长，各局领导",
            font: "仿宋",
            size: 32
          })
        ]
      }),

      new Paragraph({
        children: [
          new TextRun({
            text: `临高县公安局情报指挥中心 ${data.report_info.report_date}印发`,
            font: "仿宋",
            size: 32
          })
        ]
      })
    ]
  }]
});

// 保存文档
Packer.toBuffer(doc).then(buffer => {
  const outputPath = '/home/orangels/xm_dev/ls_dev/reportSkillMaker/output/output_2025年12月统计报告_docx.docx';
  fs.writeFileSync(outputPath, buffer);
  console.log(`报告已生成: ${outputPath}`);
});

