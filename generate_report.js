const { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel, BorderStyle } = require('docx');
const fs = require('fs');

// 创建文档
const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: "仿宋", size: 32 } // 16磅 = 32半磅
      }
    },
    paragraphStyles: [
      {
        id: "Normal",
        name: "Normal",
        run: { font: "仿宋", size: 32 },
        paragraph: {
          spacing: { line: 560, lineRule: "exact" }, // 28磅行距
          indent: { firstLine: 640 } // 2字符首行缩进
        }
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        margin: { top: 1440, right: 1800, bottom: 1440, left: 1800 }
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
            size: 110, // 55磅
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
            color: "FF0000"
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
            size: 44, // 22磅
            bold: true
          })
        ]
      }),

      // 一、整体情况
      new Paragraph({
        children: [
          new TextRun({
            text: "一、整体情况",
            font: "黑体",
            size: 32,
            bold: true
          })
        ],
        spacing: { before: 200, after: 200 }
      }),

      new Paragraph({
        children: [
          new TextRun({
            text: "2025年12月1日至31日我局共接报",
            font: "仿宋",
            size: 32
          }),
          new TextRun({
            text: "有效警情",
            font: "仿宋",
            size: 32,
            bold: true
          }),
