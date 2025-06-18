// Example of how you would create a table and add data to it

import * as fs from "fs";
import {
  Document,
  HeadingLevel,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  VerticalAlignTable,
  TextDirection,
  TableLayoutType,
  TextRun,
  WidthType,
  BorderStyle,
  ShadingType,
  AlignmentType,
} from "docx";
import { QuerySqlTool } from "langchain/tools/sql";
import { mapTable } from "../../mappings/tableMapping.js";

export async function establishSubjectInfo(gui_no: string, db: any) {
  // Varaible
  let companyNameTitle = "";
  const dbTableNameMapping = {
    company_profile: "company_profile",
    income_statement: "income_statement",
    balance_sheet: "balance_sheet",
    financial_ratios: "financial_ratios",
    cash_flow_statement: "cash_flow_statement",
    customer_transaction: "customer_transaction",
  };

  const companyProfileTableName = dbTableNameMapping["company_profile"];

  // TODO: 要改用ORM的作法去執行SQL語法
  const subjectSqlQueryStatement = `SELECT full_name_zhtw FROM ${companyProfileTableName} WHERE gui_no = ${gui_no};`;

  // 用langchain的工具建立一個QuerySqlTool，用來執行SQL查詢
  const executeQueryTool = new QuerySqlTool(db);
  // 執行項目資訊的SQL查詢
  const subjectResult = await executeQueryTool.invoke(subjectSqlQueryStatement);
  const parsedSubjectResult = JSON.parse(subjectResult);
  const subjectMappedData = Object.keys(parsedSubjectResult[0]).map(
    (item, key) => {
      if (item === "full_name_zhtw") {
        companyNameTitle = `${parsedSubjectResult[0][item]}授信報告`;
      }

      return {
        Title: mapTable.subjectMap[item],
        Content: parsedSubjectResult[0][item],
      };
    }
  );
  //手動寫死push生成項目進來
  subjectMappedData.push({
    Title: "生成項目",
    Content: "基本資料、資產負債分析、財務比率分析",
  });

  // 項目資訊：整理出第一列表頭
  const SubjectTitleRow = [
    new TableRow({
      children: [
        new TableCell({
          width: {
            size: 2000,
            type: "dxa",
          },
          borders: {
            top: {
              style: BorderStyle.SINGLE,
              size: 14,
              color: "000000",
            },
            bottom: {
              style: BorderStyle.SINGLE,
              size: 14,
              color: "000000",
            },
            left: {
              style: BorderStyle.SINGLE,
              size: 14,
              color: "000000",
            },
            right: {
              style: BorderStyle.SINGLE,
              size: 14,
              color: "000000",
            },
          },
          children: [
            new Paragraph({
              spacing: {
                before: 100,
                after: 100,
              },
              indent: {
                left: 200,
                right: 200,
              },
              children: [
                new TextRun({
                  text: "項目",
                  size: 26,
                }),
              ],
            }),
          ],
          verticalAlign: VerticalAlignTable.CENTER,
        }),
        new TableCell({
          width: {
            size: 7000,
            type: "dxa",
          },
          borders: {
            top: {
              style: BorderStyle.SINGLE,
              size: 14,
              color: "000000",
            },
            bottom: {
              style: BorderStyle.SINGLE,
              size: 14,
              color: "000000",
            },
            left: {
              style: BorderStyle.SINGLE,
              size: 14,
              color: "000000",
            },
            right: {
              style: BorderStyle.SINGLE,
              size: 14,
              color: "000000",
            },
          },
          children: [
            new Paragraph({
              spacing: {
                before: 100,
                after: 100,
              },
              indent: {
                left: 200,
                right: 200,
              },
              children: [
                new TextRun({
                  text: "資訊",
                  size: 26,
                }),
              ],
            }),
          ],

          verticalAlign: VerticalAlignTable.CENTER,
        }),
      ],
    }),
  ];

  // 項目資訊: 內容TABLE
  const SubjectContentArray = subjectMappedData.map((item) => {
    const content = [
      new TableCell({
        width: { size: 2000 },
        borders: {
          top: {
            style: BorderStyle.SINGLE,
            size: 14,
            color: "000000",
          },
          bottom: {
            style: BorderStyle.SINGLE,
            size: 14,
            color: "000000",
          },
          left: {
            style: BorderStyle.SINGLE,
            size: 14,
            color: "000000",
          },
          right: {
            style: BorderStyle.SINGLE,
            size: 14,
            color: "000000",
          },
        },
        children: [
          new Paragraph({
            spacing: {
              before: 100,
              after: 100,
            },
            indent: {
              left: 200,
              right: 200,
            },
            children: [
              new TextRun({
                text: item.Title,
                size: 24,
              }),
            ],
          }),
        ],
      }),
      new TableCell({
        width: { size: 7000 },
        borders: {
          top: {
            style: BorderStyle.SINGLE,
            size: 14,
            color: "000000",
          },
          bottom: {
            style: BorderStyle.SINGLE,
            size: 14,
            color: "000000",
          },
          left: {
            style: BorderStyle.SINGLE,
            size: 14,
            color: "000000",
          },
          right: {
            style: BorderStyle.SINGLE,
            size: 14,
            color: "000000",
          },
        },
        children: [
          new Paragraph({
            spacing: {
              before: 100,
              after: 100,
            },
            indent: {
              left: 200,
              right: 200,
            },
            children: [
              new TextRun({
                text: item.Content,
                size: 24,
              }),
            ],
          }),
        ],
      }),
    ];

    return new TableRow({
      children: content,
    });
  });

  // 項目資訊
  const SubjectFinalArray = SubjectTitleRow.concat(SubjectContentArray);

  // TODO: Implement company basic information analysis logic
  return {
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({
            text: companyNameTitle,
            size: 48,
            color: "000000",
          }),
        ],
      }),
      new Paragraph({
        spacing: {
          before: 100,
          after: 100,
        },
        children: [
          new TextRun({
            text: "項目資訊",
            size: 36,
            color: "000000",
          }),
        ],
      }),
      new Table({
        layout: TableLayoutType.FIXED,
        width: {
          size: 8000,
          type: "dxa",
        },
        rows: SubjectFinalArray,
      }),
      new Paragraph({ children: [new TextRun({ break: 2 })] }),
    ],
  };
}
