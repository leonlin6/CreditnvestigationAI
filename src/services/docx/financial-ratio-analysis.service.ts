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

// 定義TableCellDataType的interface
interface TableCellDataType {
  Title: string;
  Content: string;
}

export async function establishFinancialRatios(
  year: number,
  gui_no: string,
  aiSummaryText: string,
  db: any
) {
  // 用langchain的工具建立一個QuerySqlTool，用來執行SQL查詢
  const executeQueryTool = new QuerySqlTool(db);

  const dbTableNameMapping = {
    company_profile: "company_profile",
    income_statement: "income_statement",
    balance_sheet: "balance_sheet",
    financial_ratios: "financial_ratios",
    cash_flow_statement: "cash_flow_statement",
    customer_transaction: "customer_transaction",
  };
  // 從DB撈資產負債分析、財務比率分析的資料來組成Content
  const financialRatiosTableName = dbTableNameMapping["financial_ratios"];

  const financialRatiosSqlQueryStatement = `SELECT year,gui_no,average_collection_period ,total_asset_turnover,roe DECIMAL,average_days_sales_outstanding,net_profit_margin,debt_to_asset_ratio,pre_tax_profit_to_capital_ratio,long_term_capital_to_fixed_assets_ratio,current_ratio,interest_coverage_ratio,roa,cash_reinvestment_ratio,cash_adequacy_ratio,quick_ratio,accounts_receivable_turnover,fixed_assets_turnover,inventory_turnover,cash_flow_ratio,eps FROM ${financialRatiosTableName} WHERE year =${year} AND gui_no = ${gui_no} ;`;

  // 執行財務比率分析的SQL查詢
  const financialRatiosResult = await executeQueryTool.invoke(
    financialRatiosSqlQueryStatement
  );
  const parsedFinancialRatiosResult = JSON.parse(financialRatiosResult);

  // 財務比率: 整理出第一列表頭
  const FinancialRatiosTitleRow = [
    new TableRow({
      children: [
        new TableCell({
          width: {
            size: 5400,
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
                  text: "會計科目",
                  size: 26,
                }),
              ],
            }),
          ],
          verticalAlign: VerticalAlignTable.CENTER,
        }),
        new TableCell({
          width: {
            size: 3600,
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
              alignment: "right",
              children: [
                new TextRun({
                  text: "2024年",
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

  const financialRatiosMappedData = Object.keys(
    parsedFinancialRatiosResult[0]
  ).map((item, key) => {
    return {
      Title: mapTable.financialRatiosMap[item],
      Content: parsedFinancialRatiosResult[0][item],
    };
  });

  // 財務比率: 內容TABLE
  const FinancialRatiosContentArray = financialRatiosMappedData.map((item) => {
    const content = [
      new TableCell({
        width: { size: 5400 },
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
        width: { size: 3600 },
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
            alignment: "right",
            children: [
              new TextRun({
                text: JSON.stringify(item.Content),
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

  // 財務比率
  const FinancialRatiosFinalArray = FinancialRatiosTitleRow.concat(
    FinancialRatiosContentArray
  );

  // 整理傳來的AI資料(string)，要重新組成docx的物件，加入new TextRun的break，才有辦法換行
  // 目前只有一筆AI分析：(資產負債分析＋財務比率分析)
  const refactoredSummaryChildren = aiSummaryText
    .split("\n")
    .flatMap((line, index, arr) => {
      const runs = [
        new TextRun({
          text: line,
          size: 24,
        }),
      ];
      if (index < arr.length - 1) {
        runs.push(new TextRun({ break: 1 })); // Insert a line break
      }
      return runs;
    });

  // TODO: Implement company basic information analysis logic
  return {
    children: [
      new Paragraph({
        spacing: {
          before: 100,
          after: 100,
        },
        children: [
          new TextRun({
            text: "財稅比率分析",
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
        rows: FinancialRatiosFinalArray,
      }),
      new Paragraph({ children: [new TextRun({ break: 2 })] }),
      new Paragraph({
        children: refactoredSummaryChildren,
      }),
    ],
  };
}
