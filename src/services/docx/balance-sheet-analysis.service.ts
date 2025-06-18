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

export async function establishBalanceSheet(
  year: number,
  gui_no: string,
  db: any
) {
  let balanceSheetMappedData: {
    Q1: TableCellDataType[];
    Q2: TableCellDataType[];
    Q3: TableCellDataType[];
    Q4: TableCellDataType[];
  } = {
    Q1: [],
    Q2: [],
    Q3: [],
    Q4: [],
  };
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
  const balanceSheetTableName = dbTableNameMapping["balance_sheet"];

  const balanceSheetSqlQueryStatement = `SELECT year,quarter,retained_earnings,other_accounts_receivable,other_current_assets,inventory,accounts_receivable,total_equity,current_liabilities,current_assets,cash,capital_stock,total_liabilities_and_equity,capital_reserve,total_assets,non_current_liabilities,non_current_assets FROM ${balanceSheetTableName} WHERE year = ${year} AND gui_no = ${gui_no} ORDER BY quarter ASC;`;
  const financialRatiosSqlQueryStatement = `SELECT year,gui_no,average_collection_period ,total_asset_turnover,roe DECIMAL,average_days_sales_outstanding,net_profit_margin,debt_to_asset_ratio,pre_tax_profit_to_capital_ratio,long_term_capital_to_fixed_assets_ratio,current_ratio,interest_coverage_ratio,roa,cash_reinvestment_ratio,cash_adequacy_ratio,quick_ratio,accounts_receivable_turnover,fixed_assets_turnover,inventory_turnover,cash_flow_ratio,eps FROM ${financialRatiosTableName} WHERE year =${year} AND gui_no = ${gui_no} ;`;

  // 執行資產負債分析的SQL查詢
  const balanceSheetResult = await executeQueryTool.invoke(
    balanceSheetSqlQueryStatement
  );
  const parsedBalanceSheetResult = JSON.parse(balanceSheetResult);

  parsedBalanceSheetResult.forEach((item) => {
    switch (item.quarter) {
      case 1:
        const Q1Data = Object.keys(item).map((it, key) => {
          // 將數字轉換成千分位格式，並轉成string格式
          const formattedNumber = item[it].toLocaleString();

          return {
            Title: mapTable.balanceSheetMap[it],
            Content: formattedNumber,
          };
        });
        balanceSheetMappedData.Q1 = Q1Data;
        break;
      case 2:
        const Q2Data = Object.keys(item).map((it, key) => {
          const formattedNumber = item[it].toLocaleString();

          return {
            Title: mapTable.balanceSheetMap[it],
            Content: formattedNumber,
          };
        });
        balanceSheetMappedData.Q2 = Q2Data;
        break;
      case 3:
        const Q3Data = Object.keys(item).map((it, key) => {
          const formattedNumber = item[it].toLocaleString();

          return {
            Title: mapTable.balanceSheetMap[it],
            Content: formattedNumber,
          };
        });
        balanceSheetMappedData.Q3 = Q3Data;
        break;
      case 4:
        const Q4Data = Object.keys(item).map((it, key) => {
          const formattedNumber = item[it].toLocaleString();

          return {
            Title: mapTable.balanceSheetMap[it],
            Content: formattedNumber,
          };
        });
        balanceSheetMappedData.Q4 = Q4Data;
        break;

      default:
        break;
    }
  });

  // 資產負債：整理出第一列表頭
  const BalanceSheetTitleRow = [
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
                  text: "數值",
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

  // 資產負債Q1：內容TABLE
  const BalanceSheetQ1ContentArray = balanceSheetMappedData.Q1.map((item) => {
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

  // 資產負債Q2：內容TABLE
  const BalanceSheetQ2ContentArray = balanceSheetMappedData.Q2.map((item) => {
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

  // 資產負債Q3：內容TABLE
  const BalanceSheetQ3ContentArray = balanceSheetMappedData.Q3.map((item) => {
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

  // 資產負債Q4：內容TABLE
  const BalanceSheetQ4ContentArray = balanceSheetMappedData.Q4.map((item) => {
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

  // 資產負債 Q1~Q4
  const BalanceSheetQ1FinalArray = BalanceSheetTitleRow.concat(
    BalanceSheetQ1ContentArray
  );
  const BalanceSheetQ2FinalArray = BalanceSheetTitleRow.concat(
    BalanceSheetQ2ContentArray
  );
  const BalanceSheetQ3FinalArray = BalanceSheetTitleRow.concat(
    BalanceSheetQ3ContentArray
  );
  const BalanceSheetQ4FinalArray = BalanceSheetTitleRow.concat(
    BalanceSheetQ4ContentArray
  );

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
            text: "資產負債分析",
            size: 36,
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
            text: "2024年第1季",
            size: 24,
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
        rows: BalanceSheetQ1FinalArray,
      }),
      new Paragraph({ children: [new TextRun({ break: 2 })] }),

      new Paragraph({
        spacing: {
          before: 100,
          after: 100,
        },
        children: [
          new TextRun({
            text: "2024年第2季",
            size: 24,
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
        rows: BalanceSheetQ2FinalArray,
      }),
      new Paragraph({ children: [new TextRun({ break: 2 })] }),
      new Paragraph({
        spacing: {
          before: 100,
          after: 100,
        },
        children: [
          new TextRun({
            text: "2024年第3季",
            size: 24,
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
        rows: BalanceSheetQ3FinalArray,
      }),
      new Paragraph({ children: [new TextRun({ break: 2 })] }),
      new Paragraph({
        spacing: {
          before: 100,
          after: 100,
        },
        children: [
          new TextRun({
            text: "2024年第4季",
            size: 24,
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
        rows: BalanceSheetQ4FinalArray,
      }),
      new Paragraph({ children: [new TextRun({ break: 2 })] }),
    ],
  };
}
