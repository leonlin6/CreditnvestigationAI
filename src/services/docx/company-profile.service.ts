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

export async function establishCompanyProfile(gui_no: string, db: any) {
  // Varaible
  // let companyNameTitle = "";
  const dbTableNameMapping = {
    company_profile: "company_profile",
    income_statement: "income_statement",
    balance_sheet: "balance_sheet",
    financial_ratios: "financial_ratios",
    cash_flow_statement: "cash_flow_statement",
    customer_transaction: "customer_transaction",
  };

  // 從DB撈公司的company name來組成項目資訊
  // 從DB撈公司的基本資料來組成Content
  const companyProfileTableName = dbTableNameMapping["company_profile"];

  // TODO: 要改用ORM的作法去執行SQL語法
  const companyProfileSqlQueryStatement = `SELECT stock_code,full_name_zhtw,short_name_zhtw,gui_no,address_zhtw,phone,fax,website,email,industry_main,industry_sub,industry_national,ceo,capital,employee_count,founded_date,business_scope,accountant_firm,accountants,board_shareholding_ratio,board_pledge_ratio,listed_market,par_value,ipo_date,avg_60d_price,avg_60d_volume FROM ${companyProfileTableName} WHERE gui_no = ${gui_no};`;
  const companyEnProfileSqlQueryStatement = `SELECT full_name_enus,short_name_enus,address_enus FROM ${companyProfileTableName} WHERE gui_no = ${gui_no};`;
  const companyOtherProfileSqlQueryStatement = `SELECT management_team,registration_change_record,investment_projects FROM ${companyProfileTableName} WHERE gui_no = ${gui_no};`;

  // 用langchain的工具建立一個QuerySqlTool，用來執行SQL查詢
  const executeQueryTool = new QuerySqlTool(db);

  // 執行公司基本資料的SQL查詢
  const companyProfileResult = await executeQueryTool.invoke(
    companyProfileSqlQueryStatement
  );
  const companyEnProfileResult = await executeQueryTool.invoke(
    companyEnProfileSqlQueryStatement
  );

  const companyOtherProfileResult = await executeQueryTool.invoke(
    companyOtherProfileSqlQueryStatement
  );

  const parsedCompanyProfileResult = JSON.parse(companyProfileResult);
  const parsedCompanyEnProfileResult = JSON.parse(companyEnProfileResult);
  const parsedCompanyOtherProfileResult = JSON.parse(companyOtherProfileResult);

  const companyProfileMappedData = Object.keys(
    parsedCompanyProfileResult[0]
  ).map((item, key) => {
    // if (item === "full_name_zhtw") {
    //   companyNameTitle = `${parsedCompanyProfileResult[0][item]}授信報告`;
    // }

    if (typeof parsedCompanyProfileResult[0][item] === "number") {
      return {
        Title: mapTable.companyProfileMap[item],
        Content: parsedCompanyProfileResult[0][item].toLocaleString(),
      };
    } else {
      return {
        Title: mapTable.companyProfileMap[item],
        Content: parsedCompanyProfileResult[0][item],
      };
    }
  });

  const companyEnProfileMappedData = Object.keys(
    parsedCompanyEnProfileResult[0]
  ).map((item, key) => {
    return {
      Title: mapTable.companyEnProfileMap[item],
      Content: parsedCompanyEnProfileResult[0][item],
    };
  });

  // 公司基本資料-基本資料表：這裡整理出第一列標頭
  const CompanyInformationRow = [
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
                  text: "中文資訊",
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

  // 公司基本資料-基本資料表-中文資訊：把內容物從ARRAY中塞入TABLE中，用mpas組成陣列
  const CompanyInformationContentArray = companyProfileMappedData.map(
    (item) => {
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
    }
  );

  // 公司基本資料-基本資料表：這裡整理出第一列標頭
  const CompanyEnInformationRow = [
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
                  text: "英文資訊",
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

  // 公司基本資料-基本資料表-英文資訊：之後補
  const CompanyEnInformationContentArray = companyEnProfileMappedData.map(
    (item) => {
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
    }
  );

  // 公司基本資料:把已經組好的表頭row和內容的rows組合起來
  const CompanyInformationFinalArray = CompanyInformationRow.concat(
    CompanyInformationContentArray
  );

  const CompanyEnInformationFinalArray = CompanyEnInformationRow.concat(
    CompanyEnInformationContentArray
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
            text: "公司基本資料",
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
            text: "基本資料表",
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
        rows: CompanyInformationFinalArray,
      }),
      new Paragraph({ children: [new TextRun({ break: 2 })] }),
      new Table({
        layout: TableLayoutType.FIXED,
        width: {
          size: 8000,
          type: "dxa",
        },
        rows: CompanyEnInformationFinalArray,
      }),
      new Paragraph({ children: [new TextRun({ break: 2 })] }),
      new Paragraph({
        children: [
          new TextRun({
            text: "經營團隊",
            size: 26,
            color: "000000",
            bold: true,
          }),
        ],
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: parsedCompanyOtherProfileResult[0].management_team,
            size: 24,
            color: "000000",
          }),
        ],
      }),
      new Paragraph({ children: [new TextRun({ break: 1 })] }),
      new Paragraph({
        children: [
          new TextRun({
            text: "公司變更",
            size: 26,
            color: "000000",
            bold: true,
          }),
        ],
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: parsedCompanyOtherProfileResult[0].registration_change_record,
            size: 24,
            color: "000000",
          }),
        ],
      }),
      new Paragraph({ children: [new TextRun({ break: 1 })] }),
      new Paragraph({
        children: [
          new TextRun({
            text: "投資項目",
            size: 26,
            color: "000000",
            bold: true,
          }),
        ],
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: parsedCompanyOtherProfileResult[0].investment_projects,
            size: 24,
            color: "000000",
          }),
        ],
      }),
      new Paragraph({ children: [new TextRun({ break: 1 })] }),
    ],
  };
}
