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
import { mapTable } from "../mappings/tableMapping.js";
import { resolve } from "path";

// 定義TableCellDataType的interface
interface TableCellDataType {
  Title: string;
  Content: string;
}

export const excuteNewDoc = async (
  year: number,
  gui_no: string,
  aiSummaryText: string,
  db: any
) => {
  const dbTableNameMapping = {
    company_profile: "company_profile",
    income_statement: "income_statement",
    balance_sheet: "balance_sheet",
    financial_ratios: "financial_ratios",
    cash_flow_statement: "cash_flow_statement",
    customer_transaction: "customer_transaction",
  };

  // Varaible
  let companyNameTitle = "";
  // 從DB撈公司的company name來組成項目資訊
  // 從DB撈公司的基本資料來組成Content
  const companyProfileTableName = dbTableNameMapping["company_profile"];
  const subjectSqlQueryStatement = `SELECT full_name_zhtw FROM ${companyProfileTableName} WHERE gui_no = ${gui_no};`;
  const companyProfileSqlQueryStatement = `SELECT stock_code,full_name_zhtw,short_name_zhtw,gui_no,address_zhtw,phone,fax,website,email,industry_main,industry_sub,industry_national,ceo,capital,employee_count,founded_date,business_scope,accountant_firm,accountants,board_shareholding_ratio,board_pledge_ratio,listed_market,par_value,ipo_date,avg_60d_price,avg_60d_volume FROM ${companyProfileTableName} WHERE gui_no = ${gui_no};`;
  const companyEnProfileSqlQueryStatement = `SELECT full_name_enus,short_name_enus,address_enus FROM ${companyProfileTableName} WHERE gui_no = ${gui_no};`;
  const companyOtherProfileSqlQueryStatement = `SELECT management_team,registration_change_record,investment_projects FROM ${companyProfileTableName} WHERE gui_no = ${gui_no};`;


  // 從DB撈資產負債分析、財務比率分析的資料來組成Content
  const financialRatiosTableName = dbTableNameMapping["financial_ratios"];
  const balanceSheetTableName = dbTableNameMapping["balance_sheet"];

  const balanceSheetSqlQueryStatement = `SELECT year,quarter,retained_earnings,other_accounts_receivable,other_current_assets,inventory,accounts_receivable,total_equity,current_liabilities,current_assets,cash,capital_stock,total_liabilities_and_equity,capital_reserve,total_assets,non_current_liabilities,non_current_assets FROM ${balanceSheetTableName} WHERE year = ${year} AND gui_no = ${gui_no} ORDER BY quarter ASC;`;
  const financialRatiosSqlQueryStatement = `SELECT year,gui_no,average_collection_period ,total_asset_turnover,roe DECIMAL,average_days_sales_outstanding,net_profit_margin,debt_to_asset_ratio,pre_tax_profit_to_capital_ratio,long_term_capital_to_fixed_assets_ratio,current_ratio,interest_coverage_ratio,roa,cash_reinvestment_ratio,cash_adequacy_ratio,quick_ratio,accounts_receivable_turnover,fixed_assets_turnover,inventory_turnover,cash_flow_ratio,eps FROM ${financialRatiosTableName} WHERE year =${year} AND gui_no = ${gui_no} ;`;

  // 用langchain的工具建立一個QuerySqlTool，用來執行SQL查詢
  const executeQueryTool = new QuerySqlTool(db);

  // 執行項目資訊的SQL查詢
  const subjectResult = await executeQueryTool.invoke(subjectSqlQueryStatement);

  const parsedSubjectResult = JSON.parse(subjectResult);
  const subjectMappedData = Object.keys(parsedSubjectResult[0]).map(
    (item, key) => {
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
    if (item === "full_name_zhtw") {
      companyNameTitle = `${parsedCompanyProfileResult[0][item]}授信報告`;
    }

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

  // 執行資產負債分析的SQL查詢
  const balanceSheetResult = await executeQueryTool.invoke(
    balanceSheetSqlQueryStatement
  );
  const parsedBalanceSheetResult = JSON.parse(balanceSheetResult);

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

  // 執行財務比率分析的SQL查詢
  const financialRatiosResult = await executeQueryTool.invoke(
    financialRatiosSqlQueryStatement
  );
  const parsedFinancialRatiosResult = JSON.parse(financialRatiosResult);
  const financialRatiosMappedData = Object.keys(
    parsedFinancialRatiosResult[0]
  ).map((item, key) => {
    return {
      Title: mapTable.financialRatiosMap[item],
      Content: parsedFinancialRatiosResult[0][item],
    };
  });

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

  // 組成項目資訊
  // 生成項目為固定的：基本資料、資產負債分析、財務比率分析
  const SubjectDummyArr = [
    { Title: "測試標的", Content: "subjectCompanyName" },
    { Title: "生成項目", Content: "基本資料、資產負債分析、財務比率分析" },
  ];

  // 公司基礎資訊：寫死資料
  const CompanyInformationDummyArr = [
    { Title: "股票代碼", Content: "4960" },
    { Title: "公司中文全稱", Content: "誠美材料科技" },
    { Title: "公司中文簡稱", Content: "誠美材" },
    { Title: "統一編號", Content: "27790294" },
    { Title: "公司地址(中)", Content: "台南市善化區木柵港西路13號" },
    { Title: "電話", Content: "06-5889988" },
    { Title: "傳真", Content: "06-5091010" },
    { Title: "網址", Content: "http://www.cmmt.com.tw" },
    { Title: "電子信箱", Content: "cmmt_ir@cmmt.com.tw" },
    { Title: "主要產業名稱(證交所)", Content: "M23C光電/IO" },
    { Title: "子產業名稱", Content: "M23C8C LCD 原料" },
    { Title: "主計處產業名稱", Content: "2649 其他光電材" },
    { Title: "負責人", Content: "宋妍儀" },
    { Title: "實收資本額", Content: "2,147,483,647" },
    { Title: "員工人数", Content: "1,096" },
    { Title: "設立日期", Content: "2005-05-17" },
    { Title: "公司主要業務", Content: "偏光板;其他;無" },
    { Title: "會計師事務所", Content: "資誠聯合會計師事務所" },
    { Title: "會計師", Content: "吳建志;廖阿甚;" },
    { Title: "董監持股比率", Content: "16.37" },
    { Title: "董監設質比率", Content: "1,096" },
    { Title: "上市別", Content: "TSE" },
    { Title: "面額", Content: "10" },
    { Title: "最近上市日", Content: "2011-10-24" },
    { Title: "60日均價", Content: "12.95" },
    { Title: "60日均量", Content: "2,040.04" },
  ];

  // 資產負債分析：寫死資料
  const BalanceSheetDummyObject = {
    Q1: [
      { Title: "保留盈餘合計", Content: "3,296,538" },
      { Title: "其他應收款淨額", Content: "120,771" },
      { Title: "其他流動資產", Content: "1,396" },
      { Title: "存貨", Content: "2,707,026" },
      { Title: "應收帳款淨額", Content: "2,125,401" },
      { Title: "權益總額", Content: "9,221,830" },
      { Title: "流動負債合計", Content: "5,172,635" },
      { Title: "流動資產合計", Content: "9,694,544" },
      { Title: "現金及約當現金", Content: "4,501,007" },
      { Title: "股本合計", Content: "5,721,849" },
      { Title: "負債及權益總計", Content: "14,556,915" },
      { Title: "資本公積合計", Content: "596,099" },
      { Title: "資產總額", Content: "14,556,915" },
      { Title: "非流動負債合計", Content: "162,450" },
      { Title: "非流動資產合計", Content: "4,862,371" },
    ],
    Q2: [
      { Title: "保留盈餘合計", Content: "3,289,446" },
      { Title: "其他應收款淨額", Content: "123,859" },
      { Title: "其他流動資產", Content: "2,057" },
      { Title: "存貨", Content: "2,949,717" },
      { Title: "應收帳款淨額", Content: "2,641,804" },
      { Title: "權益總額", Content: "9,165,266" },
      { Title: "流動負債合計", Content: "6,592,410" },
      { Title: "流動資產合計", Content: "11,139,678" },
      { Title: "現金及約當現金", Content: "1,712,287" },
      { Title: "股本合計", Content: "5,721,849" },
      { Title: "負債及權益總計", Content: "15,921,571" },
      { Title: "資本公積合計", Content: "587,292" },
      { Title: "資產總額", Content: "15,921,571" },
      { Title: "非流動負債合計", Content: "163,895" },
      { Title: "非流動資產合計", Content: "4,781,893" },
    ],
    Q3: [
      { Title: "保留盈餘合計", Content: "3,056,985" },
      { Title: "其他應收款淨額", Content: "157,254" },
      { Title: "其他流動資產", Content: "1,626" },
      { Title: "存貨", Content: "2,984,527" },
      { Title: "應收帳款淨額", Content: "2,354,588" },
      { Title: "權益總額", Content: "8,957,240" },
      { Title: "流動負債合計", Content: "5,413,096" },
      { Title: "流動資產合計", Content: "9,808,265" },
      { Title: "現金及約當現金", Content: "1,548,317" },
      { Title: "股本合計", Content: "5,719,588" },
      { Title: "負債及權益總計", Content: "14,536,480" },
      { Title: "資本公積合計", Content: "588,871" },
      { Title: "資產總額", Content: "14,536,480" },
      { Title: "非流動負債合計", Content: "166,144" },
      { Title: "非流動資產合計", Content: "4,728,215" },
    ],
    Q4: [
      { Title: "保留盈餘合計", Content: "3,026,197" },
      { Title: "其他應收款淨額", Content: "159,309" },
      { Title: "其他流動資產", Content: "1,238" },
      { Title: "存貨", Content: "2,835,310" },
      { Title: "應收帳款淨額", Content: "2,355,214" },
      { Title: "權益總額", Content: "8,905,215" },
      { Title: "流動負債合計", Content: "4,096,808" },
      { Title: "流動資產合計", Content: "9,471,137" },
      { Title: "現金及約當現金", Content: "1,815,774" },
      { Title: "股本合計", Content: "5,717,054" },
      { Title: "負債及權益總計", Content: "14,416,236" },
      { Title: "資本公積合計", Content: "591,406" },
      { Title: "資產總額", Content: "14,416,236" },
      { Title: "非流動負債合計", Content: "1,414,213" },
      { Title: "非流動資產合計", Content: "4,945,099" },
    ],
  };

  // 財務比率分析：寫死資料
  const FinancialRatiosDummyArr = [
    { Title: "平均收現日數", Content: "97.33" },
    { Title: "總資產週轉率(次)", Content: "0.62" },
    { Title: "權益報酬率(%)", Content: "-2.95" },
    { Title: "平均銷貨日數", Content: "134.19" },
    { Title: "純益率(%)", Content: "-2.99" },
    { Title: "負債佔資產比率(%)", Content: "38.23" },
    { Title: "稅前純益佔實收資本比率(%)", Content: "-4.15" },
    { Title: "長期資金佔不動產、廠房及設備比率(%)", Content: "251.16" },
    { Title: "流動比率(%)", Content: "231.18" },
    { Title: "利息保障倍數(%)", Content: "-70.32" },
    { Title: "資產報酬率(%)", Content: "-1.09" },
    { Title: "現金再投資比率(%)", Content: "-1.25" },
    { Title: "現金流量允當比率(%)", Content: "19.92" },
    { Title: "速動比率(%)", Content: "155.7" },
    { Title: "應收款項週轉率(次)", Content: "3.75" },
    { Title: "不動產、廠房及設備週轉率(次)", Content: "2.22" },
    { Title: "存貨週轉率(次)", Content: "2.72" },
  ];

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

  // 公司基本資料:把已經組好的表頭row和內容的rows組合起來
  const CompanyInformationFinalArray = CompanyInformationRow.concat(
    CompanyInformationContentArray
  );

  const CompanyEnInformationFinalArray = CompanyEnInformationRow.concat(
    CompanyEnInformationContentArray
  );
  // 項目資訊
  const SubjectFinalArray = SubjectTitleRow.concat(SubjectContentArray);
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
  // 財務比率
  const FinancialRatiosFinalArray = FinancialRatiosTitleRow.concat(
    FinancialRatiosContentArray
  );

  // 一併放到table的rows底下=============================================
  const doc = new Document({
    sections: [
      {
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
              new TextRun({ text: "經營團隊", size: 26, color: "000000", bold: true,  }),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: parsedCompanyOtherProfileResult[0].management_team, size: 24, color: "000000" }),
            ],
          }),
          new Paragraph({ children: [new TextRun({ break: 1 })] }),
          new Paragraph({
            children: [
              new TextRun({ text: "公司變更", size: 26, color: "000000", bold: true, }),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: parsedCompanyOtherProfileResult[0].registration_change_record, size: 24, color: "000000" }),
            ],
          }),
          new Paragraph({ children: [new TextRun({ break: 1 })] }),
          new Paragraph({
            children: [
              new TextRun({ text: "投資項目", size: 26, color: "000000", bold: true,  }),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: parsedCompanyOtherProfileResult[0].investment_projects, size: 24, color: "000000" }),
            ],
          }),
          new Paragraph({ children: [new TextRun({ break: 1 })] }),

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
      },
    ],
  });

  // 重複呼叫Packer.toBuffer會造成doc變動，造成後續buffer產出的文件會損毀，要注意
  // await Packer.toBuffer(doc).then((buffer) => {
  //   fs.writeFileSync("My Document.docx", buffer);
  // });

  const buffer = await Packer.toBuffer(doc);
  return buffer;
};
