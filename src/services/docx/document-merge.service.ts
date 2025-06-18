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
import { resolve } from "path";
import { SqlDatabase } from "langchain/sql_db";
import { DataSource, DataSourceOptions } from "typeorm";

// 取得各章節產出的document sections object

import { establishSubjectInfo } from "./subject-information.service.js";
import { establishCompanyProfile } from "./company-profile.service.js";
import { establishIncomeStatement } from "./income-statement.service.js";
import { establishBalanceSheet } from "./balance-sheet-analysis.service.js";
import { establishFinancialRatios } from "./financial-ratio-analysis.service.js";
import { establishCashFlow } from "./cash-flow-analysis.service.js";
import { establishCustomerTransaction } from "./curstomer-transaction.service.js";
import { establishNegativeNews } from "./negative-news-analysis.service.js";

// 完目前僅作項目資訊、公司基本資料、資產負債分析、財務比率分析
// TODO：損益分析、現金流量分析、產品及進銷對象分析、負面新聞分析

export const mergeAllChapters = async (
  year: number,
  gui_no: string,
  aiSummaryText: string,
  db: SqlDatabase
) => {
  try {
    const [
      subjectInfoSection,
      companyProfileSection,
      balanceSheetSection,
      financialRatiosSection,
      // profitLossSection,
      // cashFlowSection,
      // productTradingSection,
      // negativeNewsSection,
    ] = await Promise.all([
      establishSubjectInfo(gui_no, db),
      establishCompanyProfile(gui_no, db),
      establishBalanceSheet(year, gui_no, db),
      establishFinancialRatios(year, gui_no, aiSummaryText, db),

      // establishIncomeStatement(),
      // establishCashFlow(),
      // establishProductAndTrading(),
      // establishNegativeNews(),
    ]);

    const doc = new Document({
      sections: [
        subjectInfoSection,
        companyProfileSection,
        balanceSheetSection,
        financialRatiosSection,
      ],
    });
    const buffer = await Packer.toBuffer(doc);
    return buffer;
  } catch (error: any) {
    return {
      status: "error",
      error: error.message,
    };
  }
};
