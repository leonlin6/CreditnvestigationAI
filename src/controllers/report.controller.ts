import { QuerySqlTool } from "langchain/tools/sql";
import { SqlDatabase } from "langchain/sql_db";
import { establishReport } from "../services/langchain.service.js";
import { mergeAllChapters } from "../services/docx/document-merge.service.js";

const getReport = async (
  db: SqlDatabase,
  year: number,
  companyName: string
) => {
  const sqlQueryStatement = `SELECT gui_no FROM company_profile WHERE full_name_zhtw = '${companyName}';`;
  const executeQueryTool = new QuerySqlTool(db);

  const result = await executeQueryTool.invoke(sqlQueryStatement);
  const parsedResult = JSON.parse(result);

  if (!parsedResult || parsedResult.length === 0) {
    throw new Error("Company not found");
  }

  const gui_no = parsedResult[0].gui_no;
  const aiSummaryText = await establishReport(Number(year), gui_no, db).catch(
    (err) => {
      console.log(err);
      return err.message;
    }
  );

  // return await excuteNewDoc(Number(year), gui_no, report, db);
  return await mergeAllChapters(Number(year), gui_no, aiSummaryText, db);
};

const getCompanyNames = async (db: SqlDatabase) => {
  const sqlQueryStatement = `SELECT full_name_zhtw FROM company_profile ;`;
  const executeQueryTool = new QuerySqlTool(db);
  const result = await executeQueryTool.invoke(sqlQueryStatement);

  const parsedResult = JSON.parse(result);

  return parsedResult.map((item: { full_name_zhtw: string }) => {
    return item.full_name_zhtw;
  });
};

export { getReport, getCompanyNames };
