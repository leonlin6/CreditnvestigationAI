import { ChatOllama } from "@langchain/ollama";
import { ChatOpenAI } from "@langchain/openai";
import { SqlDatabase } from "langchain/sql_db";
import { QuerySqlTool } from "langchain/tools/sql";
import { Annotation, StateGraph } from "@langchain/langgraph";
import { ChatPromptTemplate } from "@langchain/core/prompts";
import { StringOutputParser } from "@langchain/core/output_parsers";
import { pull } from "langchain/hub";
import { z } from "zod";
import EvaluationRulePrompt from "../prompts/EvaluationRulePrompt.js";
import { DataSource, DataSourceOptions } from "typeorm";

interface SummaryContainerObject {
  income_statement: string;
  balance_sheet: string;
  financial_ratios: string;
  cash_flow_statement: string;
  customer_transaction: string;
}

interface DBTableNameObject {
  [key: string]: string;
  income_statement: string;
  balance_sheet: string;
  financial_ratios: string;
  cash_flow_statement: string;
  customer_transaction: string;
}

const dbTableNameMapping: DBTableNameObject = {
  income_statement: "income_statement",
  balance_sheet: "balance_sheet",
  financial_ratios: "financial_ratios",
  cash_flow_statement: "cash_flow_statement",
  customer_transaction: "customer_transaction",
};

const GenerateReportAnnotation = Annotation.Root({
  year: Annotation<number>,
  gui_no: Annotation<string>,
  question: Annotation<string>,
  chapter: Annotation<string>,
  query: Annotation<string>,
  result: Annotation<string>,
  answer: Annotation<string>,
  summary: Annotation<string>,
  container: Annotation<SummaryContainerObject>,
});

const chapters = ["financial_ratios"];

// Initialize the LangChain model
// const chatModel = new ChatOllama({
//   baseUrl: process.env.NGROK_URL,
//   model: "llama4:latest",
//   temperature: 0.3,
// });

const chatModel = new ChatOpenAI({
  modelName: "gpt-4o-mini",
  temperature: 0.3,
  openAIApiKey: process.env.OPENAI_API_KEY,
});

// set the sql configure
const dbConfig: DataSourceOptions = {
  type: "sqlite",
  database: process.env.SQLITE_DB || "credit_check.db",
};
const creditDatasource = new DataSource(dbConfig);
const db = await SqlDatabase.fromDataSourceParams({
  appDataSource: creditDatasource,
});

// declare queryOutput format by using zod
const sqpQueryOutput = z.object({
  query: z.string().describe("Syntactically valid SQL query."),
});

const commentObjectQueryOutput = z.object({
  comment: z.string().describe("Comment content for different comment."),
});

// declare llm which using withStructuredOutput : sql format
const sqlStructuredLlm = chatModel.withStructuredOutput(sqpQueryOutput);

// condition edge function
const distiguishType = (state: typeof GenerateReportAnnotation.State) => {
  switch (state.chapter) {
    case "income_statement":
      return "income_statement";
    case "balance_sheet":
      return "balance_sheet";
    case "financial_ratios":
      return "financial_ratios";
    case "cash_flow_statement":
      return "cash_flow_statement";
    case "customer_transaction":
      return "customer_transaction";
    default:
      return "__end__";
  }
};

// load the sql-query-system-prompt from prompt template of langchain hub
const queryPromptTemplate = await pull("langchain-ai/sql-query-system-prompt");

// node function： 產生SQL語法
const generate_sql = async (state: typeof GenerateReportAnnotation.State) => {
  const chapter = state.chapter;
  const dbTableName = dbTableNameMapping[chapter];
  const sqlQueryStatement = `SELECT * FROM ${dbTableName} WHERE year =${state.year} AND gui_no = ${state.gui_no};`;

  return {
    year: state.year,
    gui_no: state.gui_no,
    question: state.question,
    chapter: state.chapter,
    query: sqlQueryStatement,
    result: "",
    answer: "",
    summary: "",
  };
};

// node function： 執行SQL語法，取得DB中的資料
const obtain_data = async (state: typeof GenerateReportAnnotation.State) => {
  const executeQueryTool = new QuerySqlTool(db);
  const result = await executeQueryTool.invoke(state.query);
  return {
    ...state,
    result: result,
  };
};

// node function：將產出的文字內容，依object格式存起來，最後再一併產成docx
const import_into_container = async (
  state: typeof GenerateReportAnnotation.State
) => {
  const generatingSummaryPrompt =
    `Use the following information to generate a Comprehensive review withdata number. And Don't use markdown format as response format.` +
    `Information ${state.answer}` +
    `And then transfer the review from engilsh into traditional chinese.`;

  const summary = await chatModel.invoke(generatingSummaryPrompt);

  const transferToTCPrompt = `Transfer the following topice into traditional chinese. ${summary.content}`;

  const summaryTC = await chatModel.invoke(transferToTCPrompt);
  const summaryWithTitle = `基於資產負債表得出結論：\n` + summaryTC.content;
  return {
    ...state,
    summary: summaryWithTitle,
  };
};

// financial_ratios node
const financial_ratios = async (
  state: typeof GenerateReportAnnotation.State
) => {
  const parsedJson = JSON.parse(state.result);
  const str = JSON.stringify(parsedJson[0]);
  const generatingAnswerPrompt =
    "If LLM cannot find the related data in the datbase, responsing `Sorry, the question is Out of my service area.` else " +
    "answering the following user question, corresponding " +
    "the given SQL data. \n" +
    `Question: Please establish two comment. Comment A use source of debt_to_asset_ratio, current_ratio, quick_ratio, roa and roe in the ${str} and only following rules.` +
    `Adding % mark after numerical value on Comment A.` +
    `Comment B use source of accounts_receivable_turnover, nventory_turnover, total_asset_turnover in the ${str} and only following rules.` +
    `Don't Add % mark after numerical value on Comment B.` +
    `Rules:` +
    `debt_to_asset_ratio: ${EvaluationRulePrompt.debt_to_asset_ratio_rule}` +
    `current_ratio_rule: ${EvaluationRulePrompt.current_ratio_rule}` +
    `quick_ratio_rule: ${EvaluationRulePrompt.quick_ratio_rule}` +
    `roa_rule: ${EvaluationRulePrompt.roa_rule}` +
    `roe_rule: ${EvaluationRulePrompt.roe_rule}` +
    `accounts_receivable_turnover_rule: ${EvaluationRulePrompt.accounts_receivable_turnover_rule}` +
    `inventory_turnover_rule: ${EvaluationRulePrompt.inventory_turnover_rule}` +
    `total_asset_turnover_rule: ${EvaluationRulePrompt.total_asset_turnover_rule}`;

  const summaryResponse = await chatModel.invoke(generatingAnswerPrompt);
  return {
    ...state,
    answer: summaryResponse.content,
  };
};

// Compile application and test
const graph = new StateGraph(GenerateReportAnnotation)
  .addNode("generate_sql", generate_sql)
  .addNode("obtain_data", obtain_data)
  .addNode("financial_ratios", financial_ratios)
  .addNode("import_into_container", import_into_container)

  .addEdge("__start__", "generate_sql")
  .addEdge("generate_sql", "obtain_data")
  .addConditionalEdges("obtain_data", distiguishType)
  .addEdge("financial_ratios", "import_into_container")
  .addEdge("import_into_container", "__end__")
  .compile();

export const establishReport = async (
  year: number,
  gui_no: string,
  db: SqlDatabase
) => {
  for (let chapter of chapters) {
    const userInputResponse: string = await new Promise(async (resolve) => {
      const result = await graph.invoke({
        year: year,
        gui_no: gui_no,
        question:
          "Establish a suggestion of specific chapter with reference data by querying db.",
        chapter: chapter,
        query: "",
        result: "",
        answer: "",
        summary: "",
      });

      resolve(result.summary);
    });

    try {
      return userInputResponse;
    } catch (error) {
      console.error("An error occurred:", error);
      throw error;
    }
  }
};
