import "reflect-metadata";
// ES Modules imports
import { ChatOpenAI } from "@langchain/openai";
import { ChatPromptTemplate } from "@langchain/core/prompts";
import { RunnableSequence } from "@langchain/core/runnables";
import { StringOutputParser } from "@langchain/core/output_parsers";
import dotenv from "dotenv";
import readline from "readline";
import { createServer } from "http";
import { ChatOllama } from "@langchain/ollama";

import { Annotation, StateGraph } from "@langchain/langgraph";
import { SqlDatabase } from "langchain/sql_db";
import { DataSource } from "typeorm";
import { pull } from "langchain/hub";
import { z } from "zod";
import { QuerySqlTool } from "langchain/tools/sql";
import sqlite3 from "sqlite3";
import fs from "fs";
import { $ } from "execa";
// import { excuteNewDoc } from './TableTransfer.ts';
import EvaluationRulePrompt from "./EvaluationRulePrompt.js";

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
// Declare
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

// const chapters = [
//   'income_statement',
//   'balance_sheet',
//   'financial_ratios',
//   'cash_flow_statement',
//   'customer_transaction'
// ];
const chapters = ["financial_ratios"];
// Initialize the OpenAI model
const chatOpenAImodel = new ChatOpenAI({
  modelName: "gpt-4o-mini",
  temperature: 0.3,
  openAIApiKey: process.env.OPENAI_API_KEY,
});

// const chatOpenAImodel = new ChatOllama({
//   model: "llama3.2:latest",
//   temperature: 0,
//   // other params...
// });

// HTTP Server handler
// function handler(req, res) {
//   console.log(req.url); // Log the request URL
//   res.write("Hello World!"); // Set response content

//   res.end(); // End the response
// }

// // Create HTTP server
// const server = createServer(handler);
// // Start HTTP server
// server.listen(5001, () => {
//   console.log("Server running at http://localhost:5001/");
// });

// declare db configure
const creditDatasource = new DataSource({
  type: "sqlite",
  database: "credit_check.db",
});

// declare db
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
const sqlStructuredLlm = chatOpenAImodel.withStructuredOutput(sqpQueryOutput);

// condtion edge function
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
// 不知道為什麼<ChatPromptTemplate>不能加在泛型裡面，會造成return false，拿掉就可以正常使用
const queryPromptTemplate = await pull("langchain-ai/sql-query-system-prompt");

// node function： 產生SQL語法
const generate_sql = async (state: typeof GenerateReportAnnotation.State) => {
  // const year = 2024;
  // const gui_no = '27790294';
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

  const summary = await chatOpenAImodel.invoke(generatingSummaryPrompt);

  const transferToTCPrompt = `Transfer the following topice into traditional chinese. ${summary.content}`;

  const summaryTC = await chatOpenAImodel.invoke(transferToTCPrompt);
  const summaryWithTitle = `基於資產負債表得出結論：\n` + summaryTC.content;
  return {
    ...state,
    summary: summaryWithTitle,
  };
};

// financial_ratios node
// 根據不同chapter，來拿取不同的欄位資料，以編寫不同的prompt來產出對應的LLM分析結果
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

  const summaryResponse = await chatOpenAImodel.invoke(generatingAnswerPrompt);
  return {
    ...state,
    answer: summaryResponse.content,
  };
};

// Compile application and test
const graph = new StateGraph(GenerateReportAnnotation)

  .addNode("generate_sql", generate_sql)
  .addNode("obtain_data", obtain_data)
  // .addNode("income_statement", income_statement)
  // .addNode("balance_sheet", balance_sheet)
  .addNode("financial_ratios", financial_ratios)
  // .addNode("cash_flow_statement", cash_flow_statement)
  // .addNode("customer_transaction", customer_transaction)
  .addNode("import_into_container", import_into_container)

  .addEdge("__start__", "generate_sql")
  .addEdge("generate_sql", "obtain_data")
  .addConditionalEdges("obtain_data", distiguishType)
  // .addEdge("income_statement", "import_into_container")
  // .addEdge("balance_sheet", "import_into_container")
  .addEdge("financial_ratios", "import_into_container")
  // .addEdge("cash_flow_statement", "import_into_container")
  // .addEdge("customer_transaction", "import_into_container")
  .addEdge("import_into_container", "__end__")
  .compile();

// Load environment variables
dotenv.config();

// 一鍵生成報告，依chapter跑迴圈執行，產出每個章節的數據表格與AI建議文字
export async function establishReport(year: number, gui_no: string) {
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
      // console.log(`establishReport result========`,result);

      resolve(result.summary);
    });

    try {

      return userInputResponse;
    } catch (error) {
      console.error("An error occurred:");
    }
  }
}

import { excuteNewDoc } from "./TableTransfer.js";
// const response = await establishReport(2024, "27790294").catch((err) => {
//   console.log(err);
//   return err.message;
// });

// await excuteNewDoc(response);
// await excuteNewDoc(2024, "27790294", response, db);

// express API Server handler
import express from 'express';
import cors from 'cors';
const app = express()
const port = 3000

const corsOptions = {
  origin: [
    'https://web-creditnvestigation-70fhvfear-leonlin6s-projects-5cda9579',
    'https://web-creditnvestigation-70fhvfear-leonlin6s-projects-5cda9579.vercel.app',
  ],
  methods: 'GET,HEAD,PUT,PATCH,POST,DELETE,OPTIONS',
  allowedHeaders: ['Content-Type', 'Authorization'],
};

app.use(cors(corsOptions));
app.use(express.json());

app.get('/api/report', async (req, res) => {
  const { year, companyName } = req.query;
  
  if (!year || !companyName) {
    return res.status(400).json({ error: 'Year and company name are required' });
  }

  try {
    const sqlQueryStatement = `SELECT gui_no FROM company_profile WHERE full_name_zhtw = '${companyName}';`;
    const executeQueryTool = new QuerySqlTool(db);
    const result = await executeQueryTool.invoke(sqlQueryStatement);
    const parsedResult = JSON.parse(result);
    
    if (!parsedResult || parsedResult.length === 0) {
      return res.status(404).json({ error: 'Company not found' });
    }

    const gui_no = parsedResult[0].gui_no;
    const report = await establishReport(Number(year), gui_no).catch((err) => {
      console.log(err);
      return err.message;
    });
    
    const buffer = await excuteNewDoc(Number(year), gui_no, report, db);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment;');
  
    // 傳送 buffer
    res.send(buffer);
  } catch (error) {
    console.error('Error processing request:', error);
    res.status(500).json({ error: 'Internal server error' });
  }
})


app.get('/api/companies/name', async (req, res) => {
  const sqlQueryStatement = `SELECT full_name_zhtw FROM company_profile ;`;
  const executeQueryTool = new QuerySqlTool(db);
  const result = await executeQueryTool.invoke(sqlQueryStatement);

  const parsedResult = JSON.parse(result);

  const companyNameList = parsedResult.map((item: { full_name_zhtw: string }) => {
    return item.full_name_zhtw;
  });
  
  res.send(companyNameList);
})

app.listen(port, () => {
  console.log(`Example app listening on port ${port}`)
})
