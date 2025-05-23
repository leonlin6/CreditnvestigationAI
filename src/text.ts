import "reflect-metadata";
// ES Modules imports
import { ChatOpenAI } from "@langchain/openai";
import { ChatPromptTemplate } from "@langchain/core/prompts";
import { RunnableSequence } from "@langchain/core/runnables";
import { StringOutputParser } from "@langchain/core/output_parsers";
import dotenv from "dotenv";
import readline from "readline";
import { createServer } from 'http';
import { ChatOllama } from '@langchain/ollama';

import { Annotation, StateGraph } from "@langchain/langgraph";
import { SqlDatabase } from "langchain/sql_db";
import { DataSource } from "typeorm";
import {pull} from "langchain/hub";
// import { writeQuery } from "./WriteQuery";
import { z } from "zod";
import { QuerySqlTool } from "langchain/tools/sql";
import sqlite3 from "sqlite3";
import fs from "fs";

import { $ } from "execa";
// import { excuteNewDoc } from './TableTransfer.ts';
import EvaluationRulePrompt from './EvaluationRulePrompt.ts';
// Declare
const GenerateReportAnnotation = Annotation.Root({
  question: Annotation<string>,
  chapter: Annotation<string>,
  query: Annotation<string>,
  result: Annotation<string>,
  answer: Annotation<string>,
  summary: Annotation<string>,
});

// const chapters = [
//   'income_statement',
//   'balance_sheet',
//   'financial_ratios',
//   'cash_flow_statement',
//   'customer_transaction'
// ];
const chapters = [
  'financial_ratios',
];
// Initialize the OpenAI model
const chatOpenAImodel = new ChatOpenAI({
    modelName: "gpt-3.5-turbo",
    temperature: 0.3,
    openAIApiKey: process.env.OPENAI_API_KEY
});
 
// HTTP Server handler
function handler(req, res) {
  console.log(req.url);  // Log the request URL
  res.write('Hello World!');   // Set response content

  res.end()  ;   // End the response
}
  
// Create HTTP server
const server = createServer(handler);
  // Start HTTP server
  server.listen(5001, () => {
  console.log('Server running at http://localhost:5001/');
});  

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
const commentObjectStructuredLlm = chatOpenAImodel.withStructuredOutput(commentObjectQueryOutput);


// condtion edge function
const conditionTest = (state) => {
  switch(state.chapter){
    case 'income_statement':
      return 'income_statement';
    case 'balance_sheet':
      return 'balance_sheet';
    case 'financial_ratios':
      return 'financial_ratios';      
    case 'cash_flow_statement':
      return 'cash_flow_statement';      
    case 'customer_transaction':
      return 'customer_transaction'; 
    default:
      return '__end__';
      break;                                               
  }
}


// load the sql-query-system-prompt from prompt template of langchain hub
// 不知道為什麼<ChatPromptTemplate>不能加在泛型裡面，會造成return false，拿掉就可以正常使用
const queryPromptTemplate = await pull(
  "langchain-ai/sql-query-system-prompt"
);

// node function： 產生SQL語法
const generate_sql = async (state: typeof GenerateReportAnnotation.State) => {
  const promptValue = await queryPromptTemplate.invoke({
    dialect: db.appDataSourceOptions.type,
    top_k: 10,
    table_info: await db.getTableInfo([state.chapter]),
    input: `Please get all data in 2024 with 'gui_no' is '27790294'`,
  });

  const result = await sqlStructuredLlm.invoke(promptValue);
console.log('sqlStructuredLlm=======',result);

  return {
      question: state.question,
      chapter: state.chapter,
      query: result.query,
      result: "",
      answer: "",
      summary:""
    };
}

// node function： 執行SQL語法，取得DB中的資料
const obtain_data = async  (state: typeof GenerateReportAnnotation.State) => {
  const executeQueryTool = new QuerySqlTool(db);
  const result = await executeQueryTool.invoke(state.query);
   
  return { 
    question: state.question,
    query: state.query,
    result: result,
    answer: "",
  };
};

// node function： 產生AI建議敘述文字
const generate_summary = async (state: typeof GenerateReportAnnotation.State) => {
  const parsedJson = JSON.parse(state.result);
  const str = JSON.stringify(parsedJson[0]);
  const generatingAnswerPrompt =
    // "If LLM cannot find the related data in the datbase, responsing `Sorry, the question is Out of my service area.` else " +
    // "Given the following user question, corresponding " +
    // " SQL result, answer the user question. \n" +
    "answer the user question. \n" +
    // `QUESTION:What are the value of debt_to_asset_ratio, current_ratio, quick_ratio by following data: ${str}` ;
    `Question: Please establish two comment. Comment A use source of debt_to_asset_ratio, current_ratio, quick_ratio, roa and roe in the ${str} and only following rules .` +
    `Comment B use source of accounts_receivable_turnover, nventory_turnover, total_asset_turnover in the ${str} and only following rules.` +
    `Rules:` +
    `debt_to_asset_ratio: ${EvaluationRulePrompt.debt_to_asset_ratio_rule}` +
    `current_ratio_rule: ${EvaluationRulePrompt.current_ratio_rule}` +
    `quick_ratio_rule: ${EvaluationRulePrompt.quick_ratio_rule}` +
    `roa_rule: ${EvaluationRulePrompt.roa_rule}` +
    `roe_rule: ${EvaluationRulePrompt.roe_rule}` +
    `accounts_receivable_turnover_rule: ${EvaluationRulePrompt.accounts_receivable_turnover_rule}` +
    `inventory_turnover_rule: ${EvaluationRulePrompt.inventory_turnover_rule}` +
    `total_asset_turnover_rule: ${EvaluationRulePrompt.total_asset_turnover_rule}`;
    

    // `If current_ratio is higher than 100, it means the company may have a short-term liquidity problem.` +
    // `If quick_ratio is higher than 100, it means the company may have a short-term liquidity problem.`;
  // const response = await llm.invoke(promptValue);

  const textResponse = await commentObjectStructuredLlm.invoke(generatingAnswerPrompt);
console.log('textResponse=======',textResponse);
  return {
    ...state, 
    answer: ''
  };

};

// node function：建立報告
const export_docx = async  (state: typeof GenerateReportAnnotation.State) => {
  const generatingSummaryPrompt = `Use the following information to generate a Comprehensive review without data number. ` +
  `Information ${state.answer}` + 
  `And then transfer the review from engilsh into traditional chinese.`;

  const summary = await chatOpenAImodel.invoke(generatingSummaryPrompt);

  const transferToTCPrompt = `Transfer the following topice into traditional chinese. ${summary.content}`

  const summaryTC = await chatOpenAImodel.invoke(transferToTCPrompt);
  const summaryWithTitle = `基於資產負債表得出結論：\n` + summaryTC.content
  return {
    ...state,
    summary: summaryWithTitle
  };
};



// 根據chapter分到的node，要在這裡作
// 1.根據chapter，讀取不同的TABLE
// 2.根據chapter，抓不同的欄位資料
// 3.根據chpater，
const financial_ratios = (state: typeof GenerateReportAnnotation.State) => {
  return ({
    chapter: state.chapter,
    summary: ''
  })
}

// Compile application and test
const graph = new StateGraph(GenerateReportAnnotation)
  // .addNode("income_statement", income_statement)
  // .addNode("balance_sheet", balance_sheet)
  .addNode("financial_ratios", financial_ratios)
  // .addNode("cash_flow_statement", cash_flow_statement)
  // .addNode("customer_transaction", customer_transaction)
  .addNode("generate_sql", generate_sql)
  .addNode("obtain_data", obtain_data)
  .addNode("generate_summary", generate_summary)
  .addNode("export_docx", export_docx)

  .addConditionalEdges("__start__", conditionTest)
  // .addEdge("income_statement", "generate_sql")
  // .addEdge("balance_sheet", "generate_sql")
  .addEdge("financial_ratios", "generate_sql")
  // .addEdge("cash_flow_statement", "generate_sql")
  // .addEdge("customer_transaction", "generate_sql")

  .addEdge("generate_sql", "obtain_data")
  .addEdge("obtain_data", "generate_summary")
  .addEdge("generate_summary", "export_docx")
  .addEdge("export_docx", "__end__")
  .compile();

  // Load environment variables
  dotenv.config();




// 一鍵生成報告，依chapter跑迴圈執行，產出每個章節的數據表格與AI建議文字
export async function establishReport() {
  for(let chapter of chapters){

    const userInputResponse = await new Promise(async (resolve) => {
      const result = await graph.invoke({
        question:'Establish a suggestion of specific chapter with reference data by querying db.',
        chapter:chapter,
        query:"",
        result:"",
        answer:"",
        summary:""
      });
      // console.log(`establishReport result========`,result);

      resolve(result);

    });

    try {
      console.log(`\nThe chapter ${chapter} is finished.\n`);
      console.log(`\nThe chapter's answer is ${userInputResponse.answer}.\n`);
      console.log(`\nThe chapter's summary is ${userInputResponse.summary}.\n`);
    } catch (error) {
      console.error("An error occurred:", error.message);
    }
  }
}

// Start the chat
// chat().catch(console.error); 
// establishReport().catch(console.error); 



