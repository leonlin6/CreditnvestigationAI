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
import { excuteNewDoc } from './TableTransfer.ts';

import { establishReport } from "./text.ts";

await establishReport().catch(console.error); 
// const sqlite333 = require('sqlite3').verbose()

// Practice of linking to database
// var ddb = new sqlite3.Database(
//   './chinook.db', 
//   sqlite3.OPEN_READWRITE, 
//   function (err) {
//       if (err) {
//           return console.log(err.message)
//       }
//       console.log('connect database successfully')
//   }
// )

// ddb.all('SELECT LastName, FirstName, Title FROM employees LIMIT 10', [], function (err, rows) {
//   if (err) {
//       return console.log('find Alice error: ', err.message)
//   }

//   console.log('find Alice: ', rows);
//   const LastNameArrat = rows.map((item)=>{
//     return item.LastName;
//   }) 

// })

// Initialize the OpenAI model
const chatOpenAImodel = new ChatOpenAI({
  modelName: "gpt-3.5-turbo",
  temperature: 0.7,
  openAIApiKey: process.env.OPENAI_API_KEY
});

// const chatOpenAImodel = new ChatOllama({
//   model: "llama3.2:latest",
//   temperature: 0,
//   // other params...
// });


// Create the prompt template
// const prompt = ChatPromptTemplate.fromMessages([
//   ["system", "You are a helpful AI assistant that provides concise and accurate information."],
//   ["human", "{input}"]
// ]);

// Create the chain
// const chain = RunnableSequence.from([
//   prompt,
//   chatOpenAImodel,
//   new StringOutputParser(),
// ]);

// HTTP Server handler
// function handler(req, res) {
//   console.log(req.url);  // Log the request URL
//   res.write('Hello World!');   // Set response content
  
//   res.end()  ;   // End the response
// }

// Create HTTP server
// const server = createServer(handler);

// Start HTTP server
// server.listen(5001, () => {
//   console.log('Server running at http://localhost:5001/');
// });



const InputStateAnnotation = Annotation.Root({
  question: Annotation<string>
});

const StateAnnotation = Annotation.Root({
  question: Annotation<string>,
  query:Annotation<string>,
  result: Annotation<string>,
  answer: Annotation<string>,
});

const GenerateReportAnnotation = Annotation.Root({
  chapter: Annotation<string>,
  query:Annotation<string>,
  result: Annotation<string>,
  summary: Annotation<string>,
});

// Step1 :先執行queryPromptTemplate，用template產出SQL語法
// dialect: 符合此db type的SQL
// top_k: 查詢結果最多top_k筆
// table_info: 指定的table
// input: user's question
const retrieve = async (state: typeof StateAnnotation.State) => {
  const promptValue = await queryPromptTemplate.invoke({
    dialect: db.appDataSourceOptions.type,
    top_k: 10,
    table_info: await db.getTableInfo(['financial_ratios']),
    input: `Please get all data in 2024 with 'gui_no' is '27790294'`,
  });

  const result = await structuredLlm.invoke(promptValue);
  return {
      question: state.question,
      query: result.query,
      result: "",
      answer: "",
    };
  };
  
// Step2: 用retrieve產出的SQL語法，用langchain的QuerySqlTool去DB撈資料
const generate = async (state: typeof  StateAnnotation.State) => {
  const executeQueryTool = new QuerySqlTool(db);
  const result = await executeQueryTool.invoke(state.query);
   
  return { 
    question: state.question,
    query: state.query,
    result: result,
    answer: "",
  };
};
  

  
const queryOutput = z.object({
  query: z.string().describe("Syntactically valid SQL query."),
});

// const structuredLlm = llm.withStructuredOutput(queryOutput);
const structuredLlm = chatOpenAImodel.withStructuredOutput(queryOutput);




// Step1: use langchain to connect to sqlite database
// const chinookDatasource = new DataSource({
//   type: "sqlite",
//   database: "chinook.db",
// });

const creditDatasource = new DataSource({
  type: "sqlite",
  database: "credit_check.db",
});

const db = await SqlDatabase.fromDataSourceParams({
  appDataSource: creditDatasource,
});

// Step2:  load the sql-query-system-prompt from prompt template of langchain hub
// 不知道為什麼<ChatPromptTemplate>不能加在泛型裡面，會造成return false，拿掉就可以正常使用
const queryPromptTemplate = await pull(
  "langchain-ai/sql-query-system-prompt"
);


// const writeQuery = async (state: typeof InputStateAnnotation.State) => {
  
//   // Use formatMessages instead of invoke
//   const promptValue = await queryPromptTemplate.invoke({
//     dialect: db.appDataSourceOptions.type,
//     top_k: 10,
//     table_info: await db.getTableInfo(),
//     input: state.question,
//   });

  
//   const result = await structuredLlm.invoke(promptValue);
//   return { query: result.query };

//   // return 'test';
// };

// Step3: 呼叫WriteQuery
// const answer =await writeQuery({ question: "what is the weather today in Taipei? " });
// console.log('answer',answer);


// const executeQuery = async (state) => {
//   const executeQueryTool = new QuerySqlTool(db);
//   return { result: await executeQueryTool.invoke(state.query) };
// };

// Step3: 設定輸出用的prompt，並用剛剛DB撈出的資料，丟給LLM產出人性化的對話給User看
const generateAnswer = async (state) => {
  console.log('generateAnswer state=======',state);
  const parsedJson = JSON.parse(state.result);
  console.log('parsedJson state=======',parsedJson);
  console.log('parsedJson state[0]=======',parsedJson[0]);
const str = JSON.stringify(parsedJson[0]);
  const promptValueExportText =
    // "If LLM cannot find the related data in the datbase, responsing `Sorry, the question is Out of my service area.` else " +
    // "Given the following user question, corresponding " +
    // " SQL result, answer the user question. \n" +
    "answer the user question. \n" +
    // `QUESTION:What are the value of debt_to_asset_ratio, current_ratio, quick_ratio by following data: ${str}` ;
    `Question: Please use debt_to_asset_ratio, current_ratio, quick_ratio in the ${str} and rules to establish a comment for the company. And then translate it into Traditional Chinese.` +
    `Rules:` +
    `If debt_to_asset_ratio is higer than 50, it means the company has a high debt burden.` +
    `If current_ratio is higher than 100, it means the company may have a short-term liquidity problem.` +
    `If quick_ratio is higher than 100, it means the company may have a short-term liquidity problem.`;
  // const response = await llm.invoke(promptValue);

  console.log('promptValueExportText',promptValueExportText);
  const textResponse = await chatOpenAImodel.invoke(promptValueExportText);

  
  return { answer: textResponse.content };
};

const conditionTest = (state) => {
  if(state.question == ''){
    return "generate";
  }else{
    return "retrieve";
  }
}
  // Compile application and test
const graph = new StateGraph(StateAnnotation)
  .addConditionalEdges("__start__", conditionTest)
  .addNode("retrieve", retrieve)
  .addNode("generate", generate)
  .addNode("generateAnswer", generateAnswer)
  .addEdge("retrieve", "generate")
  .addEdge("generate", "generateAnswer")
  .addEdge("generate", "__end__")
  .compile();
  
// Load environment variables
dotenv.config();

// Create readline interface for user input
const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

// 用rl readline以後，直接處理去CALL LLM撈SQL與處理SQL後的ANSWER
async function chat() {
  console.log("Welcome to the OpenAI Chat Example!");
  console.log("Type 'quit' to exit the chat.\n");

  while (true) {
    const userInputResponse = await new Promise((resolve) => {
      rl.question("You: ", async (text) => {
        const result = await graph.invoke({
          question:text,
          query:"",
          result:"",
          answer:"",
        });
        resolve(result);
      });
    });

    // if (userInput.toLowerCase() === "quit") {
    //   console.log("Goodbye!");
    //   rl.close();
    //   break;
    // }

    try {
      // const response = await chain.invoke({
      //   input: userInput,
      // });
      console.log(`\nAI:${userInputResponse.answer}\n`);
    } catch (error) {
      console.error("An error occurred:", error.message);
    }
  }
}

const chapters = [
  'income_statement',
  'balance_sheet',
  'financial_ratios',
  'cash_flow_statement',
  'customer_transaction'
]

// 一鍵生成報告，依chapter跑迴圈執行，產出每個章節的數據表格與AI建議文字
// async function establishReport() {

//   for(let chapter of chapters){
//     const userInputResponse = await new Promise((resolve) => {
//       rl.question("You: ", async (text) => {
//         const result = await graph.invoke({
//           question:text,
//           query:"",
//           result:"",
//           answer:"",
//         });
//         resolve(result);
//       });
//     });

//     try {
//       console.log(`\nThe chapter ${chapter} is finished.\n`);
//     } catch (error) {
//       console.error("An error occurred:", error.message);
//     }
//   }

//     const userInputResponse = await new Promise((resolve) => {
//       rl.question("You: ", async (text) => {
//         const result = await graph.invoke({
//           question:text,
//           query:"",
//           result:"",
//           answer:"",
//         });
//         resolve(result);
//       });
//     });

//     try {
//       console.log(`\nAI:${userInputResponse.answer}\n`);
//     } catch (error) {
//       console.error("An error occurred:", error.message);
//     }
// }

// Start the chat
// chat().catch(console.error); 




