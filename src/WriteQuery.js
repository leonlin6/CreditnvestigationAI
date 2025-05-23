import { z } from "zod";
import { Annotation } from "@langchain/langgraph";

const InputStateAnnotation = Annotation.Root({
  question: Annotation
});

const StateAnnotation = Annotation.Root({
  question: Annotation,
  query: Annotation,
  result: Annotation,
  answer: Annotation
});

const queryOutput = z.object({
  query: z.string().describe("Syntactically valid SQL query."),
});

const structuredLlm = llm.withStructuredOutput(queryOutput);

const writeQuery = async (state) => {
  const promptValue = await queryPromptTemplate.invoke({
    dialect: db.appDataSourceOptions.type,
    top_k: 10,
    table_info: await db.getTableInfo(),
    input: state.question,
  });
  const result = await structuredLlm.invoke(promptValue);
  return { query: result.query };
};

export { writeQuery };