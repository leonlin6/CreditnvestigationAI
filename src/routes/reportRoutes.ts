import express from "express";
import { getReport, getCompanyNames } from "../controllers/reportController.js";
import { SqlDatabase } from "langchain/sql_db";
import { DataSource } from "typeorm";
import { DataSourceOptions } from "typeorm";

const router = express.Router();
const dbConfig: DataSourceOptions = {
  type: "sqlite",
  database: process.env.SQLITE_DB || "credit_check.db",
};
const creditDatasource = new DataSource(dbConfig);
const db = await SqlDatabase.fromDataSourceParams({
  appDataSource: creditDatasource,
});

export const setupReportRoutes = (app: express.Application) => {
  /**
   * @route GET /api/report
   * @desc Generate and download a credit report for a specific company and year
   * @param {string} year - The year for the report
   * @param {string} companyName - The name of the company
   * @returns {Buffer} Word document (.docx) file
   */
  router.get("/report", async (req, res) => {
    const { year, companyName } = req.query;

    if (!year || !companyName) {
      return res
        .status(400)
        .json({ error: "Year and company name are required" });
    }

    try {
      const buffer = await getReport(db, Number(year), companyName as string);
      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      );
      res.setHeader("Content-Disposition", "attachment;");
      res.send(buffer);
    } catch (error) {
      console.error("Error processing request:", error);
      res.status(500).json({ error: "Internal server error" });
    }
  });

  /**
   * @route GET /api/companies/name
   * @desc Get a list of all company names in the database
   * @returns {string[]} Array of company names
   */
  router.get("/companies/name", async (req, res) => {
    try {
      const companyNameList = await getCompanyNames(db);
      res.send(companyNameList);
    } catch (error) {
      console.error("Error getting company names:", error);
      res.status(500).json({ error: "Internal server error" });
    }
  });

  app.use("/api", router);
};
