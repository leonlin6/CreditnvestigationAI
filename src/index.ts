import "reflect-metadata";
import dotenv from "dotenv";
import express from 'express';
import cors from 'cors';
import { setupReportRoutes } from './routes/reportRoutes.js';

// Load environment variables
dotenv.config();

// express API Server handler
const app = express();
const port = process.env.PORT || 3000;

const corsOptions = {
  origin: [
    'https://web-creditnvestigation-70fhvfear-leonlin6s-projects-5cda9579',
    'https://web-creditnvestigation-70fhvfear-leonlin6s-projects-5cda9579.vercel.app',
    "https://web-creditnvestigation-ai.vercel.app",
    "http://localhost:3001"
  ],
  methods: 'GET,HEAD,PUT,PATCH,POST,DELETE,OPTIONS',
  allowedHeaders: ['Content-Type', 'Authorization'],
};

app.use(cors(corsOptions));
app.use(express.json());

// Setup API routes
setupReportRoutes(app);

app.listen(port, () => {
  console.log(`Example app listening on port ${port}`);
});
