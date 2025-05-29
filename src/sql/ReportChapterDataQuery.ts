export const SQL = {
  income_statement: `SELECT * FROM users ORDER BY created_at DESC;`,
  balance_sheet: `SELECT * FROM users WHERE id = $1;`,
  financial_ratios: `INSERT INTO logs (type, message, created_at) VALUES ($1, $2, NOW());`,
  cash_flow_statement: `INSERT INTO logs (type, message, created_at) VALUES ($1, $2, NOW());`,
  customer_transaction: `INSERT INTO logs (type, message, created_at) VALUES ($1, $2, NOW());`,
};

