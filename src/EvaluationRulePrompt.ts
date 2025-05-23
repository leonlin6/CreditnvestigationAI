const EvaluationRulePrompt = {
  debt_to_asset_ratio_rule : `If debt_to_asset_ratio value is lower than 30, it means Conservative financial management; possibly indicates inefficient capital utilization.` +
  `If debt_to_asset_ratio value is between 30 and 50, it means the company maintains a moderate level of financial leverage—effectively using debt to expand operations while retaining a sound financial structure. Debt levels are reasonable, and the company demonstrates good debt repayment ability.` +
  `If debt_to_asset_ratio value is between 50 and 60, it means The financial structure remains relatively stable, but the debt ratio has reached a moderate level. Close attention should be paid to debt management efficiency and repayment capacity to avoid increased financial risk due to overexpansion.` +
  `If debt_to_asset_ratio value is above 60, it means Heavier debt burden with increased financial risk.` ,

  current_ratio_rule: `If current_ratio is value lower than 100, it means insufficient short-term debt repayment ability.` +
  `If current_ratio value is between 150 and 250, it means the company has a strong short-term debt repayment capacity. Current assets can effectively cover current liabilities, financial liquidity is healthy, and the company is well-positioned to meet operational funding needs and unexpected situations.` +
  `If current_ratio value is between 100 and 150, it means the short-term debt repayment ability is acceptable but cash flow management should be improved.` +
  `If current_ratio value is between 250 and 300, it means liquidity is abundant, but there may be inefficiencies in utilizing current funds. It is recommended to optimize cash allocation.` +
  `If current_ratio value is above 300, it means the company has excessive liquid assets and should improve capital utilization efficiency.`,

  quick_ratio_rule:`If quick_ratio value is lower than 80, it means the company has weak immediate debt repayment capacity.` +
  `If quick_ratio value is between 80 and 100, it means the immediate liquidity is acceptable but may require attention to cash flow stability.` +
  `If quick_ratio value is between 100 and 130, it means the company maintains an excellent level of quick ratio. Even excluding inventory, current assets can sufficiently cover current liabilities. This reflects strong immediate repayment ability, sound cash and short-term investment management, and good financial flexibility.` +
  `If quick_ratio value is between 130 and 150, it means liquidity remains generally acceptable, but excess cash may be better allocated toward more productive investments to improve overall capital efficiency.` +
  `If quick_ratio value is above 150, it means the company may be managing cash too conservatively.`,

  roa_rule:`If return_on_assets value is lower than 5, it means the company has relatively low asset utilization efficiency.` +
  `If return_on_assets value is between 5 and 10, it means the company has basic profitability and acceptable asset utilization, but there is still room for improvement. Reviewing and optimizing operational strategies is recommended to enhance returns.` +
  `If return_on_assets value is between 10 and 15, it means excellent performance. It indicates that the management team effectively utilizes company assets to generate strong profits, with high returns on investments and outstanding operational efficiency, creating significant value for shareholders.` +
  `If return_on_assets value is above 15, it means exceptional performance.`,

  roe_rule:`If return_on_equity value is lower than 10, it means the return on shareholders' equity is relatively low.` +
  `If return_on_equity value is between 10 and 15, it means the company demonstrates stable profitability and provides a reasonable return to shareholders, though there is still room to improve efficiency compared to industry leaders.` +
  `If return_on_equity value is between 15 and 25, it means strong performance. The company generates high returns on shareholders’ equity, indicating that management excels at utilizing shareholder capital. This level of performance is attractive to investors.` +
  `If return_on_equity value is above 25, it means outstanding performance.`,

  accounts_receivable_turnover_rule: `If accounts_receivable_turnover is lower than 6, it means the collection efficiency is poor. This may indicate issues such as customer credit risk or overly lenient collection policies, which can negatively impact cash flow and working capital turnover.` +
  `If accounts_receivable_turnover is between 6 and 9, it means the collection efficiency is generally acceptable. However, cash flow could still be improved through enhanced credit policies or stronger collection efforts.` +
  `If accounts_receivable_turnover is above 9, it means excellent performance. The company demonstrates effective enforcement of its credit policies, high-quality customers, stable cash flow, and efficient working capital management.`,

  inventory_turnover_rule: `If inventory_turnover is lower than 4, it means there is a serious inventory buildup. This may indicate poor sales performance, obsolete or depreciating products, and highlights the need to improve product mix planning and inventory control strategies.` +
  `If inventory_turnover is between 4 and 6, it means the inventory management is within an acceptable range. However, improving demand forecasting or supply chain coordination could help increase turnover efficiency.` +
  `If inventory_turnover is above 6, it means good performance. The company is managing inventory effectively, with smooth product sales, reduced capital lock-in, and lower risk of inventory depreciation.`,

   total_asset_turnover_rule : `If total_asset_turnover is lower than 0.5, it means the company has significantly low asset utilization efficiency. This may indicate idle capacity, poor asset allocation, or weak market demand. It is recommended to review asset management strategies and operational models.` +
  `If total_asset_turnover is between 0.5 and 1.2, it means the company operates at a generally acceptable level. However, there is still room for improvement by enhancing capacity utilization or optimizing asset allocation to increase operational efficiency.` +
  `If total_asset_turnover is above 1.2, it means excellent performance. The company is effectively using its assets to generate revenue, showing strong operational efficiency. Every dollar of assets generates more than a dollar in revenue, highlighting management’s strong resource allocation capability.`
}



export default EvaluationRulePrompt;