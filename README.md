# xcl_risk_parser

VBA Excel Macro: Financial Obligations Analyzer

This VBA macro analyses financial data on an Excel worksheet ("Sheet1") to:

Identify and sum negative values associated with a predefined list of creditors or obligations, excluding specific keywords (e.g., "Spotify").
Extract rows with recognized obligations and any income entries, copying them to a second worksheet ("Sheet2").
Calculate the 4-month average obligations and income, presenting this summary at the top of "Sheet2".
Sort the obligations in descending order by date on "Sheet2".
Remove specific columns' values (G, I, and J) for clarity.
In essence, this macro is useful for summarizing and analyzing financial obligations over a 4-month period from a detailed transaction list.
