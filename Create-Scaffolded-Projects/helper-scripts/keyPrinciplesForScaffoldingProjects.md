# Google Sheets Dashboarding Hackathon (June 1)

### Google Sheets has two major elements when working with formulas:
RANGE and VALUE -> Usually, you either refer to a value like B1 or range like B1:B most of the formulas will accept something on those lines, like:
=counta(value, range)
=sum(range)
=vlookup(value, range, number, boolean)

## Formulas:
- UNIQUE -> returns a RANGE including a null value (an extra row which you can't change)
- QUERY -> Query formula gives a SQL like interface to query google sheets data :)
- IMPORTRANGE -> will help you to not duplicate data. Whenever you need data from other sheet, you simply import it, instead of copying it.
- SPARKLINE -> Tool to visualize data properties in a single cell
- LET -> Helps you write readable, structured, mini-code/function. Parameters pair up, first is the name for a variable, followed by value assigned. Last parameter is the value to be returned.
- SUBTOTAL -> //need to go through this, went over my head


## Data Validation Rules:
- ISURL ->
- REGEXMATCH ->
- You can reject the Input
- Add a prompt or warning message

## Named Ranges:
- A variable to refer to a range

## Named Function:
- Template for a creating complex reusable functions, parameterizing input to the formula and reuse it anywhere you like

## More Automation Tool:
### AppScripts
- Use AI to get formulas to add functionality, even add visual elements like a menu to trigger an operation.
- We can customize the error message for the end-user, in case the script fails

### Data Connectors
- You can connect multitudes of with data sources

### Protect Sheets and Ranges
- You can apply authentication for editing data
- Pair it with AppScripts to unleash its true powers :)