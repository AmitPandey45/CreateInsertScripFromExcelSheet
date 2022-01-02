using Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CreateInsertScripFromExcelSheet.cs
{
    class Program
    {
        const string ScriptSeparator = "\n\n";
        const string TableName = "db_MEM.CommitteeMemberClassification";
        static Dictionary<string, BusinessRule> BusinessColumns = new Dictionary<string, BusinessRule>
        {
            { "COMMITTEEPRIMARYACTIVITYID", new BusinessRule { ColumnNumber = -1, Rule = null } },
            { "WEBSITE", new BusinessRule { ColumnNumber = -1, Rule = new int[]  { 1, 2, 3, 4, 6, 7 } } },
            { "FACILITYORGANIZATION", new BusinessRule { ColumnNumber = -1, Rule = new int[]  { 1, 2, 3, 4, 5, 6, 7, 11 } } },
            { "PARENTORGANIZATION" , new BusinessRule { ColumnNumber = -1, Rule = new int[]  { 1, 2, 3, 4 } } },
            { "CODIVISION", new BusinessRule { ColumnNumber = -1, Rule = new int[]  { 1, 2, 3, 4, 6, 7 } } }
        };

        static int BusinessDerivedColumnValue = -1;

        static void Main(string[] args)
        {
            CreateInsertScript();

            Console.ReadKey();
        }


        private static void CreateInsertScript()
        {
            int rowNumber = 1;
            int columnNumber = 1;
            int lastColumnNumber = 0;
            var insertScript = new StringBuilder();
            var columnScript = new StringBuilder();
            var rowScript = new StringBuilder();
            IEnumerable<Cell> cells;

            foreach (var worksheet in Workbook.Worksheets(@"D:\MY SOFTWARE AND FILES\CommitteeMemberClassification.xlsx"))
            {
                rowNumber = 1;
                columnScript.Clear();

                foreach (var row in worksheet.Rows)
                {
                    columnNumber = 1;
                    if (rowNumber == 1)
                    {
                        cells = row.Cells.Where(w => w != null);
                        lastColumnNumber = cells.Count();
                    }
                    else
                    {
                        cells = row.Cells;
                    }

                    foreach (var cell in cells)
                    {
                        if (rowNumber == 1)
                        {
                            string key = cell?.Text?.Trim().ToUpper();
                            if (BusinessColumns.ContainsKey(key) && BusinessColumns[key].ColumnNumber == -1)
                            {
                                BusinessColumns[key].ColumnNumber = columnNumber;
                            }

                            CreateColumnScript(columnScript, columnNumber, lastColumnNumber, cell?.Text);
                        }
                        else
                        {
                            CreateRowScript(rowScript, columnNumber, lastColumnNumber, cell?.Text);
                        }

                        columnNumber++;
                    }

                    if (rowNumber > 1)
                    {
                        insertScript.Append(columnScript);
                        insertScript.Append(rowScript);
                        insertScript.Append(ScriptSeparator);
                        rowScript.Clear();
                    }

                    rowNumber++;
                }
            }

            Console.WriteLine(insertScript.ToString());
        }

        private static void CreateColumnScript(StringBuilder columnScript, int columnNumber, int lastColumnNumber, string headerName)
        {
            if (columnNumber == 1)
            {
                columnScript.Append($"INSERT INTO {TableName} (");
            }

            columnScript.Append($"{headerName}");

            if (columnNumber != lastColumnNumber)
            {
                columnScript.Append(", ");
            }
            else
            {
                columnScript.Append(")");
            }
        }

        private static void CreateRowScript(StringBuilder rowScript, int columnNumber, int lastColumnNumber, string value)
        {
            if (columnNumber == 1)
            {
                rowScript.Append($"VALUES (");
            }

            if (string.IsNullOrEmpty(value) || value.ToLower().Equals("null"))
            {
                value = null;
            }
            else if (value.Contains("’"))
            {
                value = value.Replace("’", "''");
            }
            else if (value.Contains("'"))
            {
                value = value.Replace("'", "''");
            }

            BusinessRule businessRule = BusinessColumns
                .Select(w => w.Value)
                .FirstOrDefault(w => w.ColumnNumber.Equals(columnNumber));

            if (businessRule != null)
            {
                if (businessRule.Rule != null)
                {
                    value = businessRule.Rule.Contains(BusinessDerivedColumnValue) ? value : null;
                }
                else
                {
                    businessRule.BusinessDerivedColumnValue = value;
                    BusinessDerivedColumnValue = Convert.ToInt32(businessRule.BusinessDerivedColumnValue.Trim());
                }
            }

            if (!string.IsNullOrEmpty(value))
            {
                rowScript.Append($"'{value}'");
            }
            else
            {
                rowScript.Append("NULL");
            }

            if (columnNumber != lastColumnNumber)
            {
                rowScript.Append(", ");
            }
            else
            {
                rowScript.Append(")");
            }
        }
    }
}
