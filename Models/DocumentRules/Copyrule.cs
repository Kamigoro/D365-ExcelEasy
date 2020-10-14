using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace D365_ExcelModifier.Models.DocumentRules
{
    public class Copyrule : BaseRule
    {
        public string InputFile { get; set; }
        public string OutputFile { get; set; }
        public string InputColumn { get; set; }
        public string OutputColumn { get; set; }

        public override void Execute()
        {
            try
            {
                using var inWorkbook = new XLWorkbook(InputFile);
                using var outWorkbook = new XLWorkbook(OutputFile);
                var outputColumn = outWorkbook.Worksheets.ElementAt(0).Column(OutputColumn);

                foreach (var inputWorksheet in inWorkbook.Worksheets)
                {
                    //Get the column where the header is the correct one
                    var inputColumn = inputWorksheet.Columns()
                        .FirstOrDefault(column => column.Cell(1)
                        .Value
                        .ToString()
                        .StartsWith(InputColumn));

                    //if we found a column matching with InputColumn
                    if(inputColumn != null)
                    {
                        //We get the cells of the selected column and we append them to the output file in the specified column of the first sheet
                        var cellsOfInputColumn = inputColumn.Cells().ToList();
                        cellsOfInputColumn.RemoveAt(0);//Removing header

                        foreach (var inputCell in cellsOfInputColumn)
                        {
                            outputColumn.Cell(outputColumn.LastCell().WorksheetRow().RowNumber() + 1).Value = inputCell.Value.ToString();
                        }
                    }
                }

                inWorkbook.Save();
                outWorkbook.Save();
                RuleExecuted_EventHandler?.Invoke(this, new RuleEventArgs(this, true));
            }
            catch (Exception e)
            {
                RuleExecuted_EventHandler?.Invoke(this, new RuleEventArgs(this, true, e.Message));
                throw e;
            }
        }
    }
}
