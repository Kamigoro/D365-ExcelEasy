using ClosedXML.Excel;
using D365_ExcelModifier.Models.DocumentRules;
using System;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;

namespace D365_ExcelModifier.Models.Actions
{
    public class CopyInOtherFileAction
    {
        public IXLWorksheets InputWorksheets { get; set; }
        public IXLWorksheets OuputWorksheets { get; set; }
        public CopyInOtherFileRule Rule { get; set; }

        public CopyInOtherFileAction(CopyInOtherFileRule rule, IXLWorksheets inputWorksheets, IXLWorksheets outputWorksheets)
        {
            Rule = rule;
            InputWorksheets = inputWorksheets;
            OuputWorksheets = outputWorksheets;
        }

        public bool Execute()
        {
            try
            {
                //we assume we paste the data in the first worksheet of the output workbook
                var outputworkseet = OuputWorksheets.ElementAt(0);
                var outputColumn = outputworkseet.Columns().First(column => column.Cell(1).Value.ToString().StartsWith(Rule.OutputColumn));

                //we want to get the last row used of the outputworksheet to append after it
                int lastUsedRowIndex = outputworkseet.LastRowUsed().RowNumber();
                int outputColumnIndex = outputColumn.ColumnNumber();

                //We start writing in the output work sheet at the newt avaliable row in the specified column
                int nextNewCellRowIndex = lastUsedRowIndex + 1;
                foreach (var inputWorksheet in InputWorksheets)
                {
                    //Get the column where the header is the correct one
                    var inputColumn = inputWorksheet.Columns()
                        .FirstOrDefault(column => column.Cell(1)
                        .Value
                        .ToString()
                        .StartsWith(Rule.InputColumn));

                    //To avoid null exception
                    if (inputColumn != null)
                    {
                        //We get the cells of the selected column and we append them to the output file in the specified column of the first sheet
                        var cellsOfInputColumn = inputColumn.Cells().ToList();
                        cellsOfInputColumn.RemoveAt(0);//Removing header
                        foreach (var inputcell in cellsOfInputColumn)
                        {
                            outputworkseet.Cell(nextNewCellRowIndex, outputColumnIndex).Value = inputcell.Value;
                            nextNewCellRowIndex++;
                        }
                    }
                }
                return true;
            }
            catch(Exception e)
            {
                Debug.WriteLine(e);
                return false;
            }
        }
    }
}
