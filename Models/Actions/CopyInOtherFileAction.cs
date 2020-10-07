using ClosedXML.Excel;
using D365_ExcelModifier.Models.DocumentRules;
using System.Linq;
using System.Threading.Tasks;

namespace D365_ExcelModifier.Models.Actions
{
    public class CopyInOtherFileAction : IDocumentAction
    {
        public IXLWorksheet Worksheet { get; set; }
        public CopyInOtherFileRule Rule { get; set; }

        public CopyInOtherFileAction(CopyInOtherFileRule rule)
        {
            Rule = rule;
        }

        public async Task<bool> ExecuteAsync()
        {
            await Task.Yield();
            try
            {
                using (var inputWorkbook = new XLWorkbook(Rule.InputFile))
                {
                    using (var outputWorkbook = new XLWorkbook(Rule.OutputFile))
                    {
                        //we assume we paste the data in the first worksheet of the output workbook
                        var outputworkseet = outputWorkbook.Worksheet(1);
                        var outputColumn = outputworkseet.Columns().First(column => column.Cell(1).Value.ToString().StartsWith(Rule.OutputColumn));

                        //we want to get the last row used of the outputworksheet to append after it
                        int lastUsedRowIndex = outputworkseet.LastRowUsed().RowNumber();
                        int outputColumnIndex = outputColumn.ColumnNumber();

                        //We start writing in the output work sheet at the newt avaliable row in the specified column
                        int nextNewCellIndex = lastUsedRowIndex + 1;
                        foreach (var inputWorksheet in inputWorkbook.Worksheets)
                        {
                            var inputColumn = inputWorksheet.Columns().First(column => column.Cell(1).Value.ToString().StartsWith(Rule.InputColumn));
                            foreach (var inputcell in inputColumn.Cells())
                            {
                                outputworkseet.Cell(nextNewCellIndex, outputColumnIndex).Value = inputcell.Value;
                                nextNewCellIndex++;
                            }

                        }

                    }
                }
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
