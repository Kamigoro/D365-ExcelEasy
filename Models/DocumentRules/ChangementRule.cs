using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace D365_ExcelModifier.Models.DocumentRules
{
    public class ChangementRule : BaseRule
    {
        public string InputFile { get; set; }
        public string OldValue { get; set; }
        public string NewValue { get; set; }

        public override void Execute()
        {
            try
            {
                using (var workbook = new XLWorkbook(InputFile))
                {
                    foreach (var worksheet in workbook.Worksheets)
                    {
                        //Find all the cells that starts by OldValue
                        var cellsToReplace = worksheet.Cells().
                            Where(cell =>
                            cell.Value.ToString()
                            .StartsWith(OldValue));
                        //Replacing the value of the cells concerned
                        foreach (var cell in cellsToReplace)
                        {
                            cell.Value = NewValue;
                        }
                    }
                    workbook.Save();
                    RuleExecuted_EventHandler?.Invoke(this, new RuleEventArgs(this, true));
                }
            }
            catch (ArgumentNullException e)
            {
                RuleExecuted_EventHandler?.Invoke(this, new RuleEventArgs(this, false, "Nous n'avons pas trouvé de cellules contenant cette valeur"));
                throw e;
            }
            catch (Exception e)
            {
                RuleExecuted_EventHandler?.Invoke(this, new RuleEventArgs(this, false, e.Message));
                throw e;
            }
            
        }
    }
}
