﻿using ClosedXML.Excel;
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
                    var cellsToReplace = workbook.Search(OldValue);

                    if (cellsToReplace.Count() == 0)
                    {
                        RuleExecuted_EventHandler?.Invoke(this, new RuleEventArgs(this, false, "Nous n'avons pas trouvé de cellules contenant cette valeur"));
                    }
                    else
                    {
                        foreach (var cell in cellsToReplace)
                        {
                            cell.Value = NewValue;
                        }
                    }
                    RuleExecuted_EventHandler?.Invoke(this, new RuleEventArgs(this, true));
                    workbook.Save();
                }
            }
            catch (ArgumentNullException e)
            {
                RuleExecuted_EventHandler?.Invoke(this, new RuleEventArgs(this, false, $"Nous n'avons pas trouvé de cellules contenant cette valeur\n{e.Message}"));
            }
            catch (Exception e)
            {
                RuleExecuted_EventHandler?.Invoke(this, new RuleEventArgs(this, false, e.Message));
            }

        }
    }
}
