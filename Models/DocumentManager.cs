using ClosedXML.Excel;
using D365_ExcelModifier.Models.Actions;
using D365_ExcelModifier.Models.DocumentRules;
using System.Collections.Generic;

namespace D365_ExcelModifier.Models
{
    public class DocumentManager
    {
        public string InputFile { get; set; }
        public string OutputFile { get; set; }
        private List<CopyInOtherFileRule> CopyInOtherFileRules { get; set; } = new List<CopyInOtherFileRule>();
        private List<ValueChangementRule> ValueChangementRules { get; set; } = new List<ValueChangementRule>();

        public DocumentManager(List<DocumentRuleBase> baseRules, string inputFile, string outputFile)
        {
            InputFile = inputFile;
            OutputFile = outputFile;

            foreach (DocumentRuleBase baseRule in baseRules)
            {
                if (baseRule.OutputColumn == null)
                {
                    ValueChangementRules.Add(new ValueChangementRule(baseRule.InputColumn, baseRule.OldValue, baseRule.NewValue));
                }
                else
                {
                    CopyInOtherFileRules.Add(new CopyInOtherFileRule(baseRule.InputColumn, baseRule.OutputColumn));
                }
            }
        }

        public void ExecuteRules()
        {
            using var inputWorkbook = new XLWorkbook(InputFile);
            using var outputWorkbook = new XLWorkbook(OutputFile);
            foreach (ValueChangementRule valueChangementRule in ValueChangementRules)
            {
                ValueChangingAction action = new ValueChangingAction(valueChangementRule, inputWorkbook.Worksheets);
                action.Execute();
            }

            foreach (CopyInOtherFileRule copyInOtherFileRule in CopyInOtherFileRules)
            {
                CopyInOtherFileAction action = new CopyInOtherFileAction(copyInOtherFileRule, inputWorkbook.Worksheets, outputWorkbook.Worksheets);
                action.Execute();
            }

            outputWorkbook.Save();
            inputWorkbook.Save();
        }

    }
}
