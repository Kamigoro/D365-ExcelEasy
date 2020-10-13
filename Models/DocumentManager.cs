using ClosedXML.Excel;
using D365_ExcelModifier.Models.Actions;
using D365_ExcelModifier.Models.DocumentRules;
using System;
using System.Collections.Generic;

namespace D365_ExcelModifier.Models
{
    public class DocumentManager
    {
        public string InputFile { get; set; }
        public string OutputFile { get; set; }
        private List<CopyInOtherFileRule> CopyInOtherFileRules { get; set; } = new List<CopyInOtherFileRule>();
        private List<ValueChangementRule> ValueChangementRules { get; set; } = new List<ValueChangementRule>();
        public Action FinishedExecution_Event;
        private RuleType ruleType;

        public DocumentManager(List<CopyInOtherFileRule> copyInOtherFileRules, string inputFile, string outputFile)
        {
            CopyInOtherFileRules = copyInOtherFileRules;
            InputFile = inputFile;
            OutputFile = outputFile;
            ruleType = RuleType.CopyInOtherFile;
        }

        public DocumentManager(List<ValueChangementRule> valueChangementRules, string inputFile)
        {
            ValueChangementRules = valueChangementRules;
            InputFile = inputFile;
            ruleType = RuleType.ValueChangement;
        }

        public void ExecuteRules()
        {
            switch (ruleType)
            {
                case RuleType.CopyInOtherFile:
                    ExecuteCopyInOtherFilesRules();
                    break;
                case RuleType.ValueChangement:
                    ExecuteValueChangementRules();
                    break;
                default:
                    return;
            }

            FinishedExecution_Event?.Invoke();
        }


        private void ExecuteValueChangementRules()
        {
            using var inputWorkbook = new XLWorkbook(InputFile);
            foreach (ValueChangementRule valueChangementRule in ValueChangementRules)
            {
                ValueChangingAction action = new ValueChangingAction(valueChangementRule, inputWorkbook.Worksheets);
                action.Execute();
            }
            inputWorkbook.Save();
        }

        private void ExecuteCopyInOtherFilesRules()
        {
            using var inputWorkbook = new XLWorkbook(InputFile);
            using var outputWorkbook = new XLWorkbook(OutputFile);


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
