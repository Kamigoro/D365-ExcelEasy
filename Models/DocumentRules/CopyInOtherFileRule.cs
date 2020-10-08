namespace D365_ExcelModifier.Models.DocumentRules
{
    public class CopyInOtherFileRule : DocumentRuleBase
    {
        public CopyInOtherFileRule()
        {

        }

        public CopyInOtherFileRule(string inputColumn, string outputColumn)
        {
            InputColumn = inputColumn;
            OutputColumn = outputColumn;
        }
    }
}
