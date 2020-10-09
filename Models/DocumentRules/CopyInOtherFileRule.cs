namespace D365_ExcelModifier.Models.DocumentRules
{
    public class CopyInOtherFileRule
    {
        public string InputColumn { get; set; }
        public string OutputColumn { get; set; }

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
