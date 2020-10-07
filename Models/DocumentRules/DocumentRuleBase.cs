namespace D365_ExcelModifier.Models.DocumentRules
{
    public class DocumentRuleBase
    {
        public string InputColumn { get; set; }
        public string OutputColumn { get; set; }
        public string OldValue { get; set; }
        public string NewValue { get; set; }
    }
}
