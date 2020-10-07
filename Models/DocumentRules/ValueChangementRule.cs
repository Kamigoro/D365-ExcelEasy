namespace D365_ExcelModifier.Models.DocumentRules
{
    public class ValueChangementRule : IDocumentRule
    {
        public string TargetedColumn { get; set; }
        public string OldValue { get; set; }
        public string NewValue { get; set; }

        public ValueChangementRule(string targetedColumn, string oldValue, string newValue)
        {
            TargetedColumn = targetedColumn;
            OldValue = oldValue;
            NewValue = newValue;
        }
    }
}
