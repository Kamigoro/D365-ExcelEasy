namespace D365_ExcelModifier.Models.DocumentRules
{
    public class ValueChangementRule : DocumentRuleBase
    {
        public ValueChangementRule(string inputColumn, string oldValue, string newValue)
        {
            InputColumn = inputColumn;
            OldValue = oldValue;
            NewValue = newValue;
        }
    }
}
