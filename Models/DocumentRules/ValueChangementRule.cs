namespace D365_ExcelModifier.Models.DocumentRules
{
    public class ValueChangementRule
    {
        public string InputColumn { get; set; }
        public string OldValue { get; set; }
        public string NewValue { get; set; }

        public ValueChangementRule()
        {

        }

        public ValueChangementRule(string inputColumn, string oldValue, string newValue)
        {
            InputColumn = inputColumn;
            OldValue = oldValue;
            NewValue = newValue;
        }
    }
}
