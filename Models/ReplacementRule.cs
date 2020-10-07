namespace D365_ExcelModifier.Models
{
    public class ReplacementRule
    {
        public string FromColumnOfInputfile { get; set; }
        public string ValueToFind { get; set; }
        public string ReplacementValue { get; set; }
        public string ToColumnOfOutputFile { get; set; }
    }
}
