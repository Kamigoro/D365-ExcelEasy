namespace D365_ExcelModifier.Models.DocumentRules
{
    public class CopyInOtherFileRule : IDocumentRule
    {
        public string InputFile { get; set; }
        public string InputColumn { get; set; }
        public string OutputFile { get; set; }
        public string OutputColumn { get; set; }
    }
}
