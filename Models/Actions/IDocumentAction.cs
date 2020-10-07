using ClosedXML.Excel;
using System.Threading.Tasks;

namespace D365_ExcelModifier.Models.Actions
{
    public interface IDocumentAction
    {
        IXLWorksheet Worksheet { get; set; }
        Task<bool> ExecuteAsync();
    }
}
