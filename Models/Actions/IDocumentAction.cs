using D365_ExcelModifier.Models.DocumentRules;
using System.Threading.Tasks;

namespace D365_ExcelModifier.Models.Actions
{
    public interface IDocumentAction
    {
        bool IsRuleValid(IDocumentRule documentRule);
        Task<bool> ExecuteAsync(IDocumentRule documentRule);
    }
}
