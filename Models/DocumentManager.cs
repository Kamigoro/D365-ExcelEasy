using D365_ExcelModifier.Models.DocumentRules;
using System;
using System.Collections.Generic;
using System.Text;

namespace D365_ExcelModifier.Models
{
    public class DocumentManager
    {
        private List<CopyInOtherFileRule> CopyInOtherFileRules { get; set; } = new List<CopyInOtherFileRule>();
        private List<ValueChangementRule> ValueChangementRules { get; set; } = new List<ValueChangementRule>();

        public DocumentManager(List<DocumentRuleBase> baseRules)
        {
            foreach (DocumentRuleBase baseRule in baseRules)
            {
                if (baseRule.OutputColumn == null)
                {
                    ValueChangementRules.Add((ValueChangementRule)baseRule);
                }
                else
                {
                    CopyInOtherFileRules.Add((CopyInOtherFileRule)baseRule);
                }
            }
        }
    }
}
