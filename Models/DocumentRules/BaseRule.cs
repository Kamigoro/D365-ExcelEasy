using System;
using System.Collections.Generic;
using System.Text;

namespace D365_ExcelModifier.Models.DocumentRules
{
    public abstract class BaseRule
    {
        public EventHandler<RuleEventArgs> RuleExecuted_EventHandler;

        public abstract void Execute();
    }
}
