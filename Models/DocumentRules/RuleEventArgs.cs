using System;
using System.Buffers.Text;
using System.Collections.Generic;
using System.Text;

namespace D365_ExcelModifier.Models.DocumentRules
{
    public class RuleEventArgs : EventArgs
    {
        public BaseRule Rule { get; set; }
        public bool ExecutionStatus { get; set; }
        public string ErrorMessage { get; set; }

        public RuleEventArgs(BaseRule rule, bool executionStatus)
        {
            Rule = rule;
            ExecutionStatus = executionStatus;
        }

        public RuleEventArgs(BaseRule rule, bool executionStatus, string errorMessage)
        {
            Rule = rule;
            ExecutionStatus = executionStatus;
            ErrorMessage = errorMessage;
        }
    }
}
