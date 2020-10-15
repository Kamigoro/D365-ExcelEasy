using D365_ExcelModifier.Models.DocumentRules;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace D365_ExcelModifier.ViewModels
{
    public class MainViewModel
    {
        public string InputFile { get; set; }
        public List<ChangementRule> ChangementRules { get; set; } = new List<ChangementRule>();
        public EventHandler<RuleEventArgs> ChangementRule_EventHandler;

        #region Rules management

        public void AddChangementRule()
        {
            ChangementRule rule = new ChangementRule();
            rule.RuleExecuted_EventHandler += ChangementRule_Executed;
            ChangementRules.Add(rule);
        }

        private void ChangementRule_Executed(object sender, RuleEventArgs e)
        {
            ChangementRule_EventHandler?.Invoke(this, e);
        }

        public void RemoveChangementRule()
        {
            if (ChangementRules.Count > 0)
            {
                ChangementRules.RemoveAt(ChangementRules.Count - 1);
            }
        }

        #endregion

        #region Rules execution

        public void ExecuteChangementRules()
        {
            foreach (ChangementRule rule in ChangementRules)
            {
                rule.InputFile = InputFile;
                rule.Execute();
            }
        }
        #endregion
    }
}
