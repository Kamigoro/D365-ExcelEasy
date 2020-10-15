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
        public string OutputFile { get; set; }
        public List<CopyRule> CopyRules { get; set; } = new List<CopyRule>();
        public EventHandler<RuleEventArgs> CopyRule_EventHandler;
        public List<ChangementRule> ChangementRules { get; set; } = new List<ChangementRule>();
        public EventHandler<RuleEventArgs> ChangementRule_EventHandler;

        #region Rules management

        public void AddCopyRule()
        {
            CopyRule rule = new CopyRule();
            rule.RuleExecuted_EventHandler += CopyRule_Executed;
            CopyRules.Add(rule);
        }

        private void CopyRule_Executed(object sender, RuleEventArgs e)
        {
            CopyRule_EventHandler?.Invoke(this, e);
        }

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

        public void RemoveCopyRule()
        {
            if (CopyRules.Count > 0)
            {
                CopyRules.RemoveAt(CopyRules.Count - 1);
            }
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
        public void ExecuteCopyRules()
        {
            foreach (CopyRule rule in CopyRules)
            {
                rule.InputFile = InputFile;
                rule.OutputFile = OutputFile;
                rule.Execute();
            }
        }

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
