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
        public List<Copyrule> CopyRules { get; set; } = new List<Copyrule>();
        public List<ChangementRule> ChangementRules { get; set; } = new List<ChangementRule>();

        #region Rules management

        public void AddCopyRule() 
        {
            CopyRules.Add(new Copyrule());
        }
        
        public void AddChangementRule() 
        {
            ChangementRules.Add(new ChangementRule());
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
            foreach (Copyrule rule in CopyRules)
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
