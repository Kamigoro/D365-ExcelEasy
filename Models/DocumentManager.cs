using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace D365_ExcelModifier.Models
{
    public class DocumentManager
    {

        public string InputFile { get; set; }
        public string OutputFile { get; set; }

        public DocumentManager(string inputFile)
        {
            InputFile = inputFile;
        }

        public DocumentManager(string inputFile, string outputFile)
        {
            InputFile = inputFile;
            OutputFile = outputFile;
        }

        public async void ExecuteRules(List<ReplacementRule> replacementRules)
        {
            using (var workbook = new XLWorkbook(InputFile))
            {
                foreach (var worksheet in workbook.Worksheets)//on va aller chercher dans toutes les feuilles
                {
                    foreach (ReplacementRule replacementRule in replacementRules)//Sur toutes les feuilles on va essayer d'appliquer chaque règle
                    {
                        await ExecuteRule(replacementRule, worksheet);
                    }
                }
                workbook.Save();
            }
        }


        private async Task<bool> ExecuteRule(ReplacementRule replacementRule, IXLWorksheet worksheet)
        {
            await Task.Yield();
            if (replacementRule.ToColumnOfOutputFile == null)
            {
                return ValueInTheSameColumnReplacement(replacementRule, worksheet);
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Take a replacement rule and apply it for a column in the same document.
        /// For example we have a column where we want to replace all the cells containing "false" by "true".
        /// </summary>
        /// <param name="replacementRule"></param>
        /// <param name="worksheet"></param>
        private bool ValueInTheSameColumnReplacement(ReplacementRule replacementRule, IXLWorksheet worksheet)
        {
            try
            {
                var columnSearched = worksheet.Columns()
                .Where(column => column.Cell(1)
                .Value
                .ToString()
                .StartsWith(replacementRule.FromColumnOfInputfile)).First();

                var correspondingCells = columnSearched.Cells().
                    Where(cell => cell.Value
                    .ToString()
                    .StartsWith(replacementRule.ValueToFind));
                foreach (var cell in correspondingCells)
                {
                    cell.Value = replacementRule.ReplacementValue;
                }

                return true;
            }
            catch (Exception e)
            {
                return false;
            }

        }

    }
}
