using ClosedXML.Excel;
using D365_ExcelModifier.Models.DocumentRules;
using System;
using System.Linq;
using System.Threading.Tasks;

namespace D365_ExcelModifier.Models.Actions
{
    public class ValueChangingAction : IDocumentAction
    {
        public IXLWorksheet Worksheet { get; set; }
        public SimpleValueChangementRule Rule { get; set; }

        public ValueChangingAction(SimpleValueChangementRule rule)
        {
            Rule = rule;
        }

        public async Task<bool> ExecuteAsync()
        {
            if (Rule.TargetedColumn == "*")
            {
                return await ChangeForAllColumnsAsync();
            }
            else
            {
                return await ChangeForSpecificColumnAsync();
            }
        }

        /// <summary>
        /// Change the value of targeted by value cells by a new value in all the columns of a worksheet.
        /// </summary>
        /// <returns></returns>
        private async Task<bool> ChangeForAllColumnsAsync()
        {
            await Task.Yield();
            try
            {
                //We get all the cells from the sheet that starts with the old value
                var cellsToReplace = Worksheet.Cells().Where(cell => cell.Value
                .ToString()
                .StartsWith(Rule.OldValue));

                foreach (var cell in cellsToReplace)
                {
                    //we replace each cell with the new value
                    cell.Value = Rule.NewValue;
                }

                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }

        /// <summary>
        /// Get all the column specified in the rules and change their cells containing an old value with a new value;
        /// </summary>
        /// <returns></returns>
        private async Task<bool> ChangeForSpecificColumnAsync()
        {
            await Task.Yield();
            try
            {
                //Get the columns where the first cell value is specified in the rule
                var targetedColumns = Worksheet.Columns()
                    .Where(column => column.Cell(1).Value.ToString().StartsWith(Rule.TargetedColumn));

                foreach (var column in targetedColumns)
                {
                    //We find each cell that start with the old value
                    var cellsToReplace = column.Cells().Where(cell => cell.Value.ToString().StartsWith(Rule.OldValue));
                    foreach (var cell in cellsToReplace)
                    {
                        //we replace the old valu by the new value
                        cell.Value = Rule.NewValue;
                    }
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
