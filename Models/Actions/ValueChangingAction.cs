using ClosedXML.Excel;
using D365_ExcelModifier.Models.DocumentRules;
using System;
using System.Diagnostics;
using System.Linq;

namespace D365_ExcelModifier.Models.Actions
{
    public class ValueChangingAction
    {
        public IXLWorksheets InputWorksheets { get; set; }
        public ValueChangementRule Rule { get; set; }

        public ValueChangingAction(ValueChangementRule rule, IXLWorksheets inputWorksheets)
        {
            Rule = rule;
            InputWorksheets = inputWorksheets;
        }

        public bool Execute()
        {
            if (Rule.InputColumn == "*")
            {
                return ChangeForAllColumns();
            }
            else
            {
                return ChangeForSpecificColumn();
            }
        }

        /// <summary>
        /// Change the value of targeted by value cells by a new value in all the columns in all worksheets.
        /// </summary>
        /// <returns></returns>
        private bool ChangeForAllColumns()
        {
            Debug.WriteLine($"OldValue : {Rule.OldValue}");
            Debug.WriteLine($"NewValue : {Rule.NewValue}");
            try
            {
                foreach (var inputWorksheet in InputWorksheets)
                {
                    //We get all the cells from the sheet that starts with the old value
                    var cellsToReplace = inputWorksheet.Cells().Where(cell => cell.Value
                    .ToString()
                    .StartsWith(Rule.OldValue));

                    foreach (var cell in cellsToReplace)
                    {
                        //we replace each cell with the new value
                        cell.Value = Rule.NewValue;
                    }
                }
                return true;
            }
            catch (Exception e)
            {
                Debug.WriteLine(e);
                throw;
            }
        }

        /// <summary>
        /// Get all the column specified in the rules and change their cells containing an old value with a new value.
        /// </summary>
        /// <returns></returns>
        private bool ChangeForSpecificColumn()
        {
            try
            {
                foreach (var inputWorksheet in InputWorksheets)
                {
                    //Get the columns where the first cell value is specified in the rule
                    var targetedColumns = inputWorksheet.Columns()
                    .Where(column => column.Cell(1).Value.ToString().StartsWith(Rule.InputColumn));

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
                }
                return true;
            }
            catch (Exception e)
            {
                throw;
            }
        }

    }
}
