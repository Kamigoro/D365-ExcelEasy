using ClosedXML.Excel;
using D365_ExcelModifier.Models;
using D365_ExcelModifier.Models.Actions;
using D365_ExcelModifier.Models.DocumentRules;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;

namespace D365_ExcelModifier
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        private List<DocumentRuleBase> DocumentRules = new List<DocumentRuleBase>();

        public MainWindow()
        {
            InitializeComponent();
            ReplacementItems.ItemsSource = DocumentRules;
        }

        #region Rules gui management

        private void BTNAddRule_Click(object sender, RoutedEventArgs e)
        {
            DocumentRules.Add(new DocumentRuleBase());
            RefreshRulesItem();
        }
        private void BTNRemoveRule_Click(object sender, RoutedEventArgs e)
        {
            DocumentRules.RemoveAt(DocumentRules.Count - 1);
            RefreshRulesItem();
        }

        private void RefreshRulesItem()
        {
            Dispatcher.BeginInvoke((Action)(() =>
            {
                ReplacementItems.ItemsSource = null;
                ReplacementItems.ItemsSource = DocumentRules;
            }));
        }

        #endregion

        #region Files picker

        private void BTNSearchInputFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Fichier xlsx (.xlsx)|*.xlsx";
            if (openFileDialog.ShowDialog() == true)
            {
                TBXInputFile.Text = openFileDialog.FileName;
            }
        }

        private void BTNSearchOutputFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Fichier xlsx (.xlsx)|*.xlsx";
            if (openFileDialog.ShowDialog() == true)
            {
                TBXOutputFile.Text = openFileDialog.FileName;
            }
        }

        //Warning the user that he can't execute rules without an existing file
        private void TBXInputFile_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            if (File.Exists(TBXInputFile.Text))
            {
                BTNExecuteRules.Foreground = new SolidColorBrush(Colors.Green);
                BTNExecuteRules.IsEnabled = true;
            }
            else
            {
                BTNExecuteRules.Foreground = new SolidColorBrush(Colors.Red);
                BTNExecuteRules.IsEnabled = false;
            }
        }

        #endregion

        #region Rules execution
        private void BTNExecuteRules_Click(object sender, RoutedEventArgs e)
        {
            /*using (var inputWorkbook = new XLWorkbook("input.xlsx"))
            {
                using (var outputWorkbook = new XLWorkbook("output.xlsx"))
                {
                    CopyInOtherFileRule rule = new CopyInOtherFileRule("InputColumn", "OutputColumn");
                    CopyInOtherFileAction action = new CopyInOtherFileAction(rule, inputWorkbook.Worksheets, outputWorkbook.Worksheets);
                    action.Execute();

                    ValueChangementRule rule2 = new ValueChangementRule("InputColumn", "Tu pues","Gros PD");
                    ValueChangingAction action2 = new ValueChangingAction(rule2, inputWorkbook.Worksheets);
                    action2.Execute();

                    outputWorkbook.Save();
                    inputWorkbook.Save();
                }
            }*/
        }
        #endregion
    }
}
