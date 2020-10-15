using D365_ExcelModifier.Models;
using D365_ExcelModifier.Models.DocumentRules;
using D365_ExcelModifier.ViewModels;
using MaterialDesignThemes.Wpf;
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
        public MainViewModel ViewModel { get; set; }

        public MainWindow()
        {
            ViewModel = new MainViewModel();
            ViewModel.ChangementRule_EventHandler += ChangementRule_Executed;
            ViewModel.CopyRule_EventHandler += CopyRule_Executed;
            DataContext = ViewModel;
            InitializeComponent();
        }



        #region Rules gui management

        #region ValueChangement Rules
        private void BTNAddValueChangementRule_Click(object sender, RoutedEventArgs e)
        {
            ViewModel.AddChangementRule();
            RefreshValueChangementRulesItem();
        }
        private void BTNRemoveValueChangementRule_Click(object sender, RoutedEventArgs e)
        {
            ViewModel.RemoveChangementRule();
            RefreshValueChangementRulesItem();
        }
        private void RefreshValueChangementRulesItem()
        {
            Dispatcher.BeginInvoke((Action)(() =>
            {
                ChangementItems.ItemsSource = null;
                ChangementItems.ItemsSource = ViewModel.ChangementRules;
            }));
        }
        #endregion

        #region CopyInOtherFile Rules
        private void BTNAddCopyInOtherFileRule_Click(object sender, RoutedEventArgs e)
        {
            ViewModel.AddCopyRule();
            RefreshCopyInOtherFileRulesItem();
        }
        private void BTNRemoveCopyInOtherFileRule_Click(object sender, RoutedEventArgs e)
        {
            ViewModel.RemoveCopyRule();
            RefreshCopyInOtherFileRulesItem();
        }
        private void RefreshCopyInOtherFileRulesItem()
        {
            Dispatcher.BeginInvoke((Action)(() =>
            {
                CopyItems.ItemsSource = null;
                CopyItems.ItemsSource = ViewModel.CopyRules;
            }));
        }
        #endregion

        #region Rule execution event

        private void CopyRule_Executed(object sender, RuleEventArgs e)
        {
            if (e.ExecutionStatus == false)
            {
                CopyRule rule = (CopyRule)e.Rule;
                string errorMessage = $"Règle de copie\n Colonne d'entrée : {rule.InputColumn}\t Colonne de sortie : {rule.OutputColumn}\n Mesage d'erreur : {e.ErrorMessage}\n\n";
                TBKErrorMessage.Text += errorMessage;
            }
        }

        private void ChangementRule_Executed(object sender, RuleEventArgs e)
        {
            if (e.ExecutionStatus == false)
            {
                ChangementRule rule = (ChangementRule)e.Rule;
                string errorMessage = $"Règle de Changement de valeur\n Valeur à remplacer : {rule.OldValue}\t Valeur de remplacement : {rule.NewValue}\n Mesage d'erreur : {e.ErrorMessage}\n\n";
                TBKErrorMessage.Text += errorMessage;
            }
        }

        #endregion

        #endregion

        #region Files picker

        private void BTNSearchInputFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Fichier xlsx (.xlsx)|*.xlsx"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                TBXInputFile.Text = openFileDialog.FileName;
            }
        }

        private void BTNSearchOutputFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Fichier xlsx (.xlsx)|*.xlsx"
            };
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
                BTNExecuteChangementRules.Foreground = new SolidColorBrush(Colors.Green);
                BTNExecuteChangementRules.IsEnabled = true;
                ViewModel.InputFile = TBXInputFile.Text;
            }
            else
            {
                BTNExecuteChangementRules.Foreground = new SolidColorBrush(Colors.Red);
                BTNExecuteChangementRules.IsEnabled = true;
            }
        }

        private void TBXOutputFile_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            if (File.Exists(TBXOutputFile.Text) && File.Exists(TBXInputFile.Text))
            {
                BTNExecuteCopyRules.Foreground = new SolidColorBrush(Colors.Green);
                BTNExecuteCopyRules.IsEnabled = true;
                ViewModel.OutputFile = TBXOutputFile.Text;
            }
            else
            {
                BTNExecuteCopyRules.Foreground = new SolidColorBrush(Colors.Red);
                BTNExecuteCopyRules.IsEnabled = true;
            }
        }

        #endregion

        #region Rules execution

        #region ValueChangement rules

        private void BTNExecuteChangementRules_Click(object sender, RoutedEventArgs e)
        {
            ViewModel.ExecuteChangementRules();
        }

        #endregion

        #region CopyInOtherFile rules

        private void BTNExecuteCopyRules_Click(object sender, RoutedEventArgs e)
        {

            ViewModel.ExecuteCopyRules();
        }

        #endregion

        #endregion


    }
}
