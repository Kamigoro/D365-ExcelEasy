using D365_ExcelModifier.Models;
using D365_ExcelModifier.Models.DocumentRules;
using D365_ExcelModifier.ViewModels;
using MaterialDesignThemes.Wpf;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
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

        #region Rule execution event

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

        #endregion

        #region Rules execution

        #region ValueChangement rules

        private void BTNExecuteChangementRules_Click(object sender, RoutedEventArgs e)
        {
            ViewModel.ExecuteChangementRules();
        }

        #endregion

        #endregion


    }
}
