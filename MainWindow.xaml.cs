using D365_ExcelModifier.Models;
using D365_ExcelModifier.Models.DocumentRules;
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
        private List<ValueChangementRule> ValueChangementRules = new List<ValueChangementRule>();
        private List<CopyInOtherFileRule> CopyInOtherFileRules = new List<CopyInOtherFileRule>();

        public MainWindow()
        {
            InitializeComponent();
            ValueChangementItems.ItemsSource = ValueChangementRules;
            CopyInOtherFileItems.ItemsSource = CopyInOtherFileRules;
        }

        #region Rules gui management

        #region ValueChangement Rules
        private void BTNAddValueChangementRule_Click(object sender, RoutedEventArgs e)
        {
            ValueChangementRules.Add(new ValueChangementRule());
            RefreshValueChangementRulesItem();
        }
        private void BTNRemoveValueChangementRule_Click(object sender, RoutedEventArgs e)
        {
            if (ValueChangementRules.Count > 0)
            {
                ValueChangementRules.RemoveAt(ValueChangementRules.Count - 1);
                RefreshValueChangementRulesItem();
            }
        }

        private void RefreshValueChangementRulesItem()
        {
            Dispatcher.BeginInvoke((Action)(() =>
            {
                ValueChangementItems.ItemsSource = null;
                ValueChangementItems.ItemsSource = ValueChangementRules;
            }));
        }
        #endregion

        #region CopyInOtherFile Rules
        private void BTNAddCopyInOtherFileRule_Click(object sender, RoutedEventArgs e)
        {
            CopyInOtherFileRules.Add(new CopyInOtherFileRule());
            RefreshCopyInOtherFileRulesItem();
        }
        private void BTNRemoveCopyInOtherFileRule_Click(object sender, RoutedEventArgs e)
        {
            if (ValueChangementRules.Count > 0)
            {
                CopyInOtherFileRules.RemoveAt(ValueChangementRules.Count - 1);
                RefreshCopyInOtherFileRulesItem();
            }
        }

        private void RefreshCopyInOtherFileRulesItem()
        {
            Dispatcher.BeginInvoke((Action)(() =>
            {
                CopyInOtherFileItems.ItemsSource = null;
                CopyInOtherFileItems.ItemsSource = CopyInOtherFileRules;
            }));
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
                BTNExecuteValueChangementRules.Foreground = new SolidColorBrush(Colors.Green);
                BTNExecuteValueChangementRules.IsEnabled = true;
            }
            else
            {
                BTNExecuteValueChangementRules.Foreground = new SolidColorBrush(Colors.Red);
                BTNExecuteValueChangementRules.IsEnabled = true;
            }
        }

        private void TBXOutputFile_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            if (File.Exists(TBXOutputFile.Text) && File.Exists(TBXInputFile.Text))
            {
                BTNExecuteCopyInOtherFileRules.Foreground = new SolidColorBrush(Colors.Green);
                BTNExecuteCopyInOtherFileRules.IsEnabled = true;
            }
            else
            {
                BTNExecuteCopyInOtherFileRules.Foreground = new SolidColorBrush(Colors.Red);
                BTNExecuteCopyInOtherFileRules.IsEnabled = true;
            }
        }

        #endregion

        #region Rules execution

        #region ValueChangement rules

        private void BTNExecuteValueChangementRules_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DocumentManager documentManager = new DocumentManager(ValueChangementRules, TBXInputFile.Text);
                documentManager.FinishedExecution_Event += FinishedChangementValue_EventHandler;
                Task.Run(() =>
                {
                    documentManager.ExecuteRules();
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }

        }

        private void FinishedChangementValue_EventHandler()
        {
            Dispatcher.Invoke(() =>
            {
                var messageQueue = new SnackbarMessageQueue(TimeSpan.FromMilliseconds(2000));
                ValueChangementSnackbar.MessageQueue = messageQueue;
                ValueChangementSnackbar.MessageQueue.Enqueue("Opération terminée");
            });
        }

        #endregion

        #region CopyInOtherFile rules

        private void BTNExecuteCopyInOtherFile_Click(object sender, RoutedEventArgs e)
        {

            DocumentManager documentManager = new DocumentManager(CopyInOtherFileRules, TBXInputFile.Text, TBXOutputFile.Text);
            documentManager.FinishedExecution_Event += FinishedCopyInOtherFile_EventHandler;
            Task.Run(() =>
            {
                documentManager.ExecuteRules();
            });
        }

        private void FinishedCopyInOtherFile_EventHandler()
        {
            Dispatcher.Invoke(() =>
            {
                var messageQueue = new SnackbarMessageQueue(TimeSpan.FromMilliseconds(2000));
                CopySnackbar.MessageQueue = messageQueue;
                CopySnackbar.MessageQueue.Enqueue("Opération terminée");
            });
        }

        #endregion

        #endregion


    }
}
