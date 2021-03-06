﻿using D365_ExcelModifier.Models;
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
        private double ChangementRulesProgressBarIncrement { get; set; }

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
            Dispatcher.Invoke(() =>
            {
                if (e.ExecutionStatus == false)
                {
                    ChangementRule rule = (ChangementRule)e.Rule;
                    string errorMessage = $"Règle de Changement de valeur\n Valeur à remplacer : {rule.OldValue}\t Valeur de remplacement : {rule.NewValue}\n Mesage d'erreur : {e.ErrorMessage}\n\n";
                    TBKErrorMessage.Text += errorMessage;
                }
                else
                {
                    ChangementRulesProgressBar.Value += ChangementRulesProgressBarIncrement;
                }
            });

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

        #endregion

        #region Rules execution

        private void BTNExecuteChangementRules_Click(object sender, RoutedEventArgs e)
        {
            ChangementRulesProgressBarIncrement = 100 / ViewModel.ChangementRules.Count;
            ChangementRulesProgressBar.Value = 0;
            Thread executionThread = new Thread(ViewModel.ExecuteChangementRules);
            executionThread.Start();
        }

        #endregion


    }
}
