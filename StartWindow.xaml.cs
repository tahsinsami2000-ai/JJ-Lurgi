using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;
using Aucotec.EngineeringBase.Client.Runtime;
// Alias EB Application to avoid clash
using EbApp = Aucotec.EngineeringBase.Client.Runtime.Application;

namespace JJ_Lurgi_Piping_EB
{
    public partial class StartWindow : Window
    {
        private readonly EbApp _app;
        private readonly IList<ItemRow> _pickedOrDb;

        public StartWindow(EbApp app, IList<ItemRow> pickedOrDb = null)
        {
            _app = app;                 // allow null for designer mode
            _pickedOrDb = pickedOrDb;   // can be null
            InitializeComponent();
        }

        // 1) Generate Component Datasheets
        private void BtnMod1_Click(object sender, RoutedEventArgs e)
        {
            Hide();
            try
            {
                MainWindow w = new MainWindow(_app) { Owner = this };
                w.ShowDialog();
            }
            catch (Exception err)
            {
                MessageBox.Show("Module 1 failed: " + err.Message, "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            Close();
        }

        // 2) Generate Pipe Class Summary (Module 2)
        private void BtnMod2_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                PipeClassWindow w = new PipeClassWindow(_app) { Owner = this };
                Hide();
                w.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "--- Pipe Class Summary ---\n\nFailed to open Module 2 window:\n" + ex.Message,
                    "Module 2", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                Show();
            }
        }

        // 3) Generate Valve Application Summary (stub)
        // 3) Generate Valve Application Summary (Module 3)
        // 3) Generate Valve Application Summary (was a stub)
        private void BtnMod3_Click(object sender, RoutedEventArgs e)
        {
            var ok = MessageBox.Show(
                "Open Valve Application Summary and choose what to generate?",
                "Confirm",
                MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (ok != MessageBoxResult.Yes) return;

            try
            {
                var w = new ValveAppWindow(_app) { Owner = this };
                Hide();
                w.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "--- Valve Application Summary ---\n\nFailed to open Module 3 window:\n" + ex.Message,
                    "Module 3", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                Show();
            }
        }



        // 4) Quit
        private void BtnQuit_Click(object sender, RoutedEventArgs e) => Close();

        private void StartWindow_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.D1 || e.Key == Key.NumPad1) { BtnMod1_Click(null, null); e.Handled = true; }
            else if (e.Key == Key.D2 || e.Key == Key.NumPad2) { BtnMod2_Click(null, null); e.Handled = true; }
            else if (e.Key == Key.D3 || e.Key == Key.NumPad3) { BtnMod3_Click(null, null); e.Handled = true; }
            else if (e.Key == Key.D4 || e.Key == Key.NumPad4 || e.Key == Key.Escape) { BtnQuit_Click(null, null); e.Handled = true; }
        }
    }
}
