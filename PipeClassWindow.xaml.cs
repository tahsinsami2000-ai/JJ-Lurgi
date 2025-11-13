#nullable enable
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using Aucotec.EngineeringBase.Client.Runtime;
using EbApp = Aucotec.EngineeringBase.Client.Runtime.Application;

namespace JJ_Lurgi_Piping_EB
{
    public partial class PipeClassWindow : Window
    {
        private readonly EbApp _app;
        private List<string> _all = new List<string>();
        private List<string> _asme = new List<string>();
        private List<string> _din = new List<string>();

        public PipeClassWindow(EbApp app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            InitializeComponent();
            Loaded += PipeClassWindow_Loaded;
        }

        private void PipeClassWindow_Loaded(object? sender, RoutedEventArgs e)
        {
            try
            {
                // Fetch and split on load (fast cached lists)
                _all = PipeClassServiceEb.FetchAllPipeClassesFromEb(_app);
                PipeClassServiceEb.SplitAsAsmeDin(_all, out _asme, out _din);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error fetching pipe classes: " + ex.Message);
            }

            var sb = new StringBuilder();
            sb.AppendLine("Please select a generation mode:");
            sb.AppendLine("  1. Generate for a single Pipe Class");
            sb.AppendLine("  2. Generate for all ASME Pipe Classes (ANSI)");
            sb.AppendLine("  3. Generate for all DIN Pipe Classes");
            sb.AppendLine("  4. Generate for ALL Pipe Classes");
            sb.AppendLine("  5. Back to Main Menu");
            sb.AppendLine("  6. Quit");
            sb.AppendLine();
            sb.Append("Enter your choice (1-6):");
            TxtMenu.Text = sb.ToString();
            TxtChoice.Focus();
        }

        private void Back_Click(object sender, RoutedEventArgs e) => Close();

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            var raw = (TxtChoice.Text ?? "").Trim();
            if (!int.TryParse(raw, out int choice))
            {
                MessageBox.Show("Please enter a number between 1–6.", "Pipe Class Menu");
                return;
            }

            try
            {
                switch (choice)
                {
                    case 1:
                        ShowSinglePicker();
                        break;

                    case 2:
                        MessageBox.Show("Generating all ASME (ANSI) Pipe Classes...", "Generating");
                        var msgAsme = PipeClassServiceEb.GenerateForClasses(_app, _asme);
                        MessageBox.Show(msgAsme, "Pipe Class Summary");
                        break;

                    case 3:
                        MessageBox.Show("Generating all DIN Pipe Classes...", "Generating");
                        var msgDin = PipeClassServiceEb.GenerateForClasses(_app, _din);
                        MessageBox.Show(msgDin, "Pipe Class Summary");
                        break;

                    case 4:
                        MessageBox.Show("Generating ALL Pipe Classes...", "Generating");
                        var msgAll = PipeClassServiceEb.GenerateForClasses(_app, _all);
                        MessageBox.Show(msgAll, "Pipe Class Summary");
                        break;

                    case 5:
                        Close();
                        break;

                    case 6:
                        System.Windows.Application.Current.Shutdown();
                        break;

                    default:
                        MessageBox.Show("Please enter a number 1–6.", "Pipe Class Menu");
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Pipe Class Menu");
            }
        }

        private void ShowSinglePicker()
        {
            var win = new PipeClassSinglePicker(_app, _all);
            win.Owner = this;
            win.ShowDialog();
        }
    }
}
#nullable disable
