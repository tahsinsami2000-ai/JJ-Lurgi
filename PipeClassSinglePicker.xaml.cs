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
    public partial class PipeClassSinglePicker : Window
    {
        private readonly EbApp _app;
        private readonly List<string> _all;
        private readonly List<string> _flat = new List<string>(); // 1..N → code

        public PipeClassSinglePicker(EbApp app, IEnumerable<string> all)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            _all = new List<string>(all ?? Enumerable.Empty<string>());
            InitializeComponent();
            Loaded += PipeClassSinglePicker_Loaded;
        }

        private void PipeClassSinglePicker_Loaded(object? sender, RoutedEventArgs e)
        {
            List<string> asme;
            List<string> din;
            PipeClassServiceEb.SplitAsAsmeDin(_all, out asme, out din);

            var sb = new StringBuilder();
            sb.AppendLine("Please select a Pipe Class to generate (Enter 'b' to go back):");
            sb.AppendLine();

            int counter = 1;

            // ======== ASME ========
            if (asme.Count > 0)
            {
                sb.AppendLine("========== ASME ==========");
                var grouped = PipeClassServiceEb.GroupByRating(asme);

                if (grouped.Asme.Count > 0)
                {
                    sb.AppendLine("--- ASME Ratings --------------------");
                    sb.Append(FormatThreeColumns(grouped.Asme, ref counter));
                }
                if (grouped.Din.Count > 0)
                {
                    sb.AppendLine("--- DIN Ratings --------------------");
                    sb.Append(FormatThreeColumns(grouped.Din, ref counter));
                }
                if (grouped.Other.Count > 0)
                {
                    sb.AppendLine("--- Other Ratings --------------------");
                    sb.Append(FormatThreeColumns(grouped.Other, ref counter));
                }

                sb.AppendLine();
            }

            // ======== DIN ========
            if (din.Count > 0)
            {
                sb.AppendLine("========== DIN ==========");
                var grouped = PipeClassServiceEb.GroupByRating(din);

                if (grouped.Asme.Count > 0)
                {
                    sb.AppendLine("--- ASME Ratings --------------------");
                    sb.Append(FormatThreeColumns(grouped.Asme, ref counter));
                }
                if (grouped.Din.Count > 0)
                {
                    sb.AppendLine("--- DIN Ratings --------------------");
                    sb.Append(FormatThreeColumns(grouped.Din, ref counter));
                }
                if (grouped.Other.Count > 0)
                {
                    sb.AppendLine("--- Other Ratings --------------------");
                    sb.Append(FormatThreeColumns(grouped.Other, ref counter));
                }

                sb.AppendLine();
            }

            TxtList.Text = sb.ToString();
            TxtChoice.Focus();
        }

        private string FormatThreeColumns(IList<string> items, ref int counter)
        {
            // 3 columns, width ~ 24 chars each
            var sb = new StringBuilder();
            int width = 24;
            for (int i = 0; i < items.Count; i += 3)
            {
                var a = (i + 0 < items.Count) ? items[i + 0] : null;
                var b = (i + 1 < items.Count) ? items[i + 1] : null;
                var c = (i + 2 < items.Count) ? items[i + 2] : null;

                if (a != null) { _flat.Add(a); sb.Append(counter.ToString().PadLeft(3)).Append(". ").Append(a.PadRight(width)); counter++; }
                if (b != null) { _flat.Add(b); sb.Append(counter.ToString().PadLeft(3)).Append(". ").Append(b.PadRight(width)); counter++; }
                if (c != null) { _flat.Add(c); sb.Append(counter.ToString().PadLeft(3)).Append(". ").Append(c); counter++; }
                sb.AppendLine();
            }
            return sb.ToString();
        }

        private void Back_Click(object sender, RoutedEventArgs e) => Close();

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            var raw = (TxtChoice.Text ?? "").Trim().ToLowerInvariant();
            if (raw == "b") { Close(); return; }

            if (!int.TryParse(raw, out int idx) || idx < 1 || idx > _flat.Count)
            {
                MessageBox.Show("Enter a number between 1 and " + _flat.Count + " or 'b' to go back.", "Pipe Class");
                return;
            }

            var chosen = _flat[idx - 1];
            var msg = PipeClassServiceEb.GenerateForClasses(_app, new string[] { chosen });
            MessageBox.Show(msg, "Pipe Class Summary");
        }
    }
}
#nullable disable
