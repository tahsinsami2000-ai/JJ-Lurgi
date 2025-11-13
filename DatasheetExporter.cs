#nullable enable
using System.Collections.Generic;
using Aucotec.EngineeringBase.Client.Runtime;
// Alias EB Application to avoid any ambiguity
using EbApp = Aucotec.EngineeringBase.Client.Runtime.Application;

namespace JJ_Lurgi_Piping_EB
{
    public static class DatasheetExporter
    {
        /// <summary>
        /// Called from MainWindow → Generate. Do NOT open StartWindow here.
        /// Generate datasheets for the picked items and return a status message.
        /// </summary>
        public static string Generate(EbApp app, IList<ItemRow> picked)
        {
            if (app == null)
                return "No EB application instance.";
            if (picked == null || picked.Count == 0)
                return "No items selected. Tick rows in the grid and try again.";

            var res = DatasheetServiceEb.GenerateForSelection(app, picked);
            return res != null ? (res.Message ?? "Done.") : "No result.";
        }
    }
}
#nullable disable
