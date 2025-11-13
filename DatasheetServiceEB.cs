using Aucotec.EngineeringBase.Client.Runtime;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace JJ_Lurgi_Piping_EB
{
    /// <summary>
    /// Aggregated workbook generator (PP/Gasket/Valve/Bolts), no macros.
    /// One workbook per category, one sheet per selected item.
    /// </summary>
    public static class DatasheetServiceEb
    {
        // ======= Project folders (FIXED to your current location) =======
        private static readonly string ProjectRoot = @"E:\Aucotec Developer\Tahsin\JJ_Lurgi_Piping_EB";
        private static readonly string TemplatesRoot = Path.Combine(ProjectRoot, "templates");
        private static readonly string OutputRoot = Path.Combine(ProjectRoot, "output");
        // ================================================================

        // ---------- Attribute-key maps (PP / Gasket unchanged skeletons) ----------
        private static readonly Dictionary<string, string> MapPP =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "Code", "Code" }, { "Type", "Type" }, { "Seamless / Welded", "Seamless / Welded" },
                { "Material", "Material" }, { "Acc to Standard", "Acc to Standard" }, { "Schedule", "Schedule" },
                { "Class", "Class" }, { "Rating", "Rating" }, { "Length", "Length" },
                { "Additional info 1", "Additional info 1" }, { "Additional info 2", "Additional info 2" },
                { "Piping connection 1", "Piping connection 1" }, { "Piping connection 2", "Piping connection 2" },
                { "Colour marking 1", "Colour marking 1" }, { "Colour marking 2", "Colour marking 2" },
                { "Piping class", "Piping class" }, { "Size min", "Size min" }, { "Size max", "Size max" },
                { "Odd sizes allowed", "Odd sizes allowed" }
            };

        private static readonly Dictionary<string, string> MapG =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                // kept for fallback lookups when Description (C1/C2) isn’t present in EB
                { "Code", "Code" }, { "Type", "Type" }, { "Inside / Outside Ring", "Inside / Outside Ring" },
                { "Material 1 column 1", "Material 1 column 1" }, { "Material 1 column 2", "Material 1 column 2" },
                { "Material 2 column 1", "Material 2 column 1" }, { "Material 2 column 2", "Material 2 column 2" },
                { "Material 3 column 1", "Material 3 column 1" }, { "Material 3 column 2", "Material 3 column 2" },
                { "Material 4 column 1", "Material 4 column 1" }, { "Material 4 column 2", "Material 4 column 2" },
                { "Acc to Standard", "Acc to Standard" }, { "Class", "Class" },
                { "Flange Facing", "Flange Facing" }, { "Thickness", "Thickness" },
                { "Design pressure", "Design pressure" }, { "Design temperature", "Design temperature" },
                { "Piping class", "Piping class" }, { "Size min", "Size min" }, { "Size max", "Size max" },
                { "Odd sizes allowed", "Odd sizes allowed" }
            };

        // ---- BOLTS mapping (unchanged) ----
        private static readonly Dictionary<string, string> MapB =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "Code", "Device designation" },
                { "Type", "Additional Comment" },
                { "Material number", "Material number" },
                { "Nut Material", "Nut Material" },
                { "Coating", "Coating" },
                { "Material Standard", "Material Standard" },
                { "Remark 1", "Remark 1" }
            };

        private enum Category { PipingPart, Gasket, Valve, Bolt }

        // Canonical NPS sizes
        private static readonly List<string> NpsSizes = new List<string>
        {
            "1/2","3/4","1","1 1/4","1 1/2","2","2 1/2","3","4","5","6",
            "8","10","12","14","16","18","20","22","24","26","28","30","32","34","36"
        };

        public static DatasheetResult GenerateForSelection(Application app, IList<ItemRow> picked)
        {
            if (picked == null || picked.Count == 0)
                return new DatasheetResult { Created = 0, Message = "No items selected." };

            if (!Directory.Exists(TemplatesRoot))
                return new DatasheetResult { Created = 0, Message = "Templates folder not found at:\n" + TemplatesRoot };

            Directory.CreateDirectory(OutputRoot);

            // Template files (exact names)
            string ppTemplate = Path.Combine(TemplatesRoot, "Piping parts data sheet template - for program.xlsx");
            string gkTemplate = Path.Combine(TemplatesRoot, "Piping gasket data sheet template - for program.xlsx");
            string vlTemplate = Path.Combine(TemplatesRoot, "Valve data sheet template - for program.xlsx");
            string btTemplate = Path.Combine(TemplatesRoot, "Bolts data sheet template - for program.xlsx");

            // Output files
            string outPP = Path.Combine(OutputRoot, "Piping parts data sheets.xlsx");
            string outGK = Path.Combine(OutputRoot, "Gasket data sheets.xlsx");
            string outVL = Path.Combine(OutputRoot, "Valve data sheets.xlsx");
            string outBT = Path.Combine(OutputRoot, "Bolts data sheets.xlsx");

            // Copy templates → outputs
            if (File.Exists(ppTemplate)) File.Copy(ppTemplate, outPP, true);
            if (File.Exists(gkTemplate)) File.Copy(gkTemplate, outGK, true);
            if (File.Exists(vlTemplate)) File.Copy(vlTemplate, outVL, true);
            if (File.Exists(btTemplate)) File.Copy(btTemplate, outBT, true);

            // Open workbooks
            XLWorkbook wbPP = File.Exists(outPP) ? new XLWorkbook(outPP) : null;
            XLWorkbook wbGK = File.Exists(outGK) ? new XLWorkbook(outGK) : null;
            XLWorkbook wbVL = File.Exists(outVL) ? new XLWorkbook(outVL) : null;
            XLWorkbook wbBT = File.Exists(outBT) ? new XLWorkbook(outBT) : null;

            int created = 0;
            var skippedBecauseMissingValveTemplate = new List<string>();

            foreach (ItemRow row in picked)
            {
                string typeHint = FirstNonEmpty(
                    Attr(row, "Type"),
                    Attr(row, "Valve type"),
                    Attr(row, "Additional Comment"),
                    Attr(row, "Specification"),
                    Attr(row, "Comment"),
                    row.Name
                );

                Category cat = ClassifyFromText(typeHint, row);

                string sheetNameSeed = FirstNonEmpty(
                    Attr(row, "Code"),
                    Attr(row, "Device designation"),
                    Attr(row, "Comment"),
                    row.Name
                );
                string sheetName = SafeSheetName(sheetNameSeed);

                if (cat == Category.Valve)
                {
                    if (wbVL != null && File.Exists(vlTemplate))
                    {
                        var tmpl = FindTemplateSheet(wbVL, "Template-V", 1);
                        var ws = tmpl.CopyTo(sheetName);
                        PopulateValve(ws, row);
                        created++;
                    }
                    else
                    {
                        skippedBecauseMissingValveTemplate.Add(sheetName);
                    }
                    continue;
                }

                if (cat == Category.Gasket && wbGK != null && File.Exists(gkTemplate))
                {
                    var tmpl = FindTemplateSheet(wbGK, "Template-G", 1);
                    var ws = tmpl.CopyTo(sheetName);
                    PopulateGasket(ws, row);
                    created++;
                }
                else if (cat == Category.Bolt && wbBT != null && File.Exists(btTemplate))
                {
                    var tmpl = FindTemplateSheet(wbBT, "Template-B", 1);
                    var ws = tmpl.CopyTo(sheetName);
                    PopulateBolts(ws, row);
                    created++;
                }
                else if (wbPP != null && File.Exists(ppTemplate))
                {
                    var tmpl = FindTemplateSheet(wbPP, "Template-PP", 1);
                    var ws = tmpl.CopyTo(sheetName);
                    PopulatePipingPart(ws, row);
                    created++;
                }
            }

            if (wbPP != null) { TryDelete(wbPP, "Template-PP"); wbPP.Save(); }
            if (wbGK != null) { TryDelete(wbGK, "Template-G"); wbGK.Save(); }
            if (wbVL != null) { TryDelete(wbVL, "Template-V"); wbVL.Save(); }
            if (wbBT != null) { TryDelete(wbBT, "Template-B"); wbBT.Save(); }

            string msg = created > 0
                ? $"Generated {created} datasheet(s).\nOutput: {OutputRoot}"
                : "No datasheets generated.";

            if (skippedBecauseMissingValveTemplate.Count > 0)
            {
                msg += "\n\nValve items skipped (missing Valve template/workbook):\n - " +
                       string.Join("\n - ", skippedBecauseMissingValveTemplate);
            }

            return new DatasheetResult { Created = created, Message = msg };
        }

        // ---------- Populate: Piping Parts (UPDATED) ----------
        private static void PopulatePipingPart(IXLWorksheet ws, ItemRow it)
        {
            // CODE (EB shows it in "Comment" for your pipe items)
            string code = FirstNonEmpty(
                Attr(it, "Comment"),
                Attr(it, "Device designation"),
                Attr(it, "Code"),
                it?.Name
            );

            // TYPE from Additional Comment (e.g., LINED PIPE)
            string type = FirstNonEmpty(
                Attr(it, "Additional Comment"),
                Attr(it, "Type"),
                Attr(it, "Specification")
            );

            // SEAMLESS / WELDED
            string seamWeld = Attr(it, "Seamless / Welded");
            if (string.IsNullOrWhiteSpace(seamWeld))
            {
                var bft = Attr(it, "Body/Fitting type");
                if (!string.IsNullOrWhiteSpace(bft))
                {
                    var up = bft.ToUpperInvariant();
                    if (up.Contains("SEAMLESS")) seamWeld = "SEAMLESS";
                    else if (up.Contains("WELDED")) seamWeld = "WELDED";
                    else seamWeld = bft;
                }
            }

            // MATERIAL & STANDARD & SCHEDULE
            string material = FirstNonEmpty(
                Attr(it, "Material number"),
                Attr(it, "Material")
            );
            string accToStd = FirstNonEmpty(
                Attr(it, "Material Standard"),
                Attr(it, "Acc to Standard"), Attr(it, "ACC. TO STANDARD")
            );
            string schedule = FirstNonEmpty(
                Attr(it, "Pipe schedule no"),
                Attr(it, "Schedule")
            );

            // CLASS / RATING
            string rating = Attr(it, "Rating");
            string cls = FirstNonEmpty(Attr(it, "Class"), rating);

            // LENGTH
            string length = Attr(it, "Length");
            if (string.IsNullOrWhiteSpace(length))
            {
                var bft = Attr(it, "Body/Fitting type");
                if (!string.IsNullOrWhiteSpace(bft) && bft.ToUpperInvariant().Contains("LENGTH"))
                    length = bft;
            }

            // PIPING CONNECTIONS (two rows under the label)
            string pc1 = Attr(it, "Piping Connection 1");
            string pc2 = Attr(it, "Piping Connection 2");
            if (!string.IsNullOrWhiteSpace(pc1) || !string.IsNullOrWhiteSpace(pc2))
            {
                Set(ws, "D18", "PIPING CONNECTIONS");
                if (!string.IsNullOrWhiteSpace(pc1)) Set(ws, "G18", pc1);
                if (!string.IsNullOrWhiteSpace(pc2)) Set(ws, "G19", pc2);
            }

            // ADDITIONAL SPECIFICATION (if you store these)
            string addSpec1 = FirstNonEmpty(Attr(it, "Additional specification 1"), Attr(it, "Additional Specification 1"));
            string addSpec2 = FirstNonEmpty(Attr(it, "Additional specification 2"), Attr(it, "Additional Specification 2"));
            if (!string.IsNullOrWhiteSpace(addSpec1) || !string.IsNullOrWhiteSpace(addSpec2))
            {
                Set(ws, "D20", "ADDITIONAL SPECIFICATION");
                if (!string.IsNullOrWhiteSpace(addSpec1)) Set(ws, "G20", addSpec1);
                if (!string.IsNullOrWhiteSpace(addSpec2)) Set(ws, "G21", addSpec2);
            }

            // COLOUR MARKING
            string c1 = Attr(it, "Color Mark 1");
            string c2 = Attr(it, "Color Mark 2");

            // PIPING CLASS
            string pipingClass = FirstNonEmpty(
                Attr(it, "JLE Pipe Class"),
                Attr(it, "JLE Pipe Class 1"),
                Attr(it, "Piping class")
            );

            // ===== Write to template cells (JJ Lurgi PP template layout) =====
            Set(ws, "G7", code);
            Set(ws, "G9", type);

            if (!string.IsNullOrWhiteSpace(seamWeld))
            {
                Set(ws, "G10", seamWeld);
                Set(ws, "D10", "SEAMLESS/WELDED");
            }

            Set(ws, "G12", material);
            Set(ws, "G13", accToStd);
            if (!string.IsNullOrWhiteSpace(schedule))
            {
                Set(ws, "G14", schedule);
                Set(ws, "D14", "SCHEDULE");
            }

            if (!string.IsNullOrWhiteSpace(cls))
            {
                Set(ws, "G15", cls);
                Set(ws, "D15", string.IsNullOrWhiteSpace(rating) ? "CLASS" : "RATING");
            }

            if (!string.IsNullOrWhiteSpace(length))
            {
                Set(ws, "G16", length);
                Set(ws, "D16", "LENGTH");
            }

            if (!string.IsNullOrWhiteSpace(c1) || !string.IsNullOrWhiteSpace(c2))
            {
                Set(ws, "D21", "COLOUR MARKING");
                ws.Cell("F21").Value = "1ST";
                ws.Cell("F22").Value = "2ND";
                Set(ws, "G21", c1);
                Set(ws, "G22", c2);
                ws.Cell("D23").Value = "*along entire length of item";
            }

            Set(ws, "G24", pipingClass);

            Set(ws, "D4", "FITTING SPECIFICATION"); // some PP templates use this title row
                                                    // ===== NEW: Write sub-category heading =====
            string subcat = DetectPipingSubCategory(it);

            if (!string.IsNullOrWhiteSpace(subcat))
            {
                ws.Cell("D5").Value = subcat;   // bold heading under title
                ws.Cell("D5").Style.Font.Bold = true;
            }

            Set(ws, "D57", "1.0");

            // ===== Sizes table (auto-detect header row, MIN = 29) =====
            int dataStartRow = DetectSizeTableStartRow(ws, "Size", 29);
            string sMin = FirstNonEmpty(Attr(it, "Size (Min)"), Attr(it, "Size min"), Attr(it, "Size Min"));
            string sMax = FirstNonEmpty(Attr(it, "Size (Max)"), Attr(it, "Size max"), Attr(it, "Size Max"));

            if (!string.IsNullOrWhiteSpace(sMin) && !string.IsNullOrWhiteSpace(sMax))
            {
                string smin = NormalizeInch(sMin);
                string smax = NormalizeInch(sMax);

                int i1 = NpsSizes.FindIndex(x => string.Equals(NormalizeInch(x), smin, StringComparison.OrdinalIgnoreCase));
                int i2 = NpsSizes.FindIndex(x => string.Equals(NormalizeInch(x), smax, StringComparison.OrdinalIgnoreCase));
                if (i1 >= 0 && i2 >= 0)
                {
                    if (i2 < i1) { var t = i1; i1 = i2; i2 = t; }

                    int r = dataStartRow;
                    for (int i = i1; i <= i2; i++, r++)
                    {
                        string inch = NpsSizes[i];

                        // Columns: D size, E qty required, F qty ordered, G unit price, H total
                        ws.Cell(r, 4).Value = inch;            // D: Size text

                        // E (Quantity required) – leave untouched

                        // F: Quantity ordered → 0.00 (keep template formatting)
                        var qty = ws.Cell(r, 6);
                        qty.Value = 0.0;
                        // If the template is General, ensure decimals:
                        // qty.Style.NumberFormat.Format = "0.00";

                        // G: Unit Price → clear contents only (preserve shading/borders/numfmt)
                        ws.Cell(r, 7).Clear(XLClearOptions.Contents);

                        // H: Total = G * F
                        var total = ws.Cell(r, 8);
                        if (total.IsEmpty() && !total.HasFormula)
                            total.FormulaA1 = $"IFERROR({ws.Cell(r, 7).Address.ToStringRelative()}*{ws.Cell(r, 6).Address.ToStringRelative()},0)";
                    }
                }
            }
        }


        // Find the row below the “Size” header; fallback to 29 if not found.
        // Find the row below the “Size” header; never start before row 29 (to avoid overlapping headers)
        private static int DetectSizeTableStartRow(IXLWorksheet ws, string headerText, int minRow = 29)
        {
            var hdr = FindCellByText(ws, headerText);
            int candidate = hdr != null ? hdr.Address.RowNumber + 1 : minRow;
            return candidate < minRow ? minRow : candidate;
        }


        // ---------- Populate: Gasket (FINAL, EB-compatible) ----------
        private static void PopulateGasket(IXLWorksheet ws, ItemRow it)
        {
            string Clean(string s)
            {
                if (string.IsNullOrWhiteSpace(s)) return string.Empty;
                s = s.ToUpperInvariant();
                s = s.Replace("BAR(G)", "")
                     .Replace("BAR (G)", "")
                     .Replace("(G)", "")
                     .Replace("BARG", "")
                     .Replace("BAR", "")
                     .Replace("°C", "")
                     .Replace("DEGC", "")
                     .Trim();
                return s;
            }

            string AttrLike(string keyPart)
            {
                if (it?.Attributes == null || string.IsNullOrWhiteSpace(keyPart)) return "";
                foreach (var kv in it.Attributes)
                {
                    if (kv.Key != null &&
                        kv.Key.IndexOf(keyPart, StringComparison.OrdinalIgnoreCase) >= 0 &&
                        !string.IsNullOrWhiteSpace(kv.Value))
                    {
                        return kv.Value;
                    }
                }
                return "";
            }

            string code = FirstNonEmpty(
                Attr(it, "Device designation"),
                Attr(it, "Code"),
                Attr(it, "Comment"),
                it.Name
            );

            string typeBase = FirstNonEmpty(
                Attr(it, "Additional Comment"),
                Attr(it, "Type"),
                Attr(it, "Specification"),
                Attr(it, "Comment")
            );
            string ringInfo = Attr(it, "Gasket Inside/Outside Ring");
            string typeText = string.IsNullOrWhiteSpace(ringInfo) ? typeBase : $"{typeBase}\nWITH {ringInfo.ToUpper()}";

            string m1L = FirstNonEmpty(Attr(it, "Description 1 (C1)"), Attr(it, "Material 1 column 1"));
            string m1R = FirstNonEmpty(Attr(it, "Description 1 (C2)"), Attr(it, "Material 1 column 2"));
            string m2L = FirstNonEmpty(Attr(it, "Description 2 (C1)"), Attr(it, "Material 2 column 1"));
            string m2R = FirstNonEmpty(Attr(it, "Description 2 (C2)"), Attr(it, "Material 2 column 2"));
            string m3L = FirstNonEmpty(Attr(it, "Description 3 (C1)"), Attr(it, "Material 3 column 1"));
            string m3R = FirstNonEmpty(Attr(it, "Description 3 (C2)"), Attr(it, "Material 3 column 2"));
            string m4L = FirstNonEmpty(Attr(it, "Description 4 (C1)"), Attr(it, "Material 4 column 1"));
            string m4R = FirstNonEmpty(Attr(it, "Description 4 (C2)"), Attr(it, "Material 4 column 2"));

            string accStd = FirstNonEmpty(Attr(it, "Material Standard"), Attr(it, "Acc to Standard"), Attr(it, "ACC. TO STANDARD"));
            string cls = Attr(it, "Class");
            string facing = FirstNonEmpty(Attr(it, "Facing"), Attr(it, "Flange Facing"));
            string thickness = Attr(it, "Thickness");

            string dpMinRaw = FirstNonEmpty(
                Attr(it, "JLE Design pressure min"),
                Attr(it, "JLE Design pressure Min"),
                AttrLike("JLE Design pressure min"),
                AttrLike("Design pressure min")
            );

            string dpMaxRaw = FirstNonEmpty(
                Attr(it, "JLE Design pressure max"),
                Attr(it, "JLE Design pressure Max"),
                AttrLike("JLE Design pressure max"),
                AttrLike("Design pressure max")
            );

            string dtRaw = FirstNonEmpty(
                Attr(it, "JLE Design temperature max"),
                Attr(it, "JLE Design Temperature max"),
                Attr(it, "Design temperature"),
                Attr(it, "Design Temperature"),
                AttrLike("Design temperature")
            );

            string dpMin = Clean(dpMinRaw);
            string dpMax = Clean(dpMaxRaw);
            string dt = Clean(dtRaw);

            if (!string.IsNullOrEmpty(dpMax) && !dpMax.StartsWith("-") && !dpMax.StartsWith("+"))
                dpMax = "+" + dpMax;

            string dp;
            if (!string.IsNullOrEmpty(dpMin) && !string.IsNullOrEmpty(dpMax))
                dp = dpMin + " / " + dpMax + " BARG";
            else
                dp = (dpMin + dpMax).Trim();
            if (!string.IsNullOrEmpty(dp)) dp += " BARG";
            if (!string.IsNullOrEmpty(dt)) dt += " °C";

            Set(ws, "G7", code);
            Set(ws, "G9", typeText);

            Set(ws, "G12", m1L); Set(ws, "I12", m1R);
            Set(ws, "G13", m2L); Set(ws, "I13", m2R);
            Set(ws, "G14", m3L); Set(ws, "I14", m3R);
            Set(ws, "G15", m4L); Set(ws, "I15", m4R);

            Set(ws, "G16", accStd);
            Set(ws, "G17", cls);
            Set(ws, "G18", facing);
            Set(ws, "G19", thickness);

            if (!string.IsNullOrWhiteSpace(dp))
            {
                ws.Cell("G21").Value = dp;
                ws.Cell("D21").Value = "DESIGN PRESSURE";
            }

            if (!string.IsNullOrWhiteSpace(dt))
            {
                ws.Cell("G22").Value = dt;
                ws.Cell("D22").Value = "DESIGN TEMPERATURE";
            }

            Set(ws, "D4", "GASKET DATA SHEET");
            Set(ws, "D57", "1.0");
            int gkStart = DetectSizeTableStartRow(ws, "Size");
            if (gkStart < 29) gkStart = 29; // safe fallback so we start at D29 minimum
            WriteGasketSizesDn(ws, it, gkStart);

        }

        // ---------- Populate: Valve (IMPROVED MATERIAL MAPPING + endConn/addSpec restored) ----------
        private static void PopulateValve(IXLWorksheet ws, ItemRow it)
        {
            string AttrLike(string keyPart)
            {
                if (it?.Attributes == null || string.IsNullOrWhiteSpace(keyPart)) return "";
                foreach (var kv in it.Attributes)
                {
                    if (kv.Key != null &&
                        kv.Key.IndexOf(keyPart, StringComparison.OrdinalIgnoreCase) >= 0 &&
                        !string.IsNullOrWhiteSpace(kv.Value))
                    {
                        return kv.Value;
                    }
                }
                return "";
            }

            // --- Code / basic texts -------------------------------------------------
            string code = FirstNonEmpty(
                Attr(it, "Code"),
                Attr(it, "Device designation"),
                Attr(it, "Tag"),
                Attr(it, "Valve code"),
                Attr(it, "Comment"),
                it.Name
            );

            string vtype = FirstNonEmpty(
                Attr(it, "Additional Comment"),
                Attr(it, "Valve Type"),
                Attr(it, "Valve type"),
                Attr(it, "Type"),
                Attr(it, "Specification"),
                Attr(it, "Comment")
            );

            // MEDIUM & CORROSIVE
            string medium = FirstNonEmpty(
                Attr(it, "Fluid Name"),
                AttrLike("Fluid Name"),
                Attr(it, "Medium"),
                Attr(it, "Media"),
                Attr(it, "Fluid"),
                Attr(it, "Service")
            );

            string corrosive = FirstNonEmpty(
                Attr(it, "Corrosive Component"),
                AttrLike("Corrosive Component"),
                Attr(it, "Corrosive component"),
                Attr(it, "Corrosive"),
                Attr(it, "Corrosion component"),
                Attr(it, "Corrosive media")
            );

            // --- Design pressure / temperature --------------------------------------
            string dpMin = FirstNonEmpty(
                Attr(it, "JLE Design pressure min"),
                AttrLike("JLE Design pressure min"),
                AttrLike("Design pressure min")
            );
            string dpMax = FirstNonEmpty(
                Attr(it, "JLE Design pressure max"),
                AttrLike("JLE Design pressure max"),
                AttrLike("Design pressure max")
            );

            string dp1 = string.Empty;
            if (!string.IsNullOrWhiteSpace(dpMin) || !string.IsNullOrWhiteSpace(dpMax))
            {
                string left = (dpMin ?? string.Empty).Trim();
                string right = (dpMax ?? string.Empty).Trim();
                dp1 = (!string.IsNullOrEmpty(left) && !string.IsNullOrEmpty(right)) ? left + " / " + right : left + right;
            }

            string dt1 = FirstNonEmpty(
                Attr(it, "JLE Design temperature max"),
                AttrLike("JLE Design temperature max"),
                Attr(it, "Design temperature 1"),
                Attr(it, "Design temperature"),
                Attr(it, "Design Temperature")
            );

            string dp2 = FirstNonEmpty(
                Attr(it, "Design pressure 2"), Attr(it, "Design pressure 2 (bar g)"),
                Attr(it, "Design Pressure 2"), Attr(it, "DP2")
            );
            string dt2 = FirstNonEmpty(
                Attr(it, "Design temperature 2"), Attr(it, "Design temperature 2 (°C)"),
                Attr(it, "Design Temperature 2"), Attr(it, "DT2")
            );
            string dp3 = FirstNonEmpty(
                Attr(it, "Design pressure 3"), Attr(it, "Design pressure 3 (bar g)"),
                Attr(it, "Design Pressure 3"), Attr(it, "DP3")
            );
            string dt3 = FirstNonEmpty(
                Attr(it, "Design temperature 3"), Attr(it, "Design temperature 3 (°C)"),
                Attr(it, "Design Temperature 3"), Attr(it, "DT3")
            );

            // --- Design code / rating / class ---------------------------------------
            string designCode = FirstNonEmpty(
                Attr(it, "Design code"),
                Attr(it, "Design Code"),
                Attr(it, "Standard"),
                Attr(it, "Code")
            );
            string rating = Attr(it, "Rating");
            string cls = Attr(it, "Class");

            // --- MATERIAL block (C1→F, C2→H, C3→I; rows 19..22) --------------------
            string d1c1 = Attr(it, "Description 1 (C1)");
            string d2c1 = Attr(it, "Description 2 (C1)");
            string d3c1 = Attr(it, "Description 3 (C1)");
            string d4c1 = Attr(it, "Description 4 (C1)");

            string d1c2 = Attr(it, "Description 1 (C2)");
            string d2c2 = Attr(it, "Description 2 (C2)");
            string d3c2 = Attr(it, "Description 3 (C2)");
            string d4c2 = Attr(it, "Description 4 (C2)");

            string d1c3 = Attr(it, "Description 1 (C3)");
            string d2c3 = Attr(it, "Description 2 (C3)");
            string d3c3 = Attr(it, "Description 3 (C3)");
            string d4c3 = Attr(it, "Description 4 (C3)");

            string materialGroup = FirstNonEmpty(Attr(it, "Material Group"), Attr(it, "Material group"));
            string materialNumber = FirstNonEmpty(Attr(it, "Material number"), Attr(it, "Material Number"), Attr(it, "Material"));

            // C1 -> F19..F22
            Set(ws, "F19", d1c1);
            Set(ws, "F20", d2c1);
            Set(ws, "F21", d3c1);
            Set(ws, "F22", d4c1);

            // C2 -> H19..H22 (H19 fallback → Material Group)
            Set(ws, "H19", string.IsNullOrWhiteSpace(d1c2) ? materialGroup : d1c2);
            Set(ws, "H20", d2c2);
            Set(ws, "H21", d3c2);
            Set(ws, "H22", d4c2);

            // C3 -> I19..I22 (I20 fallback → Material number == material code)
            Set(ws, "I19", d1c3);
            Set(ws, "I20", string.IsNullOrWhiteSpace(d2c3) ? materialNumber : d2c3);
            Set(ws, "I21", d3c3);
            Set(ws, "I22", d4c3);

            // --- Piping connection & Additional specifications (RESTORED) ------------
            string endConn = FirstNonEmpty(
                Attr(it, "Piping Connection 1"),
                AttrLike("Piping Connection 1"),
                Attr(it, "Piping connection 1"),
                Attr(it, "Piping connection"),
                Attr(it, "End connection"),
                Attr(it, "Connection"),
                Attr(it, "Ends"),
                Attr(it, "Flange Facing"),
                Attr(it, "Facing")
            );

            string add1 = FirstNonEmpty(
                Attr(it, "Additional specification 1"),
                Attr(it, "Additional Specification 1"),
                Attr(it, "Remark 1"),
                AttrLike("Remark 1")
            );
            string add2 = FirstNonEmpty(
                Attr(it, "Additional specification 2"),
                Attr(it, "Additional Specification 2"),
                Attr(it, "Remark 2"),
                AttrLike("Remark 2")
            );

            string operation = FirstNonEmpty(Attr(it, "Operation"), Attr(it, "Operator"));

            string pipingClass = FirstNonEmpty(
                Attr(it, "Piping class"),
                Attr(it, "Pipe class"),
                Attr(it, "JLE Possible Pipe Class"),
                Attr(it, "JLE Pipe Class"),
                Attr(it, "JLE Pipe Class 1")
            );

            // --- Write to sheet -----------------------------------------------------
            Set(ws, "F7", code);
            Set(ws, "F9", vtype);

            Set(ws, "F11", medium);
            Set(ws, "F12", corrosive);

            Set(ws, "F14", dp1);
            Set(ws, "F15", dt1);
            Set(ws, "G14", dp2);
            Set(ws, "G15", dt2);
            Set(ws, "H14", dp3);
            Set(ws, "H15", dt3);

            Set(ws, "F17", designCode);

            // Optional label for rating/class area (depends on template layout)
            if (!string.IsNullOrWhiteSpace(rating)) Set(ws, "D17", "RATING");
            else if (!string.IsNullOrWhiteSpace(cls)) Set(ws, "D17", "CLASS");

            // RESTORED: PIPING CONNECTION + OPERATION + ADDITIONAL SPEC
            if (!string.IsNullOrWhiteSpace(endConn))
            {
                Set(ws, "D24", "PIPING CONNECTION");
                Set(ws, "F24", endConn);
            }
            if (!string.IsNullOrWhiteSpace(operation))
            {
                Set(ws, "D25", "OPERATION");
                Set(ws, "F25", operation);
            }
            Set(ws, "F26", add1);
            Set(ws, "F27", add2);

            // Piping class (footer / info row if present in template)
            Set(ws, "F29", pipingClass);

            // Header / version
            Set(ws, "D4", "VALVE DATA SHEET");
            Set(ws, "D57", "1.0");

            // --- Size table ---------------------------------------------------------
            string sminRaw = FirstNonEmpty(
                Attr(it, "Size (Min)"),
                AttrLike("Size (Min)"),
                Attr(it, "Size min"),
                Attr(it, "Size Min"),
                Attr(it, "MIN SIZE"),
                Attr(it, "DN min"),
                Attr(it, "DN Min")
            );

            string smaxRaw = FirstNonEmpty(
                Attr(it, "Size (Max)"),
                AttrLike("Size (Max)"),
                Attr(it, "Size max"),
                Attr(it, "Size Max"),
                Attr(it, "MAX SIZE"),
                Attr(it, "DN max"),
                Attr(it, "DN Max")
            );

            // start at row 33 so first size is under the header (D33)
            WriteValveSizesAcceptingDn(ws, sminRaw, smaxRaw, 33);
        }

        // ========== Size helpers ==========
        private static string JoinParts(params string[] parts)
        {
            return string.Join(" ",
                parts.Where(p => !string.IsNullOrWhiteSpace(p)).Select(p => p.Trim()));
        }

        private static string DnToInchKey(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            s = s.Trim().ToUpperInvariant();
            s = s.Replace("\"", "").Replace(" IN", "").Replace("IN ", "").Trim();

            if (!s.StartsWith("DN")) return s; // already inch
            string dn = s;

            var rev = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "DN15","1/2" },{ "DN20","3/4" },{ "DN25","1" },{ "DN32","1 1/4" },{ "DN40","1 1/2" },
                { "DN50","2" },{ "DN65","2 1/2" },{ "DN80","3" },{ "DN100","4" },{ "DN125","5" },{ "DN150","6" },
                { "DN200","8" },{ "DN250","10" },{ "DN300","12" },{ "DN350","14" },{ "DN400","16" },
                { "DN450","18" },{ "DN500","20" },{ "DN550","22" },{ "DN600","24" }
            };
            return rev.TryGetValue(dn, out var inch) ? inch : s;
        }

        // Valve sizes: accept DN or inch, write NPS sizes starting at 'startRow'
        private static void WriteValveSizesAcceptingDn(IXLWorksheet ws, string sizeMin, string sizeMax, int startRow)
        {
            if (string.IsNullOrWhiteSpace(sizeMin) || string.IsNullOrWhiteSpace(sizeMax))
                return;

            string smin = DnToInchKey(sizeMin);
            string smax = DnToInchKey(sizeMax);

            smin = NormalizeInch(smin);
            smax = NormalizeInch(smax);
            if (string.IsNullOrWhiteSpace(smin) || string.IsNullOrWhiteSpace(smax)) return;

            int i1 = NpsSizes.FindIndex(x => string.Equals(NormalizeInch(x), smin, StringComparison.OrdinalIgnoreCase));
            int i2 = NpsSizes.FindIndex(x => string.Equals(NormalizeInch(x), smax, StringComparison.OrdinalIgnoreCase));
            if (i1 < 0 || i2 < 0) return;
            if (i2 < i1) { int t = i1; i1 = i2; i2 = t; }

            int row = startRow;
            for (int i = i1; i <= i2; i++)
            {
                string inch = NpsSizes[i];

                ws.Cell(row, 4).Value = inch;   // D: Size NPS/DN column
                ws.Cell(row, 6).Clear();        // F: Unit price (blank)
                ws.Cell(row, 7).Clear();        // G: Qty (blank)

                var total = ws.Cell(row, 8);    // H: Total = G * F
                if (total.IsEmpty() && !total.HasFormula)
                    total.FormulaA1 = $"IFERROR({ws.Cell(row, 7).Address.ToStringRelative()}*{ws.Cell(row, 6).Address.ToStringRelative()},0)";

                row++;
            }
        }

        // Gasket sizes: always write DN text starting at row below the header (usually 26)
        // Gasket sizes: write DN text starting at the detected start row.
        // Column layout (per your template):
        // D = Size NPS/DN, E = Quantity required, F = Quantity ordered, G = Unit Price, H = Total Price
        private static void WriteGasketSizesDn(IXLWorksheet ws, ItemRow it, int startRow)
        {
            string smin = NormalizeInch(Attr(it, "Size (Min)"));
            string smax = NormalizeInch(Attr(it, "Size (Max)"));
            if (string.IsNullOrWhiteSpace(smin) || string.IsNullOrWhiteSpace(smax)) return;

            int i1 = NpsSizes.FindIndex(x => string.Equals(NormalizeInch(x), smin, StringComparison.OrdinalIgnoreCase));
            int i2 = NpsSizes.FindIndex(x => string.Equals(NormalizeInch(x), smax, StringComparison.OrdinalIgnoreCase));
            if (i1 < 0 || i2 < 0) return;
            if (i2 < i1) { int t = i1; i1 = i2; i2 = t; }

            int row = startRow;
            for (int i = i1; i <= i2; i++)
            {
                string inch = NpsSizes[i];
                string dn = InchToDn(inch);

                // D: Size text (DNxx)
                ws.Cell(row, 4).Value = dn;

                // E (Quantity required) – leave as-is (don’t touch formatting/content).

                // F: Quantity ordered → set default 0.00 but KEEP the template's formatting
                var qtyCell = ws.Cell(row, 6);
                qtyCell.Value = 0.0; // template number format should render as 0.00
                                     // If your template sometimes has General, force format:
                                     // qtyCell.Style.NumberFormat.Format = "0.00";

                // G: Unit Price → empty, but preserve styling (shade/borders/number format)
                ws.Cell(row, 7).Clear(XLClearOptions.Contents);

                // H: Total = G * F (Unit Price * Quantity ordered)
                var total = ws.Cell(row, 8);
                if (total.IsEmpty() && !total.HasFormula)
                    total.FormulaA1 = $"IFERROR({ws.Cell(row, 7).Address.ToStringRelative()}*{ws.Cell(row, 6).Address.ToStringRelative()},0)";

                row++;
            }
        }





        // ---------- Populate: Bolts & Nuts (PP Python-aligned) ----------
        private static void PopulateBolts(IXLWorksheet ws, ItemRow it)
        {
            string code = FirstNonEmpty(Attr(it, "Device designation"), Attr(it, "Code"), it.Name);
            string type = FirstNonEmpty(Attr(it, "Additional Comment"), Attr(it, "Type"), Attr(it, "Comment"));

            string matBolts = FirstNonEmpty(
                Attr(it, "Material - Bolts"),
                Attr(it, "Material number"),
                Attr(it, "Bolt Material"),
                Attr(it, "Bolts Material"));

            string matNuts = FirstNonEmpty(
                Attr(it, "Material - Nuts"),
                Attr(it, "Nut Material"),
                Attr(it, "Nuts Material"));

            string coating = Attr(it, "Coating");

            string accStd = FirstNonEmpty(
                Attr(it, "Acc to standard"),
                Attr(it, "Acc to Standard"),
                Attr(it, "ACC. TO STANDARD"),
                Attr(it, "Material Standard"));

            string remarks = FirstNonEmpty(Attr(it, "Remark 1"), Attr(it, "Remark 2"), Attr(it, "Remarks"));

            Set(ws, "G7", code);   // Code
            Set(ws, "G9", type);   // Type

            Set(ws, "G11", matBolts);
            Set(ws, "G12", matNuts);
            Set(ws, "G13", coating);
            Set(ws, "G15", accStd);
            Set(ws, "G20", remarks);

            // Optional named ranges (no-op if missing)
            TrySetByNamedRange(ws, "Code", code);
            TrySetByNamedRange(ws, "Type", type);
            TrySetByNamedRange(ws, "MaterialBolts", matBolts);
            TrySetByNamedRange(ws, "MaterialNuts", matNuts);
            TrySetByNamedRange(ws, "Coating", coating);
            TrySetByNamedRange(ws, "AccToStandard", accStd);
            TrySetByNamedRange(ws, "Remarks", remarks);
        }

        // ---------- Helpers ----------
        private static string Attr(ItemRow r, params string[] keys)
        {
            if (r == null || r.Attributes == null || keys == null || keys.Length == 0) return "";
            foreach (string k in keys)
            {
                if (k == null) continue;
                if (r.Attributes.TryGetValue(k, out var v) && v != null) return v;
            }
            foreach (string k in keys)
            {
                if (k == null) continue;
                var hit = r.Attributes.FirstOrDefault(kv => kv.Key.Equals(k, StringComparison.OrdinalIgnoreCase));
                if (!string.IsNullOrEmpty(hit.Key) && hit.Value != null) return hit.Value;
            }
            return "";
        }

        private static string FirstNonEmpty(params string[] values)
        {
            if (values == null) return "";
            foreach (string v in values) if (!string.IsNullOrWhiteSpace(v)) return v;
            return "";
        }

        private static void Set(IXLWorksheet ws, string addr, string value)
        {
            if (!string.IsNullOrWhiteSpace(value)) ws.Cell(addr).Value = value;
        }

        private static bool TrySetByNamedRange(IXLWorksheet ws, string name, string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return true;
            try
            {
                var nr = ws.Workbook.DefinedNames
                    .FirstOrDefault(dn => dn.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
                if (nr != null && nr.Ranges.Any())
                {
                    nr.Ranges.First().Cells().First().Value = value;
                    return true;
                }
            }
            catch { }
            return false;
        }

        private static IXLCell FindCellByText(IXLWorksheet ws, string label)
        {
            string norm = (label ?? "").Trim().ToUpperInvariant();

            foreach (IXLCell cell in ws.CellsUsed(c => !c.IsMerged()))
            {
                string txt = (cell.GetString() ?? "").Trim().ToUpperInvariant();
                if (txt == norm) return cell;
            }
            foreach (IXLCell cell in ws.CellsUsed(c => !c.IsMerged()))
            {
                string txt = (cell.GetString() ?? "").Trim().ToUpperInvariant().Replace(".", "");
                if (txt == norm.Replace(".", "")) return cell;
            }
            return null;
        }

        private static void WriteToRightOfLabel(IXLWorksheet ws, string label, string value, int offsetColumns, string fallbackAddr)
        {
            if (string.IsNullOrWhiteSpace(value)) return;

            var cell = FindCellByText(ws, label);
            if (cell != null)
            {
                try
                {
                    ws.Cell(cell.Address.RowNumber, cell.Address.ColumnNumber + offsetColumns).Value = value;
                    return;
                }
                catch { /* fallback below */ }
            }
            Set(ws, fallbackAddr, value);
        }

        private static string NormalizeInch(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            return s.Replace("\"", string.Empty).Replace("IN", string.Empty).Trim().ToUpperInvariant();
        }

        private static string InchToDn(string inch)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "1/2", "DN15" }, { "3/4", "DN20" }, { "1", "DN25" }, { "1 1/4", "DN32" },
                { "1 1/2", "DN40" }, { "2", "DN50" }, { "2 1/2", "DN65" }, { "3", "DN80" },
                { "4", "DN100" }, { "5", "DN125" }, { "6", "DN150" }, { "8", "DN200" },
                { "10", "DN250" }, { "12", "DN300" }, { "14", "DN350" }, { "16", "DN400" },
                { "18", "DN450" }, { "20", "DN500" }, { "22", "DN550" }, { "24", "DN600" }
            };
            string key = NormalizeInch(inch);
            return map.TryGetValue(key, out var dn) ? dn : inch;
        }

        private static void WriteSizesFromRange(IXLWorksheet ws, ItemRow it, int startRow, bool toDn)
        {
            string smin = NormalizeInch(Attr(it, "Size min"));
            string smax = NormalizeInch(Attr(it, "Size max"));
            if (string.IsNullOrWhiteSpace(smin) || string.IsNullOrWhiteSpace(smax)) return;

            int i1 = NpsSizes.FindIndex(x => string.Equals(NormalizeInch(x), smin, StringComparison.OrdinalIgnoreCase));
            int i2 = NpsSizes.FindIndex(x => string.Equals(NormalizeInch(x), smax, StringComparison.OrdinalIgnoreCase));
            if (i1 < 0 || i2 < 0) return;
            if (i2 < i1) { int t = i1; i1 = i2; i2 = t; }

            int row = startRow;
            for (int i = i1; i <= i2; i++)
            {
                string inch = NpsSizes[i];
                string display = toDn ? InchToDn(inch) : inch;

                ws.Cell(row, 4).Value = display;
                ws.Cell(row, 6).Clear();
                ws.Cell(row, 7).Clear();

                var total = ws.Cell(row, 8);
                if (total.IsEmpty() && !total.HasFormula)
                    total.FormulaA1 = $"IFERROR({ws.Cell(row, 7).Address.ToStringRelative()}*{ws.Cell(row, 6).Address.ToStringRelative()},0)";
                row++;
            }
        }

        private static Category ClassifyFromText(string typeHint, ItemRow row)
        {
            if (IsLikelyValve(row)) return Category.Valve;

            string blob = (FirstNonEmpty(typeHint, "") + " " +
                           FirstNonEmpty(Attr(row, "Type"), "") + " " +
                           FirstNonEmpty(Attr(row, "Additional Comment"), "") + " " +
                           FirstNonEmpty(Attr(row, "Comment"), "") + " " +
                           (row?.Name ?? "")).ToUpperInvariant();

            if (blob.Contains("GASKET")) return Category.Gasket;
            if (blob.Contains("BOLT") || blob.Contains("NUT")) return Category.Bolt;
            return Category.PipingPart;
        }

        private static string DetectPipingSubCategory(ItemRow row)
        {
            string blob =
                (FirstNonEmpty(
                    Attr(row, "Type"),
                    Attr(row, "Additional Comment"),
                    Attr(row, "Specification"),
                    Attr(row, "Body/Fitting type"),
                    Attr(row, "Comment"),
                    row.Name
                ) ?? string.Empty).ToUpperInvariant();

            // ----- FLANGES -----
            if (blob.Contains("FLANGE") ||
                blob.Contains("WELD NECK") ||
                blob.Contains("WN") ||
                blob.Contains("SLIP ON") ||
                blob.Contains("SO ") ||
                blob.Contains("BLIND"))
                return "FLANGES";

            // ----- BRANCH FITTINGS -----
            if (blob.Contains("OLET") ||
                blob.Contains("BRANCH"))
                return "BRANCH FITTINGS";

            // ----- FITTINGS -----
            if (blob.Contains("ELBOW") ||
                blob.Contains("TEE") ||
                blob.Contains("REDUCER") ||
                blob.Contains("COUPLING") ||
                blob.Contains("CAP") ||
                blob.Contains("BEND"))
                return "FITTINGS";

            // ----- PIPE -----
            if (blob.Contains("PIPE") ||
                blob.Contains("SEAMLESS") ||
                blob.Contains("WELDED") ||
                blob.Contains("SMLS"))
                return "PIPE";

            return "PIPING PARTS";
        }


        private static bool IsLikelyValve(ItemRow row)
        {
            try
            {
                if (row?.Attributes != null)
                {
                    foreach (var kv in row.Attributes)
                    {
                        if (kv.Key != null && kv.Key.IndexOf("VALVE", StringComparison.OrdinalIgnoreCase) >= 0) return true;
                        if (!string.IsNullOrWhiteSpace(kv.Value) && kv.Value.IndexOf("VALVE", StringComparison.OrdinalIgnoreCase) >= 0) return true;
                    }
                }
            }
            catch { }

            string blob = (FirstNonEmpty(
                              Attr(row, "Valve type"),
                              Attr(row, "Type"),
                              Attr(row, "Additional Comment"),
                              Attr(row, "Specification"),
                              Attr(row, "Comment"),
                              row?.Name
                          ) ?? string.Empty).ToUpperInvariant();

            string[] valveWords =
            {
                "VALVE","GATE","GLOBE","CHECK","BALL","BUTTERFLY","PLUG","DIAPHRAGM",
                "CONTROL","SOLENOID","SAFETY","RELIEF","PRESSURE REDUCING","PSV","PRV","SRV","NRV"
            };
            if (valveWords.Any(w => blob.Contains(w))) return true;

            string[] valveKeys =
            {
                "Medium","Media","Service","Design pressure","Design temperature","Design code",
                "Piping connection","End connection","Additional specification 1","Additional specification 2",
                "Body material","Seat material","Trim material"
            };
            int hits = valveKeys.Count(k => !string.IsNullOrWhiteSpace(Attr(row, k)));
            return hits >= 1;
        }

        private static string SafeSheetName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) name = "ITEM";
            foreach (char ch in Path.GetInvalidFileNameChars()) name = name.Replace(ch, '_');
            return name.Length > 31 ? name.Substring(0, 31) : name;
        }

        private static IXLWorksheet FindTemplateSheet(XLWorkbook wb, string preferredName, int fallbackSheetIndex)
        {
            var ws = wb.Worksheets.FirstOrDefault(s => s.Name.Equals(preferredName, StringComparison.OrdinalIgnoreCase));
            return ws ?? wb.Worksheet(fallbackSheetIndex);
        }

        private static void TryDelete(XLWorkbook wb, string sheetName)
        {
            try
            {
                var ws = wb.Worksheets.FirstOrDefault(s => s.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
                if (ws != null && wb.Worksheets.Count > 1) ws.Delete();
            }
            catch { }
        }
    }

    public class DatasheetResult
    {
        public int Created { get; set; }
        public string Message { get; set; }
        public DatasheetResult() { Created = 0; Message = ""; }
    }
}
