
// PipeClassServiceEb.cs — C# 8 SAFE (no target-typed new, no empty char literals)
#nullable enable
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using Aucotec.EngineeringBase.Client.Runtime;
// Alias EB Application to avoid clash with WPF/System.Windows.Application
using EbApp = Aucotec.EngineeringBase.Client.Runtime.Application;

namespace JJ_Lurgi_Piping_EB
{
    /// <summary>
    /// Pipe Class summary generator:
    /// 1) Scans EB Catalogs (narrowed to JLE/Materials/{Bolts & Nuts, Pipe & Fittings, Valves}) to find items for a selected pipe class.
    /// 2) Writes a per-class workbook using JJ-Lurgi layout (PIPING PARTS, then BOLTING, GASKET, VALVE).
    /// Robust fuzzy attribute lookups to prevent blank/missing fields.
    /// </summary>
    public static class PipeClassServiceEb
    {
        // ======= Paths (edit if needed) =======
        private static readonly string ProjectRoot = @"E:\Aucotec Developer\Tahsin\JJ_Lurgi_Piping_EB";
        private static readonly string TemplatesRoot = Path.Combine(ProjectRoot, "templates");
        private static readonly string OutputRoot = Path.Combine(ProjectRoot, "output");
        private const string PipeClassTemplateFile = "Pipe class template - WIP.xlsx";

        // ======= Column mapping (Excel) =======
        private const int COL_TYPE = 1;       // A
        private const int COL_DESC = 5;       // E
        private const int COL_SIZE_MIN = 10;  // J
        private const int COL_SIZE_MAX = 11;  // K
        private const int COL_SCHCLASS = 12;  // L
        private const int COL_CODE = 14;      // N

        // Pipe-class code (e.g., 150JX00)
        private static readonly Regex PipeClassCodeRx =
            new Regex(@"^[0-9]{1,4}J[A-Z]{1,2}[0-9]{2}$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // Attributes that may contain pipe-class membership
        private static readonly string[] ClassAttrCandidates = new string[]
        {
            "JLE Pipe Class",
            "JLE Pipe Class 1",
            "JLE Possible Pipe Class",
            "Piping class",
            "Pipe class",
            "Piping classes that use this item (for piping class)"
        };

        // ========== PUBLIC API ==========

        public static List<string> FetchAllPipeClassesFromEb(EbApp app)
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (app == null) return new List<string>();

            ObjectItem catalogs = null;
            try { catalogs = app.Folders?.Catalogs; } catch { }
            if (catalogs == null) return new List<string>();

            // Limit traversal to JLE → Materials → {Bolts & Nuts, Pipe & Fittings, Valves}
            var jle = GetChild(catalogs, "JLE");
            var materials = jle != null ? GetChild(jle, "Materials") : null;
            var roots = new List<ObjectItem>();
            if (materials != null)
            {
                var bolts = GetChild(materials, "Bolts & Nuts");
                var pf = GetChild(materials, "Pipe & Fittings");
                var valves = GetChild(materials, "Valves");
                if (bolts != null) roots.Add(bolts);
                if (pf != null) roots.Add(pf);
                if (valves != null) roots.Add(valves);
            }
            if (roots.Count == 0) roots.Add(catalogs);

            foreach (var root in roots)
            {
                foreach (ObjectItem obj in WalkDeep(root))
                {
                    TryAddFromNamedAttributes(obj, set, ClassAttrCandidates);
                    string nm = SafeName(obj);
                    nm = Normalize(nm);
                    if (IsPipeClassCode(nm)) set.Add(nm);
                }
            }

            var all = new List<string>(set.Where(IsPipeClassCode).Select(Normalize)
                                          .Distinct(StringComparer.OrdinalIgnoreCase));
            all.Sort(StringLogicalComparer.Instance);
            return all;
        }

        public static void SplitAsAsmeDin(IEnumerable<string> all, out List<string> asme, out List<string> din)
        {
            var a = new List<string>();
            var d = new List<string>();
            foreach (string code in all ?? Enumerable.Empty<string>())
            {
                string ratingTok = ExtractRatingToken(code);
                if (IsAsmeRating(ratingTok)) a.Add(code);
                else d.Add(code);
            }
            a.Sort(StringLogicalComparer.Instance);
            d.Sort(StringLogicalComparer.Instance);
            asme = a; din = d;
        }

        public static ClassGroups GroupByRating(IEnumerable<string> classes)
        {
            var result = new ClassGroups();
            if (classes == null) return result;

            foreach (var pc in classes)
            {
                if (string.IsNullOrWhiteSpace(pc)) continue;
                var tok = ExtractRatingToken(pc);
                if (string.IsNullOrEmpty(tok)) result.Other.Add(pc);
                else if (IsAsmeRating(tok)) result.Asme.Add(pc);
                else result.Din.Add(pc);
            }

            result.Asme.Sort(StringLogicalComparer.Instance);
            result.Din.Sort(StringLogicalComparer.Instance);
            result.Other.Sort(StringLogicalComparer.Instance);

            result.Ordered.AddRange(result.Asme);
            result.Ordered.AddRange(result.Din);
            result.Ordered.AddRange(result.Other);
            return result;
        }

        public static string GenerateForClasses(EbApp app, IEnumerable<string> classes)
        {
            if (app == null) return "No EB instance.";
            Directory.CreateDirectory(OutputRoot);

            string templatePath = Path.Combine(TemplatesRoot, PipeClassTemplateFile);
            if (!File.Exists(templatePath))
                return "Template not found: " + templatePath;

            int ok = 0, fail = 0;
            foreach (string pclass in classes ?? Enumerable.Empty<string>())
            {
                if (string.IsNullOrWhiteSpace(pclass)) continue;
                try
                {
                    string outFile = Path.Combine(OutputRoot, SanitizeFileName(pclass) + ".xlsx");
                    File.Copy(templatePath, outFile, true);

                    using (var wb = new XLWorkbook(outFile))
                    {
                        var sheet = wb.Worksheets.Count > 0 ? wb.Worksheets.First() : null;
                        if (sheet == null) throw new InvalidOperationException("Template workbook has no sheets.");
                        sheet.Name = "PIPING PARTS";
                        sheet.Cell("D3").Value = "PIPING CLASS BASIC DATAs - " + pclass;

                        FillSheetForClass_All(app, pclass, sheet);

                        wb.Save();
                    }
                    ok++;
                }
                catch (Exception ex)
                {
                    fail++;
                    System.Diagnostics.Debug.WriteLine("[JJLEM] GenerateForClasses error for " + pclass + ": " + ex.Message);
                }
            }

            return "Pipe class summaries generated: OK=" + ok + ", Failed=" + fail + ". Output: " + OutputRoot;
        }

        // ========== CORE FILLER ==========

        private static void FillSheetForClass_All(EbApp app, string pipeClass, IXLWorksheet sheet)
        {
            if (app == null || sheet == null || string.IsNullOrWhiteSpace(pipeClass)) return;

            List<ObjectItem> allForClass = CollectItemsForClass(app, pipeClass);

            var listPP = new List<ObjectItem>();
            var listB = new List<ObjectItem>();
            var listG = new List<ObjectItem>();
            var listV = new List<ObjectItem>();

            foreach (ObjectItem it in allForClass)
            {
                if (LooksLikeValve(it)) listV.Add(it);
                else if (LooksLikeGasket(it)) listG.Add(it);
                else if (LooksLikeBoltsOrNuts(it)) listB.Add(it);
                else listPP.Add(it);
            }

            int row = 8;

            // ---- PIPING PARTS (grouped by Sorting for piping class: 1/2/3/4) ----
            // ---- PIPING PARTS (PIPE / FITTINGS / BRANCH FITTINGS / FLANGES ) ----
            var ppViews = new List<PPView>();
            foreach (ObjectItem it in listPP)
            {
                string code = Attr(it, true, "Code", "Item code", "Part code");
                string sort = Attr(it, true, "Sorting for piping class", "Sorting For Piping Class");
                string group = FirstChar(sort);
                ppViews.Add(new PPView { Obj = it, Code = code, GroupKey = group });
            }

            // Order properly (1=PIPE, 2=FITTINGS, 3=BRANCH FITTINGS, 4=FLANGES)
            ppViews = ppViews
                .OrderBy(v => GroupSortKey(v.GroupKey))
                .ThenBy(v => v.Code ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .ToList();

            string currentGroup = string.Empty;

            foreach (var view in ppViews)
            {
                if (!string.Equals(view.GroupKey, currentGroup, StringComparison.OrdinalIgnoreCase))
                {
                    currentGroup = view.GroupKey;

                    string header = GroupHeader(currentGroup);
                    if (!string.IsNullOrEmpty(header))
                    {
                        row++; // move down so header isn't swallowed inside template merged rows
                        row = WriteHeader(sheet, row, header);
                        row++; // add spacing
                    }
                }

                row = WritePPRows(sheet, row, view.Obj);
            }



            // ---- BOLTING ----
            if (listB.Count > 0)
            {
                row = WriteHeader(sheet, row, "BOLTING");
                foreach (var it in listB) row = WriteBoltRows(sheet, row, it);
            }

            // ---- GASKET ----
            if (listG.Count > 0)
            {
                row = WriteHeader(sheet, row, "GASKET");
                foreach (var it in listG) row = WriteGasketRows(sheet, row, it);
            }

            // ---- VALVE ----
            if (listV.Count > 0)
            {
                row = WriteHeader(sheet, row, "VALVE");
                foreach (var it in listV) row = WriteValveRows(sheet, row, it);
            }

            try
            {
                sheet.PageSetup.PrintAreas.Clear();
                sheet.PageSetup.PrintAreas.Add("A1:N" + row.ToString());
            }
            catch { }
        }

        // ========== WRITERS ==========

        private static int WriteHeader(IXLWorksheet s, int row, string text)
        {
            // FIX: Unmerge any merged cells in this entire row (ClosedXML-safe)
            try
            {
                var rng = s.Range(row, 1, row, 20);
                foreach (var cell in rng.Cells())
                {
                    if (cell.IsMerged())
                        cell.MergedRange().Unmerge();
                }
            }
            catch { }

            // Write header like Bolting/Gasket/Valve sections
            var c = s.Cell(row, COL_TYPE);
            c.Value = (text ?? string.Empty).ToUpperInvariant();
            c.Style.Font.Bold = true;

            return row + 1;
        }



        // ---------- PIPING PARTS ----------
        // Python write_pp_row naming:
        //   A(row)   = Type
        //   A(row+1) = Seamless / Welded
        private static int WritePPRows(IXLWorksheet s, int row, ObjectItem obj)
        {
            // Line 1 -> Type (but avoid pure numeric codes like 1765; then use comments instead)
            string type = FirstNonEmpty(
                Attr(obj, true, "Type"),
                Attr(obj, true, "Body/Fitting type", "Body/Fitting Type", "Body / Fitting type")
            );
            if (LooksLikeNumericCode(type))
            {
                type = FirstNonEmpty(
                    Attr(obj, true, "Additional Comment", "Additional comment"),
                    Attr(obj, true, "Comment", "Item Tag", "Tag"),
                    type
                );
            }

            // Line 2 -> Seamless / Welded
            string seamWeld = FirstNonEmpty(
                Attr(obj, true, "Seamless / Welded", "Seamless/Welded"),
                Attr(obj, true, "Seamless - Welded", "Seamless / welded"),
                Attr(obj, true, "Seamless"),
                Attr(obj, true, "Welded")
            );

            s.Cell(row + 0, COL_TYPE).Value = (type ?? string.Empty).Trim();
            s.Cell(row + 1, COL_TYPE).Value = (seamWeld ?? string.Empty).Trim();

            // Column E: Material / Acc to Standard / Piping connection 1
            string matNum = Attr(obj, true, "Material", "Material number", "Material Number");
            string matStd = Attr(obj, true, "Acc to Standard", "Acc to standard",
                                          "Material Standard", "Material standard");
            string pc1 = Attr(obj, true, "Piping connection 1", "Piping Connection 1", "End Connection 1");

            s.Cell(row + 0, COL_DESC).Value = matNum;
            s.Cell(row + 1, COL_DESC).Value = matStd;
            s.Cell(row + 2, COL_DESC).Value = pc1;

            // Sizes
            string szMin = Attr(obj, true, "Size min", "Size (Min)", "Size Min");
            string szMax = Attr(obj, true, "Size max", "Size (Max)", "Size Max");

            s.Cell(row + 0, COL_SIZE_MIN).Value = szMin;
            s.Cell(row + 0, COL_SIZE_MAX).Value = szMax;

            // Schedule / Class / Rating
            string schNo = Attr(obj, true, "Schedule", "Pipe schedule no", "Pipe Schedule No", "Pipe schedule number");
            string cls = Attr(obj, true, "Class");
            string rating = Attr(obj, true, "Rating");

            if (!string.IsNullOrWhiteSpace(schNo))
                s.Cell(row + 0, COL_SCHCLASS).Value = "SCH" + schNo.Trim();
            else if (!string.IsNullOrWhiteSpace(cls))
                s.Cell(row + 0, COL_SCHCLASS).Value = cls;
            else if (!string.IsNullOrWhiteSpace(rating))
                s.Cell(row + 0, COL_SCHCLASS).Value = rating;
            else
                s.Cell(row + 0, COL_SCHCLASS).Clear();

            // Item code in N
            string itemTag = FirstNonEmpty(
                Attr(obj, true, "Code", "Item code", "Part code"),
                Attr(obj, true, "Comment", "Additional Comment", "Item Tag", "Tag")
            );
            s.Cell(row + 0, COL_CODE).Value = itemTag;

            ClearRow(s, row + 3);
            return row + 4;
        }

        // ---------- BOLTING ----------
        // Python write_bolt_row naming:
        //   A(row) = Type
        private static int WriteBoltRows(IXLWorksheet s, int row, ObjectItem obj)
        {
            // A = Type (prefer human-friendly comments if Type is numeric)
            string type = Attr(obj, true, "Type");
            if (LooksLikeNumericCode(type))
            {
                type = FirstNonEmpty(
                    Attr(obj, true, "Additional Comment", "Additional comment"),
                    Attr(obj, true, "Comment", "Item Tag", "Tag"),
                    type
                );
            }
            s.Cell(row + 0, COL_TYPE).Value = (type ?? string.Empty).Trim();

            // E: Acc to standard, then materials/coating
            string accStd = Attr(obj, true, "Acc to standard", "Acc to Standard", "Standard");

            string matBolts = FirstNonEmpty(
                Attr(obj, true, "Material - Bolts", "Material bolts", "Bolts material"),
                AttrLike(obj, "bolt material")
            );
            string matNuts = FirstNonEmpty(
                Attr(obj, true, "Material - Nuts", "Material nuts", "Nuts material"),
                AttrLike(obj, "nut material")
            );
            string coating = Attr(obj, true, "Coating", "Finish", "Surface coating");

            s.Cell(row + 0, COL_DESC).Value = accStd;

            if (!string.IsNullOrWhiteSpace(coating) &&
                !coating.Equals("NONE", StringComparison.OrdinalIgnoreCase))
            {
                s.Cell(row + 1, COL_DESC).Value = JoinIf(matBolts, " - ", coating);
                s.Cell(row + 2, COL_DESC).Value = JoinIf(matNuts, " - ", coating);
            }
            else
            {
                s.Cell(row + 1, COL_DESC).Value = matBolts;
                s.Cell(row + 2, COL_DESC).Value = matNuts;
            }

            // J: MATCHING FLANGE
            s.Cell(row + 0, COL_SIZE_MIN).Value = "MATCHING FLANGE";
            LeftAlign(s.Cell(row + 0, COL_SIZE_MIN));

            // N: Code
            string code = FirstNonEmpty(
                Attr(obj, true, "Code", "Item code", "Part code"),
                Attr(obj, true, "Comment", "Item Tag", "Tag")
            );
            s.Cell(row + 0, COL_CODE).Value = code;

            ClearRow(s, row + 3);
            return row + 4;
        }

        // ---------- GASKET ----------
        // Python write_gasket_row naming:
        //   A(row)   = Type
        //   A(row+1) = Inside / Outside Ring
        private static int WriteGasketRows(IXLWorksheet s, int row, ObjectItem obj)
        {
            // A line1: Type (prefer comments if numeric)
            string type = Attr(obj, true, "Type");
            if (LooksLikeNumericCode(type))
            {
                type = FirstNonEmpty(
                    Attr(obj, true, "Additional Comment", "Additional comment"),
                    Attr(obj, true, "Comment", "Item Tag", "Tag"),
                    type
                );
            }

            // A line2: Inside / Outside Ring
            string insideOutside = FirstNonEmpty(
                Attr(obj, true, "Inside / Outside Ring", "Inside/Outside Ring"),
                AttrLike(obj, "Inside / Outside Ring"),
                AttrLike(obj, "Inside/Outside"),
                AttrLike(obj, "Inside Outside")
            );

            s.Cell(row + 0, COL_TYPE).Value = (type ?? string.Empty).Trim();
            s.Cell(row + 1, COL_TYPE).Value = (insideOutside ?? string.Empty).Trim();

            // E: Material rows + Acc to Standard + Thickness
            string mat1_c1 = Attr(obj, true, "Material 1 column 1", "Material 1 col 1");
            string mat1_c2 = Attr(obj, true, "Material 1 column 2", "Material 1 col 2");

            int off = 0;
            if (!string.IsNullOrWhiteSpace(mat1_c1))
            {
                if (!string.IsNullOrWhiteSpace(mat1_c2))
                    s.Cell(row + off, COL_DESC).Value = mat1_c1 + " - " + mat1_c2;
                else
                    s.Cell(row + off, COL_DESC).Value = mat1_c1;
                off++;
            }

            string mat2_c1 = Attr(obj, true, "Material 2 column 1", "Material 2 col 1");
            string mat2_c2 = Attr(obj, true, "Material 2 column 2", "Material 2 col 2");
            if (!string.IsNullOrWhiteSpace(mat2_c1))
            {
                if (!string.IsNullOrWhiteSpace(mat2_c2))
                    s.Cell(row + off, COL_DESC).Value = mat2_c1 + " - " + mat2_c2;
                else
                    s.Cell(row + off, COL_DESC).Value = mat2_c1;
                off++;
            }

            string std = FirstNonEmpty(
                Attr(obj, true, "Acc to Standard", "Acc to standard", "Standard"),
                AttrLike(obj, "ansi"), AttrLike(obj, "asme"), AttrLike(obj, "dn")
            );
            if (!string.IsNullOrWhiteSpace(std))
            {
                s.Cell(row + off, COL_DESC).Value = std;
                off++;
            }

            string thick = FirstNonEmpty(
                Attr(obj, true, "Thickness"),
                AttrLike(obj, "thickness")
            );
            if (!string.IsNullOrWhiteSpace(thick))
            {
                s.Cell(row + off, COL_DESC).Value = "THICKNESS = " + thick;
            }

            // Sizes / Class / Code
            string szMin = FirstNonEmpty(
                Attr(obj, true, "Size min", "Size (Min)", "Size Min"),
                AttrLike(obj, "min size")
            );
            string szMax = FirstNonEmpty(
                Attr(obj, true, "Size max", "Size (Max)", "Size Max"),
                AttrLike(obj, "max size")
            );
            string cls = FirstNonEmpty(
                Attr(obj, true, "Class"),
                AttrLike(obj, "class")
            );
            string code = FirstNonEmpty(
                Attr(obj, true, "Code", "Item code", "Part code"),
                Attr(obj, true, "Comment", "Item Tag", "Tag")
            );

            s.Cell(row + 0, COL_SIZE_MIN).Value = szMin;
            s.Cell(row + 0, COL_SIZE_MAX).Value = szMax;
            s.Cell(row + 0, COL_SCHCLASS).Value = cls;
            s.Cell(row + 0, COL_CODE).Value = code;

            ClearRow(s, row + 3);
            return row + 4;
        }

        // ---------- VALVE ----------
        // Python write_valve_row naming:
        //   valve_type = Type
        //   A(row)   = first part before comma
        //   A(row+1) = joined remaining parts if there are > 2 parts
        // ---------- VALVE (MODULE 2: Pipe Class Summary, Python-style MATERIAL-STANDARD) ----------
        // ---------- VALVE (MODULE 2: Pipe Class Summary, Python-style MATERIAL-STANDARD) ----------
        private static int WriteValveRows(IXLWorksheet s, int row, ObjectItem obj)
        {
            // ===== Column A: Valve type (same as Python) =====
            string rawType = FirstNonEmpty(
                Attr(obj, true, "Type", "Valve Type", "Valve type")
            );

            // If Type is just a numeric code (1765 etc.), prefer comments instead
            if (LooksLikeNumericCode(rawType))
            {
                rawType = FirstNonEmpty(
                    Attr(obj, true, "Additional Comment", "Additional comment"),
                    Attr(obj, true, "Comment", "Item Tag", "Tag"),
                    rawType
                );
            }

            string line1 = string.Empty;
            string line2 = string.Empty;

            if (!string.IsNullOrWhiteSpace(rawType))
            {
                var parts = rawType
                    .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(p => p.Trim())
                    .ToArray();

                if (parts.Length > 0)
                    line1 = parts[0];

                // Python: second line only if there are > 2 parts
                if (parts.Length > 2)
                    line2 = string.Join(", ", parts.Skip(1));
            }

            s.Cell(row + 0, COL_TYPE).Value = line1 ?? string.Empty;
            s.Cell(row + 1, COL_TYPE).Value = line2 ?? string.Empty;

            // ===== Column E: MATERIAL-STANDARD block =====
            // BODY / SEAT / DISC / SHAFT / OPERATION etc.
            // then Additional specification 1 + 2
            // then Piping connection / Mounting line
            int off = 0;

            // helper to add one "LABEL : VALUE" line into column E
            void AddLabelValue(string label, string value)
            {
                if (string.IsNullOrWhiteSpace(label)) return;

                label = label.Trim();
                value = (value ?? string.Empty).Trim();

                string text = string.IsNullOrWhiteSpace(value)
                    ? label
                    : label + " : " + value;

                s.Cell(row + off, COL_DESC).Value = text;
                off++;
            }

            // 1..4 description/material lines
            for (int i = 1; i <= 4; i++)
            {
                // LABEL from Description i (C1) or Material i column 1
                string label = FirstNonEmpty(
                    Attr(obj, true, $"Description {i} (C1)"),
                    Attr(obj, true, $"Description{i} (C1)"),
                    Attr(obj, true, $"Material {i} column 1"),
                    Attr(obj, true, $"Material {i} col 1")
                );

                // VALUE from Description i (C3) → (C2) → Material i column 3/2
                string value = FirstNonEmpty(
                    Attr(obj, true, $"Description {i} (C3)"),
                    Attr(obj, true, $"Description{i} (C3)"),
                    Attr(obj, true, $"Description {i} (C2)"),
                    Attr(obj, true, $"Description{i} (C2)"),
                    Attr(obj, true, $"Material {i} column 3"),
                    Attr(obj, true, $"Material {i} col 3"),
                    Attr(obj, true, $"Material {i} column 2"),
                    Attr(obj, true, $"Material {i} col 2")
                );

                AddLabelValue(label, value);
            }

            // ----- Additional specification 1 / 2 -----
            string add1 = FirstNonEmpty(
                Attr(obj, true, "Additional specification 1", "Additional Specification 1", "Remark 1")
            );
            if (!string.IsNullOrWhiteSpace(add1))
            {
                s.Cell(row + off, COL_DESC).Value = add1.Trim();
                off++;
            }

            string add2 = FirstNonEmpty(
                Attr(obj, true, "Additional specification 2", "Additional Specification 2", "Remark 2")
            );
            if (!string.IsNullOrWhiteSpace(add2))
            {
                s.Cell(row + off, COL_DESC).Value = add2.Trim();
                off++;
            }

            // ----- Piping connection / mounting line -----
            // This is where "MOUNTING BETWEEN FLANGES ASME B16.5 CLASS 150#" comes from
            string mounting = FirstNonEmpty(
                Attr(obj, true, "Piping connection 1", "Piping Connection 1"),
                Attr(obj, true, "Piping connection"),
                Attr(obj, true, "End connection", "Ends", "Connection"),
                AttrLike(obj, "MOUNTING BETWEEN FLANGES"),
                AttrLike(obj, "MOUNTING"),
                AttrLike(obj, "FLANGE")
            );
            if (!string.IsNullOrWhiteSpace(mounting))
            {
                s.Cell(row + off, COL_DESC).Value = mounting.Trim();
                off++;
            }

            // (IMPORTANT) Do NOT ClearRowRange here – that was wiping some of these lines.
            // If you want to clean only numeric cells below, clear J/K/L/N for the extra rows:
            for (int r = row + 1; r <= row + 3; r++)
            {
                s.Cell(r, COL_SIZE_MIN).Clear();  // J
                s.Cell(r, COL_SIZE_MAX).Clear();  // K
                s.Cell(r, COL_SCHCLASS).Clear();  // L
                s.Cell(r, COL_CODE).Clear();      // N
            }

            // ===== NPS / CLASS / CODE =====
            string szMin = FirstNonEmpty(
                Attr(obj, true, "Size min", "Size (Min)", "Size Min", "DN min", "DN Min"),
                AttrLike(obj, "min size")
            );
            string szMax = FirstNonEmpty(
                Attr(obj, true, "Size max", "Size (Max)", "Size Max", "DN max", "DN Max"),
                AttrLike(obj, "max size")
            );
            s.Cell(row + 0, COL_SIZE_MIN).Value = szMin;
            s.Cell(row + 0, COL_SIZE_MAX).Value = szMax;

            string designCode = FirstNonEmpty(
                Attr(obj, true, "Design code", "Design Code"),
                Attr(obj, true, "DIN/ANSI", "DIN / ANSI"),
                Attr(obj, true, "Standard"),
                AttrLike(obj, "design code"),
                AttrLike(obj, "din"),
                AttrLike(obj, "ansi")
            );
            s.Cell(row + 0, COL_SCHCLASS).Value = designCode;

            string code = FirstNonEmpty(
                Attr(obj, true, "Code", "Item code", "Part code"),
                Attr(obj, true, "Comment", "Item Tag", "Tag")
            );
            s.Cell(row + 0, COL_CODE).Value = code;

            // One valve block still uses 4 template rows
            return row + 4;
        }



        // ========== EB traversal & helpers ==========

        private static IEnumerable<ObjectItem> WalkDeep(ObjectItem root)
        {
            var stack = new Stack<ObjectItem>();
            if (root != null) stack.Push(root);

            while (stack.Count > 0)
            {
                var cur = stack.Pop();
                yield return cur;

                IEnumerable<ObjectItem> kids;
                try { kids = (cur != null && cur.Children != null) ? cur.Children : Enumerable.Empty<ObjectItem>(); }
                catch { kids = Enumerable.Empty<ObjectItem>(); }

                foreach (var k in kids) stack.Push(k);
            }
        }

        private static List<ObjectItem> CollectItemsForClass(EbApp app, string pipeClass)
        {
            var list = new List<ObjectItem>();
            if (app == null || string.IsNullOrWhiteSpace(pipeClass)) return list;

            ObjectItem catalogs = null;
            try { catalogs = app.Folders != null ? app.Folders.Catalogs : null; } catch { }
            if (catalogs == null) return list;

            // Catalogs → JLE → Materials → {Bolts & Nuts, Pipe & Fittings, Valves}
            var jle = GetChild(catalogs, "JLE");
            var materials = jle != null ? GetChild(jle, "Materials") : null;

            var roots = new List<ObjectItem>();
            if (materials != null)
            {
                var bolts = GetChild(materials, "Bolts & Nuts");
                var pf = GetChild(materials, "Pipe & Fittings");
                var valves = GetChild(materials, "Valves");

                if (bolts != null) roots.Add(bolts);
                if (pf != null) roots.Add(pf);
                if (valves != null) roots.Add(valves);
            }

            // Do NOT fall back to whole Catalogs tree – keeps generation fast.
            if (roots.Count == 0) return list;

            foreach (var root in roots)
            {
                foreach (ObjectItem obj in WalkDeep(root))
                {
                    if (ObjectBelongsToClass(obj, pipeClass))
                        list.Add(obj);
                }
            }
            return list;
        }

        private static bool ObjectBelongsToClass(ObjectItem obj, string pipeClass)
        {
            if (obj == null || string.IsNullOrWhiteSpace(pipeClass)) return false;
            var attrs = obj.Attributes;
            if (attrs == null) return false;

            string want = pipeClass.Trim();
            foreach (AttributeItem a in attrs)
            {
                if (a == null) continue;
                string an = a.Name ?? string.Empty;
                for (int i = 0; i < ClassAttrCandidates.Length; i++)
                {
                    if (string.Equals(an, ClassAttrCandidates[i], StringComparison.OrdinalIgnoreCase))
                    {
                        string val = SafeAttrValue(a);
                        if (!string.IsNullOrWhiteSpace(val) &&
                            val.IndexOf(want, StringComparison.OrdinalIgnoreCase) >= 0)
                            return true;
                    }
                }
            }
            return false;
        }

        private static bool LooksLikeBoltsOrNuts(ObjectItem obj)
        {
            string name = SafeName(obj).ToLowerInvariant();
            if (name.Contains("bolt") || name.Contains("nut")) return true;

            string type = Attr(obj, true, "Type");
            if (!string.IsNullOrWhiteSpace(type))
            {
                string t = type.ToLowerInvariant();
                if (t.Contains("bolt") || t.Contains("nut")) return true;
            }

            try
            {
                ObjectItem p = obj.Parent;
                while (p != null)
                {
                    string pn = SafeName(p).ToLowerInvariant();
                    if (pn.Contains("bolt") || pn.Contains("nut")) return true;
                    p = p.Parent;
                }
            }
            catch { }
            return false;
        }

        private static bool LooksLikeGasket(ObjectItem obj)
        {
            string name = SafeName(obj).ToLowerInvariant();
            if (name.Contains("gasket")) return true;

            string type = Attr(obj, true, "Type");
            if (!string.IsNullOrWhiteSpace(type) && type.ToLowerInvariant().Contains("gasket")) return true;

            try
            {
                ObjectItem p = obj.Parent;
                while (p != null)
                {
                    string pn = SafeName(p).ToLowerInvariant();
                    if (pn.Contains("gasket")) return true;
                    p = p.Parent;
                }
            }
            catch { }
            return false;
        }

        private static bool LooksLikeValve(ObjectItem obj)
        {
            string name = SafeName(obj).ToLowerInvariant();
            if (name.Contains("valve")) return true;

            string type = Attr(obj, true, "Type");
            if (!string.IsNullOrWhiteSpace(type) && type.ToLowerInvariant().Contains("valve")) return true;

            try
            {
                ObjectItem p = obj.Parent;
                while (p != null)
                {
                    string pn = SafeName(p).ToLowerInvariant();
                    if (pn.Contains("valve")) return true;
                    p = p.Parent;
                }
            }
            catch { }
            return false;
        }

        // ---------- Attribute helpers ----------

        private static string AttrExact(ObjectItem obj, params string[] names)
        {
            if (obj == null || names == null || names.Length == 0) return string.Empty;
            var attrs = obj.Attributes;
            if (attrs == null) return string.Empty;

            foreach (string target in names)
            {
                if (string.IsNullOrWhiteSpace(target)) continue;

                foreach (AttributeItem a in attrs)
                {
                    if (a == null) continue;
                    string an = a.Name ?? string.Empty;
                    if (string.Equals(an, target, StringComparison.OrdinalIgnoreCase))
                        return SafeAttrValue(a);
                }
            }
            return string.Empty;
        }

        private static string Attr(ObjectItem obj, bool fuzzy, params string[] names)
        {
            string exact = AttrExact(obj, names);
            if (!fuzzy || !string.IsNullOrWhiteSpace(exact)) return exact;

            foreach (var name in names)
            {
                string like = AttrLike(obj, name);
                if (!string.IsNullOrWhiteSpace(like)) return like;
            }
            return string.Empty;
        }

        private static string AttrLike(ObjectItem obj, string needle)
        {
            if (obj == null || obj.Attributes == null) return string.Empty;
            string n = NormalizeKey(needle);
            if (n.Length == 0) return string.Empty;

            foreach (AttributeItem a in obj.Attributes)
            {
                try
                {
                    if (a == null || string.IsNullOrWhiteSpace(a.Name)) continue;
                    string an = NormalizeKey(a.Name);
                    if (an.IndexOf(n, StringComparison.Ordinal) >= 0)
                        return SafeAttrValue(a);
                }
                catch { }
            }
            return string.Empty;
        }

        private static void TryAddFromNamedAttributes(ObjectItem obj, HashSet<string> into, IEnumerable<string> names)
        {
            if (obj == null || obj.Attributes == null || names == null) return;

            var nameSet = new HashSet<string>(
                names.Where(x => !string.IsNullOrWhiteSpace(x)).Select(x => x.Trim()),
                StringComparer.OrdinalIgnoreCase);

            foreach (AttributeItem a in obj.Attributes)
            {
                try
                {
                    if (a == null || string.IsNullOrWhiteSpace(a.Name)) continue;
                    if (!nameSet.Contains(a.Name.Trim())) continue;

                    string raw = SafeAttrValue(a);
                    AddTokens(raw, into);
                }
                catch { }
            }
        }

        private static void TryAddFromAttributesContaining(ObjectItem obj, HashSet<string> into, string contains)
        {
            if (obj == null || obj.Attributes == null) return;
            string needle = (contains ?? string.Empty).Trim().ToLowerInvariant();
            if (needle.Length == 0) return;

            foreach (AttributeItem a in obj.Attributes)
            {
                try
                {
                    if (a == null || string.IsNullOrWhiteSpace(a.Name)) continue;
                    string an = (a.Name ?? string.Empty).Trim().ToLowerInvariant();
                    if (an.IndexOf(needle, StringComparison.Ordinal) < 0) continue;

                    string raw = SafeAttrValue(a);
                    AddTokens(raw, into);
                }
                catch { }
            }
        }

        private static void AddTokens(string raw, HashSet<string> into)
        {
            if (string.IsNullOrWhiteSpace(raw)) return;
            string[] parts = raw.Split(new char[] { ',', ';', '/', '\\', '\r', '\n', '\t', ' ' },
                                       StringSplitOptions.RemoveEmptyEntries);
            foreach (string p in parts)
            {
                string token = Normalize(p);
                if (IsPipeClassCode(token)) into.Add(token);
            }
        }

        private static string SafeAttrValue(AttributeItem a)
        {
            try { return a != null && a.Value != null ? (a.Value.ToString() ?? string.Empty) : string.Empty; }
            catch { return string.Empty; }
        }

        private static string SafeName(ObjectItem obj)
            => obj != null ? (obj.Name ?? string.Empty) : string.Empty;

        // ======== Description & Remark helpers ========
        private static string GetDesc(ObjectItem obj, int index, int col)
        {
            string i = index.ToString(CultureInfo.InvariantCulture);
            string c = col.ToString(CultureInfo.InvariantCulture);

            return FirstNonEmpty(
                Attr(obj, true, $"Description {i} (C{c})"),
                Attr(obj, true, $"Description{i} (C{c})"),
                Attr(obj, true, $"Description {i} C{c}"),
                Attr(obj, true, $"Desc {i} (C{c})"),
                Attr(obj, true, $"Desc{i} (C{c})"),
                Attr(obj, true, $"C{c} Description {i}"),
                Attr(obj, true, $"Description {i} C {c}"),
                AttrLike(obj, $"description{i}"),
                AttrLike(obj, $"desc {i} c{c}")
            );
        }

        private static string GetRemark(ObjectItem obj, int index)
        {
            string i = index.ToString(CultureInfo.InvariantCulture);
            return FirstNonEmpty(
                Attr(obj, true, $"Remark {i}"),
                Attr(obj, true, $"Remarks {i}"),
                Attr(obj, true, $"Additional specification {i}"),
                AttrLike(obj, $"remark {i}"),
                AttrLike(obj, $"additional {i}")
            );
        }

        // ---------- Folder helpers ----------

        private static string NormalizeFolderKey(string s)
        {
            var t = (s ?? string.Empty).ToLowerInvariant();
            char[] buf = new char[t.Length];
            int j = 0;
            for (int i = 0; i < t.Length; i++)
            {
                char ch = t[i];
                if (ch == ' ' || ch == '/' || ch == '\\' || ch == '-' || ch == '_' ||
                    ch == '.' || ch == ':' || ch == '(' || ch == ')' || ch == ',' ||
                    ch == '\'' || ch == '\"' || ch == '&')
                    continue;
                buf[j++] = ch;
            }
            return new string(buf, 0, j);
        }

        private static ObjectItem GetChild(ObjectItem parent, string childName)
        {
            if (parent == null || parent.Children == null) return null;
            string want = NormalizeFolderKey(childName);
            try
            {
                foreach (ObjectItem c in parent.Children)
                {
                    if (NormalizeFolderKey(c != null ? c.Name ?? "" : "") == want) return c;
                }
            }
            catch { }
            return null;
        }

        // ---------- Misc utilities ----------

        private static void ClearRow(IXLWorksheet s, int row)
        {
            s.Cell(row, COL_TYPE).Clear();
            s.Cell(row, COL_DESC).Clear();
            s.Cell(row, COL_SIZE_MIN).Clear();
            s.Cell(row, COL_SIZE_MAX).Clear();
            s.Cell(row, COL_SCHCLASS).Clear();
            s.Cell(row, COL_CODE).Clear();
        }

        private static void ClearRowRange(IXLWorksheet s, int r1, int r2)
        {
            for (int r = r1; r <= r2; r++) ClearRow(s, r);
        }

        private static void LeftAlign(IXLCell c)
        {
            try { c.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left; } catch { }
        }

        private static string JoinIf(string a, string mid, string b)
        {
            if (string.IsNullOrWhiteSpace(a)) return string.IsNullOrWhiteSpace(b) ? string.Empty : b;
            if (string.IsNullOrWhiteSpace(b)) return a;
            return a + mid + b;
        }

        private static string FirstNonEmpty(params string[] vals)
        {
            if (vals == null) return string.Empty;
            foreach (var v in vals) if (!string.IsNullOrWhiteSpace(v)) return v;
            return string.Empty;
        }

        private static string Normalize(string s) => (s ?? string.Empty).Trim().ToUpperInvariant();

        private static string NormalizeKey(string s)
        {
            var t = (s ?? string.Empty).ToLowerInvariant();
            char[] buf = new char[t.Length];
            int j = 0;
            for (int i = 0; i < t.Length; i++)
            {
                char ch = t[i];
                if (ch == ' ' || ch == '/' || ch == '\\' || ch == '-' || ch == '_' || ch == '.' || ch == ':')
                    continue;
                buf[j++] = ch;
            }
            return new string(buf, 0, j);
        }

        private static bool LooksLikeNumericCode(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;
            s = s.Trim();
            for (int i = 0; i < s.Length; i++)
            {
                if (!char.IsDigit(s[i])) return false;
            }
            return s.Length > 0;
        }

        private static bool IsPipeClassCode(string s)
            => !string.IsNullOrWhiteSpace(s) && PipeClassCodeRx.IsMatch(s.Trim());

        private static string ExtractRatingToken(string code)
        {
            if (string.IsNullOrWhiteSpace(code)) return string.Empty;
            int j = code.IndexOf('J');
            return j > 0 ? code.Substring(0, j) : string.Empty;
        }

        private static bool IsAsmeRating(string ratingToken)
        {
            double val;
            if (double.TryParse(ratingToken, NumberStyles.Any, CultureInfo.InvariantCulture, out val))
                return val >= 150.0;
            return false;
        }

        private static string FirstChar(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "0";
            string t = s.Trim();
            return t.Substring(0, 1);
        }

        private static int GroupSortKey(string k)
        {
            if (k == "1") return 1;
            if (k == "2") return 2;
            if (k == "3") return 3;
            if (k == "4") return 4;
            return 99;
        }

        private static string GroupHeader(string key)
        {
            if (string.IsNullOrWhiteSpace(key)) return string.Empty;
            string t = key.Trim();
            if (t == "1") return "PIPE";
            if (t == "2") return "FITTINGS";
            if (t == "3") return "BRANCH FITTINGS";
            if (t == "4") return "FLANGES";
            return string.Empty;
        }

        private static string SanitizeFileName(string name)
        {
            string n = string.IsNullOrEmpty(name) ? "PIPECLASS" : name;
            foreach (char bad in Path.GetInvalidFileNameChars()) n = n.Replace(bad, '_');
            return n;
        }
    }

    internal sealed class PPView
    {
        public ObjectItem Obj = null!;
        public string Code = "";
        public string GroupKey = "";
    }

    internal sealed class StringLogicalComparer : IComparer<string>
    {
        public static readonly StringLogicalComparer Instance = new StringLogicalComparer();
        public int Compare(string x, string y)
        {
            if (x == null && y == null) return 0;
            if (x == null) return -1;
            if (y == null) return 1;
            return StrCmpLogicalW(x, y);
        }

        [DllImport("shlwapi.dll", CharSet = CharSet.Unicode)]
        private static extern int StrCmpLogicalW(string x, string y);
    }

    /// <summary>Grouping container for UI (supports foreach).</summary>
    public sealed class ClassGroups : IEnumerable<string>
    {
        public List<string> Asme { get; set; } = new List<string>();
        public List<string> Din { get; set; } = new List<string>();
        public List<string> Other { get; set; } = new List<string>();
        public List<string> Ordered { get; set; } = new List<string>();

        public IEnumerator<string> GetEnumerator()
        {
            return (Ordered ?? new List<string>()).GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
#nullable disable

