using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using Aucotec.EngineeringBase.Client.Runtime;
using EbApp = Aucotec.EngineeringBase.Client.Runtime.Application;

using System.Diagnostics;
using System.IO;
using System.Reflection;

namespace JJ_Lurgi_Piping_EB
{
    static class Diag
    {
        public static string Info()
        {
            try
            {
                Assembly asm = Assembly.GetExecutingAssembly();
                string path = asm.Location;
                string ver = asm.GetName().Version != null ? asm.GetName().Version.ToString() : "n/a";
                string write = File.GetLastWriteTime(path).ToString("yyyy-MM-dd HH:mm:ss");
                return Path.GetFileName(path) + "  |  v" + ver + "\nPath: " + path + "\nWritten: " + write;
            }
            catch (Exception ex) { return "Build info unavailable: " + ex.Message; }
        }

        public static void Trace(string msg)
        {
            try { Debug.WriteLine("[JJLEM] " + msg); } catch { }
        }
    }

    public partial class MainWindow : Window
    {
        private readonly EbApp _app;
        private readonly List<ItemRow> _rows = new List<ItemRow>();
        private readonly HashSet<ObjectItem> _selectedItems = new HashSet<ObjectItem>();

        // Keep this so XAML can open in designer; runtime EB passes real app.
        public MainWindow() : this(app: null) { }

        public MainWindow(EbApp app)
        {
            _app = app ?? null;
            InitializeComponent();
            Loaded += MainWindow_Loaded;
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            string info = Diag.Info();
            if (TxtBuildInfo != null) TxtBuildInfo.Text = info;
            Diag.Trace("Window loaded\n" + info);

            try
            {
                TxtStatus.Text = "Loading…";

                if (_app == null)
                {
                    TxtStatus.Text = "Designer mode (no EB instance).";
                    return;
                }

                ObjectItem catalogs = null;
                try { catalogs = _app.Folders?.Catalogs; } catch { }

                if (catalogs == null)
                {
                    TxtStatus.Text = "No 'Catalogs' folder found.";
                    return;
                }

                // Prefer Catalogs → JLE → Materials as the root.
                var materialsRoot = FindMaterialsRoot(catalogs);

                if (materialsRoot != null)
                {
                    LoadFromMaterials(materialsRoot);
                }
                else
                {
                    // Fallback: whole Catalogs (still grouped by top-level folder)
                    LoadFromRoot(catalogs);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Init failed: " + ex.Message, "JJ Lurgi Piping EB");
            }
        }

        /// <summary>Try to find Catalogs → JLE → Materials. Returns null if not found.</summary>
        private ObjectItem FindMaterialsRoot(ObjectItem catalogs)
        {
            if (catalogs == null || catalogs.Children == null) return null;

            ObjectItem jle = FindChildByName(catalogs, "JLE");
            if (jle == null) return null;

            ObjectItem materials = FindChildByName(jle, "Materials");
            return materials;
        }

        private ObjectItem FindChildByName(ObjectItem parent, string name)
        {
            if (parent == null || parent.Children == null || string.IsNullOrWhiteSpace(name)) return null;
            try
            {
                foreach (ObjectItem c in parent.Children)
                {
                    if (c != null &&
                        string.Equals(c.Name ?? string.Empty, name, StringComparison.OrdinalIgnoreCase))
                        return c;
                }
            }
            catch { }
            return null;
        }

        /// <summary>
        /// Load just from Materials → {Bolts & Nuts, Pipe & Fittings, Valves}, grouped by folder name.
        /// </summary>
        private void LoadFromMaterials(ObjectItem materialsRoot)
        {
            _rows.Clear();

            var categories = new[] { "Bolts & Nuts", "Pipe & Fittings", "Valves" };
            var temp = new List<ItemRow>();

            foreach (string catName in categories)
            {
                var folder = FindChildByName(materialsRoot, catName);
                if (folder == null) continue;

                CollectLeafItems_Lightweight(folder, temp, catName);
            }

            for (int i = 0; i < temp.Count; i++)
            {
                var r = temp[i];
                r.Index = i + 1; // not shown, but kept if ever needed
                r.IsSelected = _selectedItems.Contains(r.Object);
            }

            _rows.AddRange(temp);
            BindGridAndGroup("Materials");
        }

        /// <summary>
        /// Fallback: load all leaf items under a root and group by the first-level folder.
        /// </summary>
        private void LoadFromRoot(ObjectItem root)
        {
            _rows.Clear();
            var temp = new List<ItemRow>();

            foreach (ObjectItem child in SafeChildren(root))
            {
                string category = child.Name ?? string.Empty;
                CollectLeafItems_Lightweight(child, temp, category);
            }

            for (int i = 0; i < temp.Count; i++)
            {
                var r = temp[i];
                r.Index = i + 1;
                r.IsSelected = _selectedItems.Contains(r.Object);
            }

            _rows.AddRange(temp);
            BindGridAndGroup(root.Name ?? string.Empty);
        }

        private void BindGridAndGroup(string contextName)
        {
            GridItems.ItemsSource = null;
            GridItems.DataContext = null;

            GridItems.ItemsSource = _rows;
            GridItems.DataContext = _rows;

            // Group by Category for the DataGrid's default view
            var view = CollectionViewSource.GetDefaultView(GridItems.ItemsSource);
            if (view != null && view.CanGroup)
            {
                view.GroupDescriptions.Clear();
                view.GroupDescriptions.Add(new PropertyGroupDescription("Category"));
            }

            UpdateStatus(contextName);
        }

        private static bool HasChildren(ObjectItem obj)
        {
            try { return obj != null && obj.Children != null && obj.Children.Count > 0; }
            catch { return false; }
        }

        private static IEnumerable<ObjectItem> SafeChildren(ObjectItem obj)
        {
            try
            {
                return (obj != null && obj.Children != null)
                    ? obj.Children
                    : Enumerable.Empty<ObjectItem>();
            }
            catch
            {
                return Enumerable.Empty<ObjectItem>();
            }
        }

        // ----------- SUPER FAST: no attributes on load at all -----------
        private void CollectLeafItems_Lightweight(ObjectItem folder, List<ItemRow> into, string category)
        {
            foreach (ObjectItem child in SafeChildren(folder))
            {
                if (!HasChildren(child))
                {
                    // Do NOT touch child.Attributes here – that’s the slow part.
                    // Just use the ObjectItem itself; ItemRow will use Name as Code.
                    var attrsLite = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    into.Add(new ItemRow(child, attrsLite, category));
                }
                else
                {
                    CollectLeafItems_Lightweight(child, into, category);
                }
            }
        }
        // ---------------------------------------------------------------

        private Dictionary<string, string> ReadAllAttributes(ObjectItem obj)
        {
            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (obj == null) return dict;

            try
            {
                var attrs = obj.Attributes;
                if (attrs != null)
                {
                    foreach (AttributeItem a in attrs)
                    {
                        string key = a != null ? (a.Name ?? a.ToString()) : "ATTR";
                        string val = a != null && a.Value != null ? a.Value.ToString() : string.Empty;
                        if (!dict.ContainsKey(key)) dict[key] = val;
                    }
                }
            }
            catch { }
            return dict;
        }

        private void ChkSelectAll_Checked(object sender, RoutedEventArgs e)
        {
            bool on = ChkSelectAll.IsChecked == true;
            foreach (ItemRow r in _rows) r.IsSelected = on;

            SyncSelectionFromRows();
            GridItems.Items.Refresh();
            UpdateStatus(string.Empty);
        }

        private void BtnGenerate_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show(Diag.Info(), "JJ Lurgi Piping EB – Loaded DLL");

            try
            {
                if (_app == null)
                {
                    MessageBox.Show("No EB instance (designer mode).", "Datasheets");
                    return;
                }

                SyncSelectionFromRows();

                // Load full attributes only now, for selected rows
                List<ItemRow> picked = BuildPickedListWithAttributes();
                if (picked == null || picked.Count == 0)
                {
                    MessageBox.Show("No items selected. Tick rows and try again.", "Datasheets");
                    return;
                }

                Diag.Trace("Generate clicked. Items picked: " + picked.Count);
                string msg = DatasheetExporter.Generate(_app, picked);
                MessageBox.Show(msg, "Datasheets");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Generation failed: " + ex.Message, "Datasheets");
            }
        }

        private void BtnReset_Click(object sender, RoutedEventArgs e)
        {
            _selectedItems.Clear();
            foreach (ItemRow r in _rows) r.IsSelected = false;
            GridItems.Items.Refresh();
            UpdateStatus(string.Empty);
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e) { Close(); }

        private void SyncSelectionFromRows()
        {
            foreach (ItemRow r in _rows)
            {
                if (r.IsSelected) _selectedItems.Add(r.Object);
                else _selectedItems.Remove(r.Object);
            }
        }

        // Load full attributes only for the rows we’re going to export
        private List<ItemRow> BuildPickedListWithAttributes()
        {
            var picked = new List<ItemRow>();

            // Rows ticked in the grid – preserve category
            picked.AddRange(_rows.Where(r => r.IsSelected).Select(r =>
                new ItemRow(r.Object, ReadAllAttributes(r.Object), r.Category)));

            // Items tracked via _selectedItems (if any)
            foreach (ObjectItem oi in _selectedItems)
            {
                bool already = picked.Any(p => ReferenceEquals(p.Object, oi));
                if (!already)
                {
                    string cat = _rows.FirstOrDefault(x => ReferenceEquals(x.Object, oi))?.Category ?? string.Empty;
                    picked.Add(new ItemRow(oi, ReadAllAttributes(oi), cat));
                }
            }

            return picked.GroupBy(p => p.Object).Select(g => g.First()).ToList();
        }

        private void UpdateStatus(string contextName)
        {
            string ctx = !string.IsNullOrEmpty(contextName) ? (" from '" + contextName + "'") : "";
            TxtStatus.Text = "Loaded " + _rows.Count + " item(s)" + ctx + ".  Selected (total): " + _selectedItems.Count + ".";
        }

        private void GridItems_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
        }
    }
}
