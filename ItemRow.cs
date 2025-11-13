using Aucotec.EngineeringBase.Client.Runtime;
using System;
using System.Collections.Generic;

namespace JJ_Lurgi_Piping_EB
{
    /// <summary>View-model row for one catalog item shown in the grid.</summary>
    public class ItemRow
    {
        public ObjectItem Object { get; private set; }

        // Not shown in UI now, but kept if needed later.
        public int Index { get; set; }

        // Checkbox in the grid
        public bool IsSelected { get; set; }

        // Code shown in the grid (falls back to Name if no Code attr)
        public string Code { get; private set; }

        // Original EB name (not shown in grid)
        public string Name { get; private set; }

        // Just kept for stats; not shown
        public int AttributeCount { get; private set; }

        /// <summary>
        /// Logical group/category for UI, e.g. "Bolts & Nuts", "Pipe & Fittings", "Valves".
        /// Used for grouping and for the "Category" column.
        /// </summary>
        public string Category { get; private set; }

        /// <summary>
        /// Item type, available if we ever want to show it.
        /// </summary>
        public string Type { get; private set; }

        public Dictionary<string, string> Attributes { get; private set; }

        public ItemRow(ObjectItem obj, Dictionary<string, string> attributes)
        {
            if (obj == null) throw new ArgumentNullException("obj");
            Object = obj;

            Attributes = attributes ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            Name = obj.Name ?? string.Empty;

            string code;
            if (!Attributes.TryGetValue("Code", out code)) code = string.Empty;
            Code = string.IsNullOrWhiteSpace(code) ? Name : code;

            string type;
            if (!Attributes.TryGetValue("Type", out type)) type = string.Empty;
            Type = type ?? string.Empty;

            Category = string.Empty;
            AttributeCount = Attributes.Count;
        }

        public ItemRow(ObjectItem obj, Dictionary<string, string> attributes, string category)
            : this(obj, attributes)
        {
            Category = category ?? string.Empty;
        }
    }
}
