using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PSDocs.Models
{
    public sealed class TableColumn
    {
        /// <summary>
        /// The name of the column.
        /// </summary>
        public string Name { get; set; }

        public Func<object> Expression { get; set; }


        public static implicit operator TableColumn(Hashtable hashtable)
        {
            var column = new TableColumn();

            // Build index to allow mapping
            var index = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

            foreach (DictionaryEntry entry in hashtable)
            {
                index.Add(entry.Key.ToString(), entry.Value);
            }

            // Start loading matching values

            object value;

            if (index.TryGetValue("label", out value))
            {
                column.Name = (string)value;
            }

            if (index.TryGetValue("name", out value))
            {
                column.Name = (string)value;
            }

            if (index.TryGetValue("expression", out value))
            {
                column.Expression = (Func<object>)value;
            }

            return column;
        }

        public static implicit operator TableColumn(string name)
        {
            var column = new TableColumn { Name = name };

            return column;
        }
    }
}
