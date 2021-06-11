using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FamilySorting
{
    public class TableEntry
    {
        public string Key { get; set; }
        public string CategoryR { get; set; }
        public string Discipline { get; set; }
        public string Gategory { get; set; }
        public string SubCategory { get; set; }
        public string PathToFamily { get; set; }
        public string PathToInstruction { get; set; }

        public List<TableEntry> GetTableEntries()
        {
            var output = new List<TableEntry>();


            return output;
        }
    }
}
