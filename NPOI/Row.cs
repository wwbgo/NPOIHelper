using System.Collections.Generic;

namespace NPOIHelper {
    public class Row {
        public Row() {
            Columns = new List<Column>();
        }
        public Row(IList<Column> cols) {
            Columns = cols;
        }
        public bool IsHead { get; set; }
        public IList<Column> Columns { get; set; }
    }
}
