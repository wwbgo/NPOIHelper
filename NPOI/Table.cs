using System.Collections.Generic;

namespace NPOIHelper {
    public class Table {
        public Table() {
            Rows = new List<Row>();
        }
        public Table(string title) {
            Title = title;
            Rows = new List<Row>();
        }
        public Table(IList<Row> rows) {
            Rows = rows;
        }
        public Table(string title,IList<Row> rows) {
            Title = title;
            Rows = rows;
        }
        public string Title { get; set; }
        public IList<Row> Rows { get; set; }
        public Dictionary<int,string[]> BandColumns { get; set; }
    }
}
