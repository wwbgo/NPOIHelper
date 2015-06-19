using System;
using System.Collections.Generic;
using System.Data;
using NPOIHelper;

namespace test {
    class NPOI {
        static void Main(string[] args) {
            var time = DateTime.Now;
            Console.WriteLine(time);

            //export(read());
            //Export.RenderToExcelForXlsx(read(),"档案信息").SaveToFile("e:/test.xlsx");
            ExportForXlsx();

            var time1 = DateTime.Now;
            var t = time1 - time;
            Console.WriteLine(time1);
            Console.WriteLine(t.TotalMilliseconds / 1000);
            Console.ReadKey();
        }

        public static void export(DataTable data) {
            if (data == null || data.Rows.Count <= 0) {
                return;
            }
            var table = new Table {
                Title = string.Format("档案信息_{0}",DateTime.Now.ToString("yyyyMMddHHmmssfff")),
                BandColumns = new Dictionary<int,string[]> { { 1,new[] { "testA","testB","testC" } },{ 3,new[] { "testD","testE","testF" } } }
            };
            var headc = new List<Column>();
            for (var i = 0; i < data.Columns.Count; i++) {
                headc.Add(new Column(data.Columns[i]));
            }
            var row = new Row {
                IsHead = true,
                Columns = headc
            };
            table.Rows.Add(row);
            for (var i = 0; i < data.Rows.Count; i++) {
                var bodyc = new List<Column>();
                for (var j = 0; j < data.Rows[0].ItemArray.Length; j++) {
                    bodyc.Add(new Column(data.Rows[i][j]));
                }
                var rows = new Row {
                    Columns = bodyc
                };
                table.Rows.Add(rows);
            }
            Export.RenderToExcelForXls(table).SaveToFile("e:/test.xls");
        }
        public static void ExportForXlsx(DataTable data) {
            if (data == null || data.Rows.Count <= 0) {
                return;
            }
            var table = new Table {
                Title = string.Format("档案信息_{0}",DateTime.Now.ToString("yyyyMMddHHmmssfff")),
                BandColumns = new Dictionary<int,string[]> { { 1,new[] { "testA","testB","testC" } },{ 3,new[] { "testD","testE","testF" } } }
            };
            var headc = new List<Column>();
            for (var i = 0; i < data.Columns.Count; i++) {
                headc.Add(new Column(data.Columns[i]));
            }
            var row = new Row {
                IsHead = true,
                Columns = headc
            };
            table.Rows.Add(row);
            for (var i = 0; i < data.Rows.Count; i++) {
                var bodyc = new List<Column>();
                for (var j = 0; j < data.Rows[0].ItemArray.Length; j++) {
                    bodyc.Add(new Column(data.Rows[i][j]));
                }
                var rows = new Row {
                    Columns = bodyc
                };
                table.Rows.Add(rows);
            }
            Export.RenderToExcelForXlsx(table).SaveToFile("e:/test1.xlsx");
        }
        public static void ExportForXlsx() {
            var table = new Table {
                Title = string.Format("档案信息_{0}",DateTime.Now.ToString("yyyyMMddHHmmssfff")),
                BandColumns = new Dictionary<int,string[]> { { 1,new[] { "testA","testB","testC" } },{ 3,new[] { "testD","testE","testF" } } }
            };
            var headc = new List<Column>
            {
                new Column("test1",CellTypes.String),
                new Column("test2",CellTypes.String),
                new Column("test3",CellTypes.String),
                new Column("test4",CellTypes.String),
                new Column("test5",CellTypes.String),
                new Column("test6",CellTypes.String),
                new Column("test7",CellTypes.String),
            };
            var row = new Row {
                IsHead = true,
                Columns = headc
            };
            table.Rows.Add(row);
            for (var i = 0; i < 10000; i++) {
                var bodyc = new List<Column>
                {
                    new Column("test1",CellTypes.String),
                    new Column("test2",CellTypes.String),
                    new Column("test3",CellTypes.String),
                    new Column("test4",CellTypes.String),
                    new Column("test5",CellTypes.String),
                    new Column("test6",CellTypes.String),
                    new Column("test7",CellTypes.String),
                };
                var rows = new Row {
                    Columns = bodyc
                };
                table.Rows.Add(rows);
            }
            Export.RenderToExcelForXlsx(table).SaveToFile("e:/test1.xlsx");
        }

        public static DataTable read() {
            var data = ReadData.ExcelToDataTable("e:/tttt.xls",null);
            //if (data != null) {
            //    for (var i = 0; i < data.Rows.Count; i++) {
            //        for (var j = 0; j < data.Rows[i].ItemArray.Count(); j++) {
            //            Console.Write(data.Rows[i][j] + " ");
            //        }
            //        Console.WriteLine();
            //    }
            //}
            return data;
        }
    }
}
