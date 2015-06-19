using System;
using System.Data;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace NPOIHelper {
    public class ReadData {
        /// <summary>
        /// 将excel中的数据导入到DataTable中
        /// </summary>
        /// <param name="file"></param>
        /// <param name="sheetName">excel工作薄sheet的名称</param>
        /// <returns>返回的DataTable</returns>
        public static DataTable ExcelToDataTable(string file,string sheetName) {
            try {
                using (var fs = new FileStream(file,FileMode.Open,FileAccess.Read)) {
                    var workbook = new HSSFWorkbook(fs);
                    var data = new DataTable();
                    ISheet sheet;
                    if (!string.IsNullOrWhiteSpace(sheetName)) {
                        sheet = workbook.GetSheet(sheetName) ?? workbook.GetSheetAt(0);
                    } else {
                        sheet = workbook.GetSheetAt(0);
                    }
                    if (sheet != null) {
                        var firstRow = sheet.GetRow(0);
                        int cellCount = firstRow.LastCellNum; //一行最后一个cell的编号 即总的列数

                        for (int i = firstRow.FirstCellNum; i < cellCount; ++i) {
                            var cell = firstRow.GetCell(i);
                            if (cell != null) {
                                var cellValue = cell.StringCellValue;
                                if (cellValue != null) {
                                    var column = new DataColumn(cellValue);
                                    data.Columns.Add(column);
                                }
                            }
                        }
                        var startRow = sheet.FirstRowNum + 1;

                        //最后一列的标号
                        var rowCount = sheet.LastRowNum;
                        for (var i = startRow; i <= rowCount; ++i) {
                            var row = sheet.GetRow(i);
                            if (row == null) continue; //没有数据的行默认是null　　　　　　　

                            var dataRow = data.NewRow();
                            for (int j = row.FirstCellNum; j < cellCount; ++j) {
                                if (row.GetCell(j) != null) //同理，没有数据的单元格都默认是null
                                    dataRow[j] = row.GetCell(j).ToString();
                            }
                            data.Rows.Add(dataRow);
                        }
                    }
                    return data;
                }
            } catch (Exception ex) {
                Console.WriteLine("Exception: " + ex.Message);
                return null;
            }
        }
    }
}
