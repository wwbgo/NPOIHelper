using System;
using System.Data;
using System.IO;
using System.Linq;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using OfficeOpenXml;

namespace NPOIHelper {
    public static class Export {
        public static MemoryStream RenderToExcelForXls(Table table) {
            using (var ms = new MemoryStream()) {
                var workbook = new HSSFWorkbook();
                var sheet = string.IsNullOrWhiteSpace(table.Title) ? workbook.CreateSheet() : workbook.CreateSheet(table.Title);
                var rowIndex = 0;
                var head = table.Rows.FirstOrDefault(r => r.IsHead);
                var body = table.Rows.Where(r => !r.IsHead).ToList();
                if (!body.Any()) {
                    throw new Exception("没有需要导出的数据");
                }
                if (head != null) {
                    var headerRow = sheet.CreateRow(rowIndex++);
                    for (var i = 0; i < head.Columns.Count; i++) {
                        headerRow.CreateCell(i).SetCellValue(head.Columns[i].Value);
                        CellSetter.SetCellStyle(workbook,headerRow.Cells[i],head.Columns[i].FontColor);
                    }
                    //第一个参数表示要冻结的列数；
                    //第二个参数表示要冻结的行数；
                    //第三个参数表示右边区域可见的首列序号；
                    //第四个参数表示下边区域可见的首行序号；
                    sheet.CreateFreezePane(0,1,0,1);

                    foreach (var row in body) {
                        var dataRow = sheet.CreateRow(rowIndex++);
                        for (var i = 0; i < row.Columns.Count; i++) {
                            dataRow.CreateCell(i,(CellType)row.Columns[i].CellType).SetCellValue(row.Columns[i].Value);
                        }
                    }
                } else {
                    foreach (var row in body) {
                        var dataRow = sheet.CreateRow(rowIndex++);
                        for (var i = 0; i < row.Columns.Count; i++) {
                            dataRow.CreateCell(i,(CellType)row.Columns[i].CellType).SetCellValue(row.Columns[i].Value);
                        }
                    }
                }
                if (table.BandColumns.Any()) {
                    foreach (var bandColumn in table.BandColumns) {
                        CellSetter.SetCellDropdownlist(sheet,bandColumn.Key,bandColumn.Value);
                    }
                }
                //            IList<CourseCodeInfo> list = StudentBus.GetSubjectInterface().GetList(0,"","Name");

                //            var CourseSheetName = "Course";
                //            var RangeName = "dicRange";

                //            sheet.Workbook.CreateRow(0).CreateCell(0).SetCellValue("课程列表（用于生成课程下拉框，请勿修改）");

                //            for (var i = 1; i < list.Count; i++) {
                //                sheet.CreateRow(i).CreateCell(0).SetCellValue(list[i - 1].Name);
                //            }
                //            IName range = workbook.CreateName();
                //range.RefersToFormula = string.Format("{0}!$A$2:$A${1}",CourseSheetName,list.Count.ToString());<br>range.NameName = RangeName;
                //            CellRangeAddressList regions = new CellRangeAddressList(1,65535,4,4);
                //            DVConstraint constraint = DVConstraint.CreateFormulaListConstraint(RangeName);
                //            HSSFDataValidation dataValidate = new HSSFDataValidation(regions,constraint);
                //            sheet.AddValidationData(dataValidate);


                workbook.Write(ms);
                //ms.Flush();
                ms.Position = 0;
                return ms;
            }
        }

        public static MemoryStream RenderToExcelForXls(DataTable table,string sheetName) {
            using (var ms = new MemoryStream()) {
                var workbook = new HSSFWorkbook();
                var sheet = string.IsNullOrWhiteSpace(sheetName) ? workbook.CreateSheet() : workbook.CreateSheet(sheetName);

                if (table == null || table.Rows.Count <= 0) {
                    throw new Exception("没有需要导出的数据");
                }
                //表头  
                var row = sheet.CreateRow(0);
                for (var i = 0; i < table.Columns.Count; i++) {
                    var cell = row.CreateCell(i);
                    cell.SetCellValue(table.Columns[i].ColumnName);
                }
                //数据  
                for (var i = 0; i < table.Rows.Count; i++) {
                    var row1 = sheet.CreateRow(i + 1);
                    for (var j = 0; j < table.Columns.Count; j++) {
                        var cell = row1.CreateCell(j);
                        cell.SetCellValue(table.Rows[i][j].ToString());
                    }
                }
                workbook.Write(ms);
                //ms.Flush();
                ms.Position = 0;
                return ms;
            }
        }

        public static Stream RenderToExcelForXlsx(Table table) {
            var excel = new ExcelPackage();
            var workbook = excel.Workbook.Worksheets;
            var sheet = string.IsNullOrWhiteSpace(table.Title) ? workbook.Add("sheet1") : workbook.Add(table.Title);

            var head = table.Rows.FirstOrDefault(r => r.IsHead);
            var body = table.Rows.Where(r => !r.IsHead).ToList();
            if (!body.Any()) {
                throw new Exception("没有需要导出的数据");
            }
            if (head != null) {
                for (var i = 0; i < head.Columns.Count; i++) {
                    sheet.Cells[1,i + 1].Value = head.Columns[i].Value;
                }

                var r = 2;
                foreach (var row in body) {
                    for (var i = 0; i < row.Columns.Count; i++) {
                        sheet.Cells[r,i + 1].Value = row.Columns[i].Value;
                    }
                    r++;
                }
            } else {
                var r = 2;
                foreach (var row in body) {
                    for (var i = 0; i < row.Columns.Count; i++) {
                        sheet.Cells[r,i + 1].Value = row.Columns[i].Value;
                    }
                    r++;
                }
            }

            excel.Save();
            excel.Stream.Position = 0;
            return excel.Stream;
        }

        public static Stream RenderToExcelForXlsx(DataTable table,string sheetName) {
            var excel = new ExcelPackage();
            var workbook = excel.Workbook.Worksheets;
            var sheet = string.IsNullOrWhiteSpace(sheetName) ? workbook.Add("sheet1") : workbook.Add(sheetName);

            if (table == null || table.Rows.Count <= 0) {
                throw new Exception("没有需要导出的数据");
            }
            //表头
            for (var i = 0; i < table.Columns.Count; i++) {
                sheet.Cells[1,i + 1].Value = table.Columns[i].ColumnName;
            }
            //数据
            for (var i = 0; i < table.Rows.Count; i++) {
                for (var j = 0; j < table.Columns.Count; j++) {
                    sheet.Cells[i + 2,j + 1].Value = table.Rows[i][j];
                }
            }
            excel.Save();
            excel.Stream.Position = 0;
            return excel.Stream;
        }

        public static void SaveToFile(this MemoryStream ms,string fileName) {
            using (var fs = new FileStream(fileName,FileMode.Create,FileAccess.ReadWrite)) {
                var data = ms.GetBuffer();
                fs.Write(data,0,data.Length);
                fs.Flush();
            }
        }
        public static void SaveToFile(this Stream ms,string fileName) {
            using (var fs = new FileStream(fileName,FileMode.Create,FileAccess.ReadWrite)) {
                ms.CopyTo(fs);
                fs.Flush();
            }
        }
    }
}
