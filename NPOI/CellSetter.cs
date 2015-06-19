using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using OfficeOpenXml;

namespace NPOIHelper {
    class CellSetter {
        /// <summary>
        /// 设置单元格为下拉框并限制输入值
        /// 第一个参数表示第一行；
        /// 第二个参数表示最后一行；
        /// 第三个参数表示第一列；
        /// 第四个参数表示最后一列；
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="column"></param>
        /// <param name="datalist"></param>
        public static void SetCellDropdownlist(ISheet sheet,int column,string[] datalist) {
            //设置生成下拉框的行和列
            var cellRegions = new CellRangeAddressList(1,65535,column,column);
            //设置 下拉框内容
            var constraint = DVConstraint.CreateExplicitListConstraint(datalist);
            //绑定下拉框和作用区域，并设置错误提示信息
            var dataValidate = new HSSFDataValidation(cellRegions,constraint);
            dataValidate.CreateErrorBox("输入不合法","请输入下拉列表中的值。");
            dataValidate.ShowPromptBox = true;
            sheet.AddValidationData(dataValidate);
        }
        public static void SetCellDropdownlistForXlsx(ExcelWorksheet sheet,int column,string[] datalist) {
            //var val = sheet.Cells[2,column + 1,10000,column + 1].DataValidation.AddListDataValidation();//设置下拉框显示的数据区域
            var val = sheet.DataValidations.AddListValidation(sheet.Cells[2,column + 1,10000,column + 1].Address);//设置下拉框显示的数据区域
            foreach (var item in datalist) {
                val.Formula.Values.Add(item);
            }
            val.Error = "请选择下拉框中的数据";
            val.ShowErrorMessage = true;
            //val.Prompt = "下拉选择";//下拉提示
            //val.ShowInputMessage = true;//显示提示内容
        }

        /// <summary>
        /// 设置单元格样式
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="cell"></param>
        /// <param name="fontColor"></param>
        public static void SetCellStyle(IWorkbook workbook,ICell cell,short? fontColor = null) {
            var fCellStyle = (HSSFCellStyle)workbook.CreateCellStyle();
            var ffont = (HSSFFont)workbook.CreateFont();
            ffont.FontHeight = 15 * 15;
            ffont.FontName = "宋体";
            ffont.Color = fontColor ?? HSSFColor.Black.Index;
            fCellStyle.SetFont(ffont);

            fCellStyle.VerticalAlignment = VerticalAlignment.Center;//垂直对齐
            fCellStyle.Alignment = HorizontalAlignment.Center;//水平对齐
            cell.CellStyle = fCellStyle;
        }
    }
}
