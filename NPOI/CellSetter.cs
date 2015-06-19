using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

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
        /// <param name="isXlsx"></param>
        public static void SetCellDropdownlist(ISheet sheet,int column,string[] datalist,bool isXlsx = false) {
            if (isXlsx) {
                //设置生成下拉框的行和列
                var cellRegions = new CellRangeAddressList(1,10000,column,column);
                //设置 下拉框内容
                var constraint = new XSSFDataValidationConstraint(datalist);
                var dataValidation = new CT_DataValidation { showDropDown = true,@operator = ST_DataValidationOperator.between };
                //绑定下拉框和作用区域，并设置错误提示信息
                var dataValidate = new XSSFDataValidation(constraint,cellRegions,dataValidation);
                dataValidate.CreateErrorBox("输入不合法","请输入下拉列表中的值。");
                dataValidate.ShowPromptBox = true;
                sheet.AddValidationData(dataValidate);
            } else {
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
        }

        /// <summary>
        /// 设置单元格样式
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="cell"></param>
        /// <param name="fontColor"></param>
        /// <param name="isXlsx"></param>
        public static void SetCellStyle(IWorkbook workbook,ICell cell,short? fontColor = null,bool isXlsx = false) {
            if (isXlsx) {
                var fCellStyle = (XSSFCellStyle)workbook.CreateCellStyle();
                var ffont = (XSSFFont)workbook.CreateFont();
                ffont.FontHeight = 4 * 4;
                ffont.FontName = "宋体";
                ffont.Color = fontColor ?? HSSFColor.Black.Index;
                fCellStyle.SetFont(ffont);

                fCellStyle.VerticalAlignment = VerticalAlignment.Center;//垂直对齐
                fCellStyle.Alignment = HorizontalAlignment.Center;//水平对齐
                cell.CellStyle = fCellStyle;
            } else {
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
}
