
namespace NPOIHelper {
    public class Column {
        public Column(dynamic value) {
            Value = value;
            CellType = CellTypes.String;
        }
        public Column(dynamic value,CellTypes cellType) {
            Value = value;
            CellType = cellType;
        }
        public Column(dynamic value,FontColors fontColor) {
            Value = value;
            CellType = CellTypes.String;
            FontColor = (short)fontColor;
        }
        public Column(dynamic value,short fontColor) {
            Value = value;
            CellType = CellTypes.String;
            FontColor = fontColor;
        }
        public Column(dynamic value,CellTypes cellType,FontColors fontColor) {
            Value = value;
            CellType = cellType;
            FontColor = (short)fontColor;
        }
        public Column(dynamic value,CellTypes cellType,short fontColor) {
            Value = value;
            CellType = cellType;
            FontColor = fontColor;
        }
        public dynamic Value { get; set; }
        public CellTypes CellType { get; set; }
        public short? FontColor { get; set; }
    }
}
