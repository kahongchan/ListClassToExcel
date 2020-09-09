using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelService {
    public class ConvertOptions {

        public string SheetName { get; set; }
        public string SavePath { get; set; }

        public string DateFormat { get; set; }
        public bool BoldHeader { get; set; }
        public ExcelColor? HeaderFontColor { get; set; }
        public ExcelColor? HeaderBackgroundColor { get; set; }

        [Obsolete("FieldsMap option are obsoleted. Please use 'FieldSettings' instead. ")]
        public Dictionary<string, string> FieldsMap { get; set; }
        public Dictionary<string, FieldSettings> FieldSettings { get; set; }
    }

    public class FieldSettings : Attribute {
        public int DisplayIndex { get; set; } = 0;
        public bool AutoFitColumn { get; set; }
        public string DisplayName { get; set; }
        public string DisplayFormat { get; set; }
    }

    public class ExcelColor {
        public int R { get; set; } = 255;
        public int G { get; set; } = 255;
        public int B { get; set; } = 255;
        public int A { get; set; } = 1;

        public ExcelColor() {

        }

        public ExcelColor(int red, int green, int blue, int alpha=1) {
            R = red; G = green; B = blue; A = alpha;
        }
    } 
}
