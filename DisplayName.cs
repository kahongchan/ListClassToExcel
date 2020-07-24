using System;

namespace ExcelService {
    public class DisplayName: Attribute {
        string _displayName;
        public string displayName {
            get {
                return _displayName;
            }
            set {
                _displayName = value;
            }
        }

        public DisplayName(string title) {
            this._displayName = title;
        }
    }
}
