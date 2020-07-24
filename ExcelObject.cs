using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Data;
using System.Reflection;
using System.Drawing;

using EPPlus;
using EPPlus.Core.Extensions;
using OfficeOpenXml;

using DataTools;
using Microsoft.Extensions.Logging;
using Serilog;
using EPPlus.Core.Extensions.Style;

namespace ExcelService {
    public class ExcelObject {

        public Dictionary<string, ConvertOptions> convertOptions;
        ExcelPackage excelPackage;

        private readonly Serilog.ILogger _log = Log.ForContext(typeof(ExcelObject));

        public ExcelObject() {
            log("Excel service Started.");
        }

        ~ExcelObject() {
            if (excelPackage != null) {
                excelPackage.Dispose();
            }
            log("Excel service Disposed.");
        }

        public void initExcelPackage() {
            excelPackage = new ExcelPackage();
            convertOptions = new Dictionary<string, ConvertOptions>();
        }

        public void log(string msg) {
            _log.Information(msg);
        }

        public void AddWorkSheet<T>(IEnumerable<T> datas, string sheetName="", string optionKeyName="") {

            string workSheetName;
            ExcelWorksheet worksheet;
            int rowIdx = 1;

            if (sheetName == "") {
                workSheetName = convertOptions[optionKeyName].SheetName;
            } else {
                workSheetName = sheetName;
            }

            worksheet = excelPackage.Workbook.Worksheets.Add(workSheetName);

            foreach (object rowObj in datas) {
                Type rowType = rowObj.GetType();
                PropertyInfo[] rowInfos = rowType.GetProperties();
                int colIdx = 1;

                foreach (var rowInfo in rowInfos) {

                    //Dictionary<string, string> ColSetting;
                    //Dictionary<string, FieldSettings> FieldSettings;

                    object ColSetting = null;

                    if (optionKeyName != "") { 
                        if (convertOptions[optionKeyName].FieldSettings == null) {
                            if (convertOptions[optionKeyName].FieldsMap != null)
                                ColSetting = convertOptions[optionKeyName].FieldsMap;
                        } else {
                            ColSetting = convertOptions[optionKeyName].FieldSettings;
                        }
                    }

                    if (rowIdx == 1) {
                        //Add Header
                        if (ColSetting != null) {
                            if (ColSetting.GetType().Equals(typeof(Dictionary<string, FieldSettings>))) {
                                //log("FieldSetting detected.");
                                var fieldSettings = (Dictionary<string, FieldSettings>)ColSetting;

                                if (fieldSettings.ContainsKey(rowInfo.Name)) {
                                    worksheet.Cells[1, colIdx].Value = fieldSettings[rowInfo.Name].DisplayName;

                                    if (fieldSettings[rowInfo.Name].AutoFitColumn)
                                        worksheet.Cells[1, colIdx].AutoFitColumns();
                                    if (convertOptions[optionKeyName].BoldHeader)
                                        worksheet.Cells[1, colIdx].Style.Font.Bold = true;
                                    if (convertOptions[optionKeyName].HeaderBackgroundColor != null) {
                                        ExcelColor excelColor = (ExcelColor)convertOptions[optionKeyName].HeaderBackgroundColor;
                                        worksheet.Cells[1, colIdx].Style.SetBackgroundColor(Color.FromArgb(excelColor.A, excelColor.R, excelColor.G, excelColor.B));
                                    }
                                    if (convertOptions[optionKeyName].HeaderFontColor != null) {
                                        ExcelColor excelColor = (ExcelColor)convertOptions[optionKeyName].HeaderFontColor;
                                        worksheet.Cells[1, colIdx].Style.SetFontColor(Color.FromArgb(excelColor.A, excelColor.R, excelColor.G, excelColor.B));
                                    }

                                    colIdx++;
                                }
                            } else {
                                //log("FieldMap detected.");
                                var fieldSettings = (Dictionary<string, string>)ColSetting;
                                if (fieldSettings.ContainsKey(rowInfo.Name)) {
                                    worksheet.Cells[1, colIdx].Value = fieldSettings[rowInfo.Name];
                                    colIdx++;
                                }
                            }
                        } else {
                            worksheet.Cells[1, colIdx].Value = rowInfo.Name;
                            colIdx++;
                        }

                    } else {
                        if (rowInfo.GetValue(rowObj) != null) { 
                            if (ColSetting != null) {
                                if (ColSetting.GetType().Equals(typeof(Dictionary<string, FieldSettings>))) {
                                    
                                    var fieldSettings = (Dictionary<string, FieldSettings>)ColSetting;

                                    if (fieldSettings.ContainsKey(rowInfo.Name)) {
                                        var cellValue = rowInfo.GetValue(rowObj);
                                        //TypeCode typeCode = Type.GetTypeCode(rowInfo.PropertyType);
                                        TypeCode typeCode = Type.GetTypeCode(DataConverter.GetDataType(cellValue));

                                        switch (typeCode) {
                                            case TypeCode.DateTime:
                                                if (fieldSettings[rowInfo.Name].DisplayFormat != null) {
                                                    //cellValue = DateTime.Parse(colInfo.GetValue(rowObj)).ToString(fieldSettings[colInfo.Name].DisplayFormat);
                                                    cellValue = ((DateTime)cellValue).ToString(fieldSettings[rowInfo.Name].DisplayFormat);
                                                } else {
                                                    if (convertOptions[optionKeyName].DateFormat != null) {
                                                        cellValue = ((DateTime)cellValue).ToString(convertOptions[optionKeyName].DateFormat);
                                                    }
                                                }
                                                //log("Date format detected.");
                                                break;
                                            default:
                                                break;
                                        }

                                        worksheet.Cells[rowIdx, colIdx].Value = cellValue;

                                        if (fieldSettings[rowInfo.Name].AutoFitColumn) {
                                            //worksheet.Cells[rowIdx, colIdx].AutoFitColumns();
                                            worksheet.Column(colIdx).AutoFit();
                                        }

                                        colIdx++;
                                    }
                                } else {
                                    //log("FieldMap detected.");
                                    var fieldSettings = (Dictionary<string, string>)ColSetting;
                                    if (fieldSettings.ContainsKey(rowInfo.Name)) {
                                        worksheet.Cells[rowIdx, colIdx].Value = rowInfo.GetValue(rowObj);
                                        colIdx++;
                                    }
                                }
                            } else {
                                worksheet.Cells[rowIdx, colIdx].Value = rowInfo.GetValue(rowObj).ToString();
                                colIdx++;
                            }

                        } 
                    }

                    //colIdx++;
                }

                rowIdx++;
            }
        }
        

        public void Save(string path) {
            excelPackage.SaveAs(new FileInfo(path));
            excelPackage.Dispose();

            log($"File has saved to {path}");
        }

        //public void Save<T>(params T[] data) {
        public void Save<T>(IEnumerable<T> datas) {

            List<string> sheetNames = convertOptions.Keys.ToList<string>();

            // check if data and sheetnames are equal in count
            if (sheetNames.Count != datas.Count()) {
                Console.WriteLine("Names and Lists are not of same count");
                return;
            }

            // check if any name is empty in the sheetNames if yes give any arbitraroy name
            sheetNames.ForEach(item => {
                if (item == string.Empty) {
                    sheetNames[sheetNames.IndexOf(item)] = new Random().Next().ToString();
                }
            });

            excelPackage = new ExcelPackage(); // created a excel package

            // create sheet and save data for each object to corresponding sheet
            int index = 0;
            //sheetNames.ForEach(item => {

                /* var cols = data[index][0].GetType().GetProperties().Select(item1 => item1.Name);

                int col = 1;
                cols.ToList().ForEach(column => {
                    worksheet.Cells[1, col].Value = column;
                    col++;
                }); */

                // add data
                /* int row = 2;
                data[index].ForEach(dataObject => {
                    col = 1;
                    cols.ToList().ForEach(column => {
                        worksheet.Cells[row, col].Value = dataObject.GetType().GetProperty(column).GetValue(dataObject);
                        col++;
                    });
                    row++;
                }); */

            //});

            //excelPackage.Dispose();
        }
        public static List<Object> ListToObject<T>(List<T> input) {

            List<Object> _listCache = new List<object>();

            foreach (object item in input) {
                _listCache.Add(item);
            }

            return _listCache;
        }
    }
}
