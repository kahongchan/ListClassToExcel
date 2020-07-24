using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Data;


namespace ExcelService {
    public class SaveToExcel<T> where T : List<Object> {

        List<string> sheetNames;
        string path;

        public SaveToExcel(List<string> sheetNames, string path) {
            this.sheetNames = sheetNames;
            this.path = path;
        }

        public void Save(params T[] data) {
            // check if data and sheetnames are equal in count
            if (sheetNames.Count != data.Length) {
                Console.WriteLine("Names and Lists are not of same count");
                return;
            }

            // check if any name is empty in the sheetNames if yes give any arbitraroy name
            sheetNames.ForEach(item => {
                if (item == string.Empty) {
                    sheetNames[sheetNames.IndexOf(item)] = new Random().Next().ToString();
                }
            });

            ExcelPackage excelPackage = new ExcelPackage(); // created a excel package

            // create sheet and save data for each object to corresponding sheet
            int index = 0;
            sheetNames.ForEach(item => {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(item);
                var cols = data[index][0].GetType().GetProperties().Select(item1 => item1.Name);

                int col = 1;
                cols.ToList().ForEach(column => {
                    worksheet.Cells[1, col].Value = column;
                    col++;
                });

                // add data
                int row = 2;
                data[index].ForEach(dataObject => {
                    col = 1;
                    cols.ToList().ForEach(column => {
                        worksheet.Cells[row, col].Value = dataObject.GetType().GetProperty(column).GetValue(dataObject);
                        col++;
                    });
                    row++;
                });

            });
            excelPackage.SaveAs(new FileInfo(path));
            excelPackage.Dispose();
        }
    }
}
