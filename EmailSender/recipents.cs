using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailSender
{
    internal class Recipents
    {
        public String Name;
        public String Email;
        public String Attachment;

        public Recipents(string name, string email, string attachment)
        {
            Name = name;
            this.Email = email;
            this.Attachment = attachment;
        }

        public static List<Recipents> Read()
        {


            // Path to your Excel file
            string filePath = "mails.xlsx";
            List<Recipents> tatok = new List<Recipents>();
            // Ensure the EPPlus license context is set
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Read the Excel file
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                // Get the first worksheet in the workbook
                var worksheet = package.Workbook.Worksheets[0];
                int rows = worksheet.Dimension.Rows;
                int columns = worksheet.Dimension.Columns;




                // Loop through each row and column to read data
                String[] mail = new string[4];
                for (int row = 2; row < rows; row++)
                {
                    for (int col = 1; col <= columns; col++)
                    {
                        // Get the value of the cell
                        var cellValue = worksheet.Cells[row, col].Text;
                        mail[col - 1] = cellValue;

                        // Console.Write(cellValue + "\t");
                    }
                    tatok.Add(new Recipents(mail[0], mail[1], "attachments/" +mail[2] + mail[3]));

                }
            }
            return tatok;
        }
    }
}
