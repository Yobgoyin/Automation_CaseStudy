using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CaseStudy.Automation.Helper
{
    public class Documents : IExcel
    {
        private string filePath;
        public string fileName { get; set; }

        public Documents(string filePath)
        {
            this.filePath = filePath;
        }

        public void CreateFile(string fileName)
        {
            using (var package = new ExcelPackage())
            {
                try
                {
                    package.Workbook.Worksheets.Add("Sheet 1");
                    package.SaveAs(new FileInfo(String.Format("{0}/{1}", filePath, fileName)));
                }
                catch (Exception ex)
                {
                    throw ex;
                }

            }
        }

        public void AddCellValue( string cell, string value, string fileName)
        {
            FileInfo fileInfo = new FileInfo(String.Format("{0}/{1}", filePath, fileName));

            using (var package = new ExcelPackage(fileInfo))
            {
                var ws = package.Workbook.Worksheets[0];
                var c = ws.Cells[cell];
                c.Value = value;
            }
        }
    }
}
