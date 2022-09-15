using CaseStudy.Automation.Helper;
using OfficeOpenXml;

namespace CaseStudy.Automation
{
    public class Tests
    {
        string fileKeep;
        string newFile;
        string directory;
        ExcelPackage package;

        [OneTimeSetUp]
        public void Setup()
        {
            fileKeep = "CaseStudy.xlsx";
            newFile = "CaseStudySample.xlsx"; // File to be Created
            directory = "../../../src";
            package = new ExcelPackage();
            //List<int> numbers = new List<int>() { 10,2,10,3,9};
            //ISeeds loadValues = new Seeds(directory);
            /*
             * Initialization of my Document(Excelfile)
             */
            IExcel excel = new Documents(directory);
            /*
             * Creating a new workbook
             */
            excel.CreateFile(newFile);
            /*
             * Navigating to a Cell, and adding value
             */
            excel.AddCellValue("L2", "Hi", fileKeep);
            excel.AddCellValue("M4", "I am Josh", fileKeep);
            excel.AddCellValue("A7", "Joshua was Here!!", fileKeep); ;
            /*
             * 
             * 
             */
            //loadValues.AddCellValues(numbers);
        }

        /* Launching an xlsx file
         * Determines if the file Exists or not. 
         * Pass if the file exists
         */
        [Test]
        [TestCase("../../../src", "CaseStudy.xlsx")] //will Pass
        [TestCase("../../../src", "CaseStudySample.xlsx")] //will Pass
        public void CheckFile_IfExists_Valid(string dir, string fileName)
        {
            string filePath = String.Format("{0}/{1}", directory, fileName);
            FileInfo fileInfo = new FileInfo(filePath);

            if (fileInfo.Exists)
            {
                Assert.Pass();
            }
            else
            {
                Assert.Fail();
            }
        }
        /*
         * Navigate to a new cell
         * Check if that cell has a value or not
         * Passes if there are value in the cell
         */
        [Test]
        [TestCase("L2", "CaseStudy.xlsx")] 
        [TestCase("M4", "CaseStudy.xlsx")] 
        [TestCase("A7", "CaseStudy.xlsx")]
        public void SearchWorkbook_CheckCellIfHasAValue_Valid(string searchCell, string fileName)
        {
            string filePath = String.Format("{0}/{1}", directory, fileName);
            FileInfo fileInfo = new FileInfo(filePath);
            package = new ExcelPackage(fileInfo);

            using (package)
            {
                var ws = package.Workbook.Worksheets[0]; //by default, let us use the first worksheet
                var cell = ws.Cells[searchCell].Value;
                if (cell != null)
                {
                    Assert.IsNotEmpty(cell.ToString());
                }
                else
                {
                    Assert.Fail();
                }
            }
        }
        /* Check if you can navigate to an expected Cell.
         * Manually added a value in the excel file >> `Hello Outsourced! This text is created within the file.` @C25
         * We can also use the values in Line 32:34 as expected --
         * the text above is now an 'expected' value for the cell C25
         * Pass if the given text is captured.
         */
        [Test]
        [TestCase("C25", "Hello Outsourced! This text is created within the file.", "CaseStudy.xlsx")] 
        [TestCase("L2", "Hi", "CaseStudy.xlsx")] 
        public void SearchWorkbook_CheckExpectedCell_Valid(string searchCell,string expectedValue, string fileName)
        {
            string filePath = String.Format("{0}/{1}", directory, fileName);
            FileInfo fileInfo = new FileInfo(filePath);
            package = new ExcelPackage(fileInfo);

            using (package)
            {
                var ws = package.Workbook.Worksheets[0]; //by default, let us use the first worksheet
                var cell = ws.Cells[searchCell].Value;
                if(cell != null)
                {
                    Assert.AreEqual(cell.ToString(), expectedValue);
                }
                else
                {
                    Assert.Fail();
                }
                
            }

        }
        [Test]
        [TestCase("C25", "Hi", "CaseStudy.xlsx")]
        public void SearchWorkbook_CheckExpectedCell_Invalid(string searchCell, string expectedValue, string fileName)
        {
            string filePath = String.Format("{0}/{1}", directory, fileName);
            FileInfo fileInfo = new FileInfo(filePath);
            package = new ExcelPackage(fileInfo);

            using (package)
            {
                var ws = package.Workbook.Worksheets[0]; //by default, let us use the first worksheet
                var cell = ws.Cells[searchCell].Value;
                if (cell != null)
                {
                    Assert.AreNotEqual(cell.ToString(), expectedValue);
                }
                else
                {
                    Assert.Fail();
                }

            }

        }
        [Test]
        public void function()
        {

        }
        [OneTimeTearDown]
        public void Cleanup()
        {
            DirectoryInfo dirCleaner = new DirectoryInfo(directory);
            foreach (FileInfo file in dirCleaner.GetFiles().Where(x=>x.Name != fileKeep).ToList())
            {
                file.Delete();
            }

        }

    }
}