namespace CaseStudy.Automation.Helper
{
    public interface IExcel : IDocuments
    {
        void AddCellValue(string cell, string value, string fileName);
    }
}