using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CaseStudy.Automation.Helper
{
    internal class Seeds : Documents, ISeeds
    {
        private string filePath;
        public Seeds(string filePath) : base(filePath)
        {
            this.filePath = filePath;
        }

        public void AddCellValues(List<int> cellNumbers)
        {
            foreach( var number in cellNumbers)
            {
                int x = 3;
                base.AddCellValue(String.Format("{0}{1}", "E", x), number.ToString(), filePath);
                x++;
            }
        }
    }
}
