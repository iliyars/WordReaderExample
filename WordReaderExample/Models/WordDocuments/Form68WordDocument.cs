using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordReaderExample.Models.WordDocuments
{
    public class Form68WordDocument : WordDocument
    {
        public override int ParametersStartRow => 2;
        public override int ParametersEndRow => 20;
        public override List<int> ParametersColls => new List<int>()
        {
            2,2,5,5,5,5,5,5,6,6,5,4,5,5,4,4,4,3
        };

        public override string FilePath => base.FilePath + "\\Resources\\Form68Template.docx";
        public override int ItemColumnsOnPage => 3;

        public Form68WordDocument()
        {
            DocContentTitle = "Карта рабочих режимов резисторов, резисторных сборок, терморезисторов, поглотителей и потенциометров";
            Name = "Форма 68";
        }
    }
}
