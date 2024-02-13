using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordReaderExample.Models.WordDocuments
{
    public class Form4WordDocument : WordDocument
    {
        public override int ExploitationParametersStartRow => 7;
        public override int ExploitationParametersEndRow => 18;
        public override int ExploitationParametersColl => 2;

        public override int ParametersStartRow => 2;
        public override int ParametersEndRow => 21;

        public override List<int> ParametersColls => new List<int>()
        {
            2,2,2,3,3,7,7,6,7,7,6,7,7,7,7,7,7,3,3,3
        };

        public override string FilePath => base.FilePath + "\\Resources\\Form4Template.docx";
        public override int ItemColumnsOnPage => 4;
        public Form4WordDocument()
        {
            DocContentTitle = "Карта оценки номенклатуры ЭРИ и сведений о соответствии условий их эксплуатации и показателей надежности требованиям НТД";
            Name = "Форма 4";
        }


    }
}
