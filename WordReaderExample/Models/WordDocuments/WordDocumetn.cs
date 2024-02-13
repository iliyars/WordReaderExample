using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordReaderExample.Models.WordDocuments
{
    public class WordDocument
    {
        public virtual int ExploitationParametersStartRow { get; }
        public virtual int ExploitationParametersEndRow { get; }
        public virtual int ExploitationParametersColl { get; }

        public virtual int ParametersStartRow { get; }
        public virtual int ParametersEndRow { get; }
        public int RowCount => ParametersEndRow - ParametersStartRow + 1;
        public virtual List<int> ParametersColls { get; }
        public virtual int ItemColumnsOnPage { get; }

        public virtual string FilePath => AppDomain.CurrentDomain.BaseDirectory;
        public string NewFilePath => FilePath.Replace(".docx", "1.docx");
        protected string Name { get; set; }
        public int PageCount { get; set; }
        public WordDocument()
        {

        }
        public virtual string DocContentTitle { get; protected set; }
    }
}
