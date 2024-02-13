using Google.Apis.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace WordReaderExample.Models
{
    public enum FormName
    {
        [EnumMember(Value = "ФОРМА 4")]
        Form4,
        [EnumMember(Value = "ФОРМА 68")]
        Form68,
    }
}
