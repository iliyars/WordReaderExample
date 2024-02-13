using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordReaderExample.Models
{
    public class DomainObject : ICloneable
    {
        public int Id { get; set; }

        [NotMapped]
        public string? PositionNames { get; set; }
        public string Name { get; set; }

        public string Type { get; set; }

        public string? Family { get; set; }

        public string? Note { get; set; }

        public string FormName { get; set; }

        [NotMapped]
        public int? Quantity { get; set; }

        public virtual List<string> GetParamsForWord()
        {
            throw new NotImplementedException();
        }
        public DomainObject()
        {

        }

        public DomainObject(string family, string name, int? quantity)
        {
            Family = family;
            Name = name;
            Quantity = quantity;
        }

        public object Clone()
        {
            return MemberwiseClone();
        }
    }
}
