using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GeracaoColetasPDF.Entities
{
    class Notes
    {
        public string Name { get; set; }
        public string Value { get; set; }

        public Notes(string name, string value)
        {
            Name = name;
            Value = value;
        }
    }
}
