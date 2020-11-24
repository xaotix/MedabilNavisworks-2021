using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MedabilNavisworks
{
    class Sequence
    {
        public string TipoObjeto { get; set; }
        public string Name { get; set; }
        public double NetWeight { get; set; } = 0;
        public List<Member> Members { get; set; } = new List<Member>();
        public Dictionary<string, int> TypesCounter { get; set; } = new Dictionary<string, int>();
        public Dictionary<string, double> TypesNetWeight { get; set; } = new Dictionary<string, double>();

        

        public Sequence()
        {

        }

        public Sequence(string tipoObjeto, string name)
        {
            TipoObjeto = tipoObjeto;
            Name = name;
            
        }
    }
}
