using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MedabilNavisworks
{
    class Sequencia
    {
        public string TipoObjeto { get; set; }
        public string Nome { get; set; }
        public double PesoLiquido { get; set; } = 0;
        public List<Membro> Membros { get; set; } = new List<Membro>();
        public Dictionary<string, int> TypesCounter { get; set; } = new Dictionary<string, int>();
        public Dictionary<string, double> TypesNetWeight { get; set; } = new Dictionary<string, double>();

        

        public Sequencia()
        {

        }

        public Sequencia(string tipoObjeto, string name)
        {
            TipoObjeto = tipoObjeto;
            Nome = name;
            
        }
    }
}
