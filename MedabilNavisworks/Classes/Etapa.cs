using Autodesk.Navisworks.Api;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MedabilNavisworks
{
    class Etapa
    {
        public string TipoObjeto { get; set; }
        public string Nome { get; set; }
        public double PesoLiquido { get; set; } = 0;
        public List<Peca> Pecas { get; set; } = new List<Peca>();
        public Dictionary<string, int> TypesCounter { get; set; } = new Dictionary<string, int>();
        public Dictionary<string, double> TypesNetWeight { get; set; } = new Dictionary<string, double>();

        public ModelItem modelItem { get; set; }
        

        public Etapa()
        {

        }

        public Etapa(string tipoObjeto, string name, ModelItem item)
        {
            TipoObjeto = tipoObjeto;
            Nome = name;
            this.modelItem = item;
        }
    }
}
