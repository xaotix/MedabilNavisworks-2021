using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MedabilNavisworks
{
    class Peca
    {
        public string TipoObjeto { get; set; }
        public string Numero { get; set; }
        public double PesoLiquido { get; set; }
        public string Marca { get; set; }
        public string Tipo { get; set; }

        
        public Peca(string tipoObjeto, string number, double netWeight, string piecemark, string type)
        {
            TipoObjeto = tipoObjeto;
            Numero = number;
            PesoLiquido = netWeight;
            Marca = piecemark;
            Tipo = type;
        }
    }
}
