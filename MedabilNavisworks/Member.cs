using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MedabilNavisworks
{
    class Member
    {
        public string TipoObjeto { get; set; }
        public string Number { get; set; }
        public double NetWeight { get; set; }
        public string Piecemark { get; set; }
        public string Type { get; set; }

        
        public Member(string tipoObjeto, string number, double netWeight, string piecemark, string type)
        {
            TipoObjeto = tipoObjeto;
            Number = number;
            NetWeight = netWeight;
            Piecemark = piecemark;
            Type = type;
        }
    }
}
