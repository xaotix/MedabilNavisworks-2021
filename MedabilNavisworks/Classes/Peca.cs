using Autodesk.Navisworks.Api;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MedabilNavisworks
{
    class Peca
    {
        public List<DataProperty> GetPropsMarca()
        {
            /*aqui tem uma baderna. cada tipo de software bota numa propriedade diferente*/
            List<DataProperty> retorno = new List<DataProperty>();
            retorno.AddRange(this.GetPropriedades().FindAll(x => x.DisplayName.ToUpper().Replace(" ", "_") == "ASSEMBLY_MARK"));
            retorno.AddRange(this.GetPropriedades().FindAll(x => x.DisplayName.ToUpper().Contains("PIECEMARK")));
            retorno.AddRange(this.GetPropriedades().FindAll(x => x.DisplayName.ToUpper().Contains("MARK")));
            retorno.AddRange(this.GetPropriedades().FindAll(x => x.DisplayName.ToUpper() == "TAG"));
            retorno.AddRange(this.GetPropriedades().FindAll(x => x.DisplayName.ToUpper().Contains("NAME")));
            retorno.AddRange(this.GetPropriedades().FindAll(x => x.DisplayName.ToUpper().Contains("REFERENCE")));
            retorno.AddRange(this.GetPropriedades().FindAll(x => x.DisplayName.ToUpper().Contains("DESCRIPTION")));
            retorno.AddRange(this.GetPropriedades().FindAll(x => x.DisplayName.ToUpper().Contains("BATID")));
            return retorno;
        }
        public List<DataProperty> GetPropsEtapa()
        {
            return this.GetPropriedades().FindAll(x =>
                            x.DisplayName.ToUpper() == "SEQUENCE" |
                            x.DisplayName.ToUpper() == "PHASE"
                           );
        }
        public List<DataProperty> GetPropsNumero()
        {
            return this.GetPropriedades().FindAll(x =>
                            x.DisplayName.ToUpper() == "MEMBER_NUMBER" |
                            x.DisplayName.ToUpper() == "GUID"
                            );
        }

        public List<DataProperty> GetPropsPeso()
        {
            return this.GetPropriedades().FindAll(x =>
                       x.DisplayName.ToUpper().Contains("NET_WEIGHT") |
                       x.DisplayName.ToUpper() == "WEIGHT" |
                       x.DisplayName.ToUpper().Replace(" ", "_").Contains("UNIT_WEIGHT")
                       );
        }

        public List<DataProperty> GetPropsTipo()
        {
            return this.GetPropriedades().FindAll(x =>
                          x.DisplayName.ToUpper().Contains("TYPE")|
                          x.DisplayName.ToUpper().Contains("MEMBER_TYPE")
                          );
        }

        public List<DataProperty> GetPropsDescricao()
        {
            return this.GetPropriedades().FindAll(x =>
                            x.DisplayName.ToUpper().Contains("DESCRIÇÃO") |
                            x.DisplayName.ToUpper().Contains("DESCRITPION")
                               );
        }

        private List<DataProperty> _propriedades { get;set; }
        public List<DataProperty> GetPropriedades()
        {
            if(_propriedades ==null)
            {
                _propriedades = Funcoes.GetPropriedades(this.ModelItem);
            }
            return _propriedades;
        }
        public override string ToString()
        {
            return this.Marca;
        }
        public ModelItem ModelItem { get; private set; }
        public string TipoObjeto { get; private set; } = "member";



        public string GetNumero()
        {
            var nums = this.GetPropsNumero();
            if(nums.Count>0)
            {
                return Funcoes.Getvalor(nums[0]);
            }
            return "";
        }


        public string GetEtapa()
        {
            var nums = this.GetPropsEtapa();
            if (nums.Count > 0)
            {
                return Funcoes.Getvalor(nums[0]);
            }
            return "";
        }

        public double GetPesoLiquido()
        {
            var nums = this.GetPropsPeso();
            if (nums.Count > 0)
            {
                return Conexoes.Utilz.Double(Funcoes.Getvalor(nums[0]));
            }
            return 0;
        }

        public string Marca { get; private set; } = "";
        public string Tipo { get; private set; } = "";
        private string GetMarca()
        {
           
            if(this.ModelItem!=null)
            {
                var nums = this.GetPropsMarca().FindAll(x=>Funcoes.Getvalor(x).ToUpper().Replace(" ","").Replace(".","").Replace("-","")!="");

                string nome_fim = "";

                if (nums.Count > 0)
                {
                    nome_fim = Funcoes.Getvalor(nums[0]);
                }

               if(nome_fim.Replace(" ", "").Replace(".", "").Replace("-", "") == "")
                {
                    var nome = this.ModelItem.DisplayName.Replace(" ", "").Replace(".", "").Replace("-", "");
                    if (nome != "")
                    {
                        nome_fim = this.ModelItem.DisplayName;
                    }
                    else
                    {

                    }
                }

                //15/04/2020 - para ler os inputs de TecnoMetal
                if (nome_fim.ToUpper().Contains("MARK") && nome_fim.ToUpper().Contains("POS"))
                {
                    var m = nome_fim.Split(' ').ToList();
                    nome_fim = m[0].ToUpper().Replace(" ", "").Replace("MARK", "").Replace(":", "");
                }

                if (nome_fim == Tipo)
                {
                    nome_fim = this.ModelItem.DisplayName;
                }
                return nome_fim;
            }
            return "";
        }


        public string GetTipo()
        {
            var nums = this.GetPropsTipo();
            if (nums.Count > 0)
            {
                return Funcoes.Getvalor(nums[0]);
            }
            return "";
        }

        public Peca(ModelItem PC, string number, double netWeight, string piecemark, string type)
        {
            this.ModelItem = PC;
        }
        public Peca(ModelItem PC)
        {
            this.ModelItem = PC;
            this.Marca = GetMarca();
            this.Tipo = GetTipo();
        }
        
    }
}
