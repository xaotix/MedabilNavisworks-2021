using Autodesk.Navisworks.Api;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MedabilNavisworks
{
    public class Peca
    {
        public List<DB.Celula> GetPropsMarca()
        {
            /*aqui tem uma baderna. cada tipo de software bota numa propriedade diferente*/
            List<DB.Celula> retorno = new List<DB.Celula>();
            retorno.AddRange(this.GetPropriedades().FindAll(x => x.Coluna.ToUpper().Replace(" ", "_") == "ASSEMBLY_MARK"));
            retorno.AddRange(this.GetPropriedades().FindAll(x => x.Coluna.ToUpper().Contains("PIECEMARK")));
            retorno.AddRange(this.GetPropriedades().FindAll(x => x.Coluna.ToUpper().Contains("MARK")));
            retorno.AddRange(this.GetPropriedades().FindAll(x => x.Coluna.ToUpper() == "TAG"));
            retorno.AddRange(this.GetPropriedades().FindAll(x => x.Coluna.ToUpper().Contains("NAME")));
            retorno.AddRange(this.GetPropriedades().FindAll(x => x.Coluna.ToUpper().Contains("REFERENCE")));
            retorno.AddRange(this.GetPropriedades().FindAll(x => x.Coluna.ToUpper().Contains("DESCRIPTION")));
            retorno.AddRange(this.GetPropriedades().FindAll(x => x.Coluna.ToUpper().Contains("BATID")));
            return retorno;
        }
        public List<DB.Celula> GetPropsEtapa()
        {
            return this.GetPropriedades().FindAll(x =>
                            x.Coluna.ToUpper() == "SEQUENCE" |
                            x.Coluna.ToUpper() == "PHASE"
                           );
        }
        public List<DB.Celula> GetPropsNumero()
        {
            return this.GetPropriedades().FindAll(x =>
                            x.Coluna.ToUpper() == "MEMBER_NUMBER" |
                            x.Coluna.ToUpper() == "GUID"
                            );
        }

        public List<DB.Celula> GetPropsPeso()
        {
            return this.GetPropriedades().FindAll(x =>
                       x.Coluna.ToUpper().Contains("PESO_TOTAL") |
                       x.Coluna.ToUpper().Contains("PESO") |
                       x.Coluna.ToUpper().Contains("NET_WEIGHT") |
                       x.Coluna.ToUpper() == "WEIGHT" |
                       x.Coluna.ToUpper().Replace(" ", "_").Contains("UNIT_WEIGHT")
                       );
        }

        public List<DB.Celula> GetPropsComp()
        {
            return this.GetPropriedades().FindAll(x =>
                       x.Coluna.ToUpper().Contains("COMPRIMENTO") |
                       x.Coluna.ToUpper().Replace(" ", "_").Contains("LENGHT")
                       );
        }
        public List<DB.Celula> GetPropsLarg()
        {
            return this.GetPropriedades().FindAll(x =>
                       x.Coluna.ToUpper().Contains("LARGURA") |
                       x.Coluna.ToUpper().Replace(" ", "_").Contains("WIDTH")
                       );
        }
        public List<DB.Celula> GetPropsEspessura()
        {
            return this.GetPropriedades().FindAll(x =>
                       x.Coluna.ToUpper().Contains("ESPESSURA") |
                       x.Coluna.ToUpper().Replace(" ", "_").Contains("THICK")
                       );
        }
        public List<DB.Celula> GetPropsTipo()
        {
            return this.GetPropriedades().FindAll(x =>
                          (x.Coluna.ToUpper().Contains("TYPE") && x.Tabela.ToUpper() == "ITEM"));
        }

        public List<DB.Celula> GetPropsDescricao()
        {
            return this.GetPropriedades().FindAll(x =>
                            x.Coluna.ToUpper().Contains("DESCRIÇÃO") |
                            x.Coluna.ToUpper().Contains("NAME") |
                            x.Coluna.ToUpper().Contains("DESCRITPION")
                               );
        }


        public List<string> GetTabs()
        {
            return this.GetPropriedades().Select(x => x.Tabela).Distinct().ToList().OrderBy(x => x).ToList();
        }


        private List<DB.Celula> _propriedades { get;set; }
        public List<DB.Celula> GetPropriedades(List<string> tab = null)
        {
            if(_propriedades ==null)
            {
                _propriedades = Funcoes.GetPropriedadesTab(this.ModelItem);
            }
            if(tab!=null)
            {
                if(tab.Count>0)
                {
                return _propriedades.FindAll(x => tab.Find(y=>y.ToUpper() == x.Tabela.ToUpper()) != null);
                }
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
                return nums[0].Valor;
            }
            return "";
        }
        public string GetDescricao()
        {
            var nums = this.GetPropsDescricao();
            if (nums.Count > 0)
            {
                return nums[0].Valor;
            }
            return "";
        }

        public string GetAtributo(string propriedade)
        {
            var retorno = this.GetPropriedades().Find(x => x.Coluna.ToUpper().Replace(" ", "_") == propriedade.ToUpper().Replace(" ", "_"));
            if(retorno!=null)
            {
                return retorno.Valor;
            }
            return "";
        }

        public string GetEtapa()
        {
            var nums = this.GetPropsEtapa();
            if (nums.Count > 0)
            {
                return nums[0].Valor;
            }
            return "";
        }

        public double GetPesoLiquido()
        {
            var nums = this.GetPropsPeso().FindAll(x=>x.Get().Double()>0);
            if (nums.Count > 0)
            {
                return nums[0].Get().Double();
            }
            return 0;
        }
        public double GetComprimento()
        {
            var nums = this.GetPropsComp().FindAll(x => x.Get().Double() > 0);
            if (nums.Count > 0)
            {
                return nums[0].Get().Double();
            }
            return 0;
        }
        public double GetLargura()
        {
            var nums = this.GetPropsLarg().FindAll(x => x.Get().Double() > 0);
            if (nums.Count > 0)
            {
                return nums[0].Get().Double();
            }
            return 0;
        }

        public double GetEspessura()
        {
            var nums = this.GetPropsEspessura().FindAll(x => x.Get().Double() > 0);
            if (nums.Count > 0)
            {
                return nums[0].Get().Double();
            }
            return 0;
        }

        public string Marca { get; private set; } = "";
        public string Tipo { get; private set; } = "";
        private string GetMarca()
        {
           
            if(this.ModelItem!=null)
            {
                var nums = this.GetPropsMarca().FindAll(x=>x.Valor.Replace(" ","").Replace(".","").Replace("-","")!="");

                string nome_fim = "";

                if (nums.Count > 0)
                {
                    nome_fim = nums[0].Valor;
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
                return nums[0].Valor;
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
