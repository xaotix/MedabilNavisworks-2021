using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.ComApi;
using Autodesk.Navisworks.Api.Interop.ComApi;
using Conexoes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MedabilNavisworks
{
   public static class Funcoes
    {
        public static void Apagar_Propriedade(ModelItem item, string tabName, string propertyName)
        {
            Funcoes.Propriedade_Edita_Cria(item, tabName, propertyName, "_remover");
        }
        public static ModelItemCollection GetPecas(string tab, string propriedade)
        {
            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
            Search s = new Search();
            s.Selection.SelectAll();

            SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName(tab, propriedade);
            s.SearchConditions.Add(oSearchCondition);

            ModelItemCollection items = s.FindAll(doc, false);
            return items;
        }
        public static ModelItemCollection GetPecas(string tab, string propriedade, string valor)
        {
            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;

            Search s = new Search();

            s.Selection.SelectAll();
            SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName(tab, propriedade).EqualValue(VariantData.FromDisplayString(valor));

            s.SearchConditions.Add(oSearchCondition);
            ModelItemCollection items = s.FindAll(doc, false);
            return items;
        }
        public static ModelItemCollection GetPecas(string marca)
        {
            Search s = new Search();
            s.Selection.SelectAll();

            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
            SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName(Constantes.Tab, Constantes.Piecemark).EqualValue(VariantData.FromDisplayString(marca));
            s.SearchConditions.Add(oSearchCondition);
            ModelItemCollection items = s.FindAll(doc, false);
            return items;
        }

        public static void Apagar_Propriedade(string tab, string propriedade, ModelItemCollection pcs = null)
        {
            if (pcs == null)
            {
                pcs = GetPecas(tab, propriedade);
            }


            Conexoes.Wait w = new Wait(pcs.Count, $"Apagando propriedade {tab} - {propriedade} de {pcs.Count} Peças");
            w.Show();

            foreach (ModelItem item in pcs)
            {
                w.somaProgresso();
                Apagar_Propriedade(item, tab, propriedade);
            }
            w.Close();
        }

        public static SearchCondition GetFiltroPorTipo(string tipo)
        {
            return SearchCondition.HasPropertyByDisplayName("Item", "Type").DisplayStringContains(tipo).IgnoreStringValueCase();
            //var ttp = SearchCondition.HasPropertyByDisplayName("Item", "Type").IgnoreStringValueCase();
            //ttp.EqualValue(new VariantData(tipo)).IgnoreStringValueCase();
            //return ttp;
        }

        public static string Getvalor(DataProperty valor)
        {
            switch (valor.Value.DataType)
            {
                case VariantDataType.None:
                    return "";
                case VariantDataType.Double:
                    return valor.Value.ToDouble().ToString();
                case VariantDataType.Int32:
                    return valor.Value.ToInt32().ToString();
                case VariantDataType.Boolean:
                    return valor.Value.ToBoolean().ToString();
                case VariantDataType.DisplayString:
                    return valor.Value.ToDisplayString();
                case VariantDataType.DateTime:
                    return valor.Value.ToDateTime().ToShortDateString();
                case VariantDataType.DoubleLength:
                    return valor.Value.ToDoubleLength().ToString();
                case VariantDataType.DoubleAngle:
                    return valor.Value.ToDoubleAngle().ToString();
                case VariantDataType.NamedConstant:
                    return valor.Value.ToNamedConstant().Value.ToString();
                case VariantDataType.IdentifierString:
                    return valor.Value.ToIdentifierString().ToString();
                case VariantDataType.DoubleArea:
                    return valor.Value.ToDoubleArea().ToString();
                case VariantDataType.DoubleVolume:
                    return valor.Value.ToDoubleVolume().ToString();
                case VariantDataType.Point3D:
                    return valor.Value.ToPoint3D().ToString();
                case VariantDataType.Point2D:
                    return valor.Value.ToPoint2D().ToString();
            }

            return "";
        }
        public static List<ModelItem> GetPecas()
        {

            Conexoes.Wait w = new Wait(5, "Mapeando...");
            w.Show();
            w.somaProgresso();

            ModelItemCollection items = new ModelItemCollection();

            Document activeDoc = Autodesk.Navisworks.Api.Application.ActiveDocument;

            Search s = new Search();
            s.Selection.SelectAll();

            List<SearchCondition> condicoes = new List<SearchCondition>();

            condicoes.Add(SearchCondition.HasCategoryByDisplayName("SDS2_Unified"));
            condicoes.Add(SearchCondition.HasCategoryByDisplayName("SDS2_General"));

            condicoes.Add(SearchCondition.HasCategoryByDisplayName("Steel & Graphics Common"));

            condicoes.Add(SearchCondition.HasCategoryByDisplayName("Tekla Common"));


            condicoes.Add(SearchCondition.HasCategoryByDisplayName("Medabil"));

          



            foreach (var cond in condicoes)
            {
                s.SearchConditions.AddGroup(new List<SearchCondition> { cond });
            }

            w.somaProgresso();
            List<string> tips = new List<string>();
            tips.Add(Constantes.IFC_MontagemMarca);
            tips.Add("IFCMEMBER");
            tips.Add("IFCBEAM");
            tips.Add("IFCCOLUMN");
            tips.Add("IFCPLATE");
            tips.Add("IFCFOOTING");
            tips.Add("IFCELEMENT");
            tips.Add("IFCBUILDINGELEMENTPROXY");
            tips.Add("IFCSTAIR");
            /*solda*/
            //tips.Add("IFCFASTENER");


            foreach(var t in tips)
            {
                s.SearchConditions.AddGroup(new List<SearchCondition> { GetFiltroPorTipo(t) });
            }


            List<ModelItem> itens = new List<ModelItem>();
            var itensLista = s.FindAll(activeDoc, false).ToList();
            List<string> tipos = new List<string>();
            w.SetProgresso(1,itensLista.Count, $"Mapeando peças {itensLista.Count}");
           
            foreach (var item in itensLista)
            {
                w.somaProgresso();
                var pai = item.Parent;

                if(pai==null)
                {
                    continue;
                }
                else
                {
                    /*todo = mapear todos os tipos possiveis*/
                    var tipopai = Funcoes.Getvalor(pai, "Type");
                   
                    if(tipopai.ToUpper() == Constantes.IFC_MontagemMarca)
                    {
                        itens.Add(pai);
                    }

                    itens.Add(item);
                }

            }


            itens = itens.GroupBy(x => x.GetHashCode()).Select(x => x.First()).ToList();
            w.Close();
            return itens;
        }

        public static void Destacar(List<ModelItem> itens)
        {
            Document oDoc = Autodesk.Navisworks.Api.Application.ActiveDocument;
            // create ModelItemCollection
            ModelItemCollection oSel_Net = new ModelItemCollection();
            oSel_Net.AddRange(itens);
            //// highlight by .NET API
            //oDoc.CurrentSelection.SelectedItems.CopyFrom(oSel_Net);
            //highlight by COM API
            Autodesk.Navisworks.Api.Interop.ComApi.InwOpState10 state = ComApiBridge.State;
            Autodesk.Navisworks.Api.Interop.ComApi.InwOpSelection comSelectionOut = ComApiBridge.ToInwOpSelection(oSel_Net);
            state.CurrentSelection = comSelectionOut;
        }
        public static void Limpar()
        {

            var items = GetPecas();
            Conexoes.Wait w = new Wait(items.Count, $"Limpando...{items.Count} itens...");
            w.Show();

            foreach (ModelItem item in items)
            {
                var props_medabil = Funcoes.GetPropriedades(item, Constantes.Tab);
                foreach (var prop in props_medabil)
                {
                    Apagar_Propriedade(item, Constantes.Tab, prop.DisplayName);
                }
                //Apagar_Propriedade(item, Constantes.Tab, Constantes.Hierarquia);
                //Apagar_Propriedade(item, Constantes.Tab, Constantes.Nome);
                //Apagar_Propriedade(item, Constantes.Tab, Constantes.Etapa);
                //Apagar_Propriedade(item, Constantes.Tab, Constantes.Piecemark);
                //Apagar_Propriedade(item, Constantes.Tab, Constantes.Numero);
                //Apagar_Propriedade(item, Constantes.Tab, Constantes.Tipo);
                //Apagar_Propriedade(item, Constantes.Tab, Constantes.Peso);
                //Apagar_Propriedade(item, Constantes.Tab, "SKID");
                w.somaProgresso();
            }
            w.Close();
            Conexoes.Utilz.Alerta("Finalizado", "", System.Windows.MessageBoxImage.Information);

        }
        public static void CriaTabDePropriedades(ModelItem item, string user_name) //string tipoObjeto, object objeto)
        {
            if (item == null) return;
            InwOpState10 state;
            state = ComApiBridge.State;

            // set received item as selection
            ModelItemCollection modelItemCollectionIn = new ModelItemCollection();
            modelItemCollectionIn.Add(item);

            // get the selection in COM
            InwOpSelection comSelectionOut =
            ComApiBridge.ToInwOpSelection(modelItemCollectionIn);

            // get paths within the selection
            InwSelectionPathsColl oPaths = comSelectionOut.Paths();
            InwOaPath3 oPath = (InwOaPath3)oPaths.Last();

            // get properties collection of the path
            InwGUIPropertyNode2 propn = (InwGUIPropertyNode2)state.GetGUIPropertyNode(oPath, false);


            // create new property category
            // (new tab in the properties dialog)
            InwOaPropertyVec newPvec = (InwOaPropertyVec)state.ObjectFactory(nwEObjectType.eObjectType_nwOaPropertyVec, null, null);

            // create new property
            InwOaProperty newP = (InwOaProperty)state.ObjectFactory(nwEObjectType.eObjectType_nwOaProperty, null, null);

            bool hasTab = false;
            foreach (InwGUIAttribute2 nwAtt in propn.GUIAttributes())
            {
                if (nwAtt.UserDefined && nwAtt.ClassUserName == user_name) hasTab = true;
            }

            if (!hasTab) propn.SetUserDefined(0, user_name, Conexoes.Utilz.RemoverCaracteresEspeciais(user_name,true), newPvec);
        }
        public static void Propriedade_Edita_Cria(ModelItem item, string tabName, string propertyName, string propertyValue)
        {
            if (item == null) return;
            InwOpState10 state;
            state = ComApiBridge.State;

            // set received item as selection
            ModelItemCollection modelItemCollectionIn = new ModelItemCollection();
            modelItemCollectionIn.Add(item);

            // get the selection in COM
            InwOpSelection comSelectionOut =
            ComApiBridge.ToInwOpSelection(modelItemCollectionIn);

            // get paths within the selection
            InwSelectionPathsColl oPaths = comSelectionOut.Paths();
            InwOaPath3 oPath = (InwOaPath3)oPaths.Last();

            // get properties collection of the path
            InwGUIPropertyNode2 propn = (InwGUIPropertyNode2)state.GetGUIPropertyNode(oPath, false);

            //CreateNewTab
            InwOaPropertyVec newPvec = (InwOaPropertyVec)state.ObjectFactory(nwEObjectType.eObjectType_nwOaPropertyVec, null, null);



            int indexTab = 1;
            bool foundProperty = false;



            foreach (InwGUIAttribute2 nwAtt in propn.GUIAttributes())
            {

                if (!nwAtt.UserDefined) continue;
                if (nwAtt.ClassUserName != tabName)
                {
                    indexTab++;
                    continue;
                }



                //adiciona as propriedades existentes, já modificando a solicitada
                foreach (InwOaProperty nwProp in nwAtt.Properties())
                {
                    InwOaProperty nwNewProp = state.ObjectFactory(nwEObjectType.eObjectType_nwOaProperty);

                    nwNewProp.UserName = nwProp.UserName;
                    nwNewProp.name = nwProp.name;
                    nwNewProp.value = nwProp.value;
                    //modifica a proprieade existente solicitada

                    if (nwNewProp.name == propertyName)
                    {
                        foundProperty = true;
                        nwNewProp.value = propertyValue;
                    }
                    if (nwNewProp.value != "_remover") newPvec.Properties().Add(nwNewProp);
                }


                //caso não tenha achado a propriedade, cria a propriedade
                if (!foundProperty && propertyValue != "_remover")
                {
                    InwOaProperty nwNewProp = state.ObjectFactory(nwEObjectType.eObjectType_nwOaProperty);

                    nwNewProp.UserName = propertyName;
                    nwNewProp.name = Conexoes.Utilz.RemoverCaracteresEspeciais(propertyName,true);
                    nwNewProp.value = propertyValue;
                    newPvec.Properties().Add(nwNewProp);
                }

                newPvec.Properties().Sort();
                propn.SetUserDefined(indexTab, nwAtt.ClassUserName, nwAtt.ClassName, newPvec);
            }
        }
        public static DataProperty GetPropriedade(ModelItem item, string tab, string propriedade)
        {
            var s = item.PropertyCategories.FindPropertyByDisplayName(tab, propriedade);

            if (s == null)
            {
                s = new DataProperty(propriedade, propriedade, new VariantData(""));
            }
            return s;
        }
        public static DataProperty GetPropriedade(ModelItem item, string propriedade)
        {
            var props = GetPropriedades(item);
            var s = props.Find(x => x.DisplayName.ToUpper().Replace(" ", "_") == propriedade.ToUpper().Replace(" ", "_"));
            if (s == null)
            {
                s = new DataProperty(propriedade, propriedade, new VariantData(""));
            }
            return s;
        }

        public static string Getvalor(ModelItem item, string propriedade)
        {
            return Getvalor(GetPropriedade(item, propriedade));
        }

        public static List<DataProperty> GetPropriedades(ModelItem item, string tab = null)
        {
            List<DataProperty> retorno = new List<DataProperty>();
            var s = item.PropertyCategories.ToList();
            foreach (var t in s)
            {
                if (tab != null)
                {
                    if (t.DisplayName.ToUpper().Replace(" ", "_") == tab.ToUpper().Replace(" ", "_"))
                    {
                        retorno.AddRange(t.Properties.ToList());
                    }
                }
                else
                {
                    retorno.AddRange(t.Properties.ToList());
                }
            }
            retorno = retorno.OrderBy(x => x.DisplayName).ToList();




            return retorno;
        }
    }
}
