using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.ComApi;
using Autodesk.Navisworks.Api.Interop.ComApi;
using Autodesk.Navisworks.Api.Plugins;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Web.Script.Serialization;
using Excel = Microsoft.Office.Interop.Excel;
using Application = System.Windows.Forms.Application;

namespace MedabilNavisworks
{
    [Plugin("MedabilRibbon", "CONN",DisplayName = "Medabil - DLM" )]
    [RibbonLayout("MedabilRibbon.xaml")]
    [RibbonTab("MedabilRibbonTab1", DisplayName = "Medabil")]
    [Command("MedabilButtonLimpar", Icon = @"Resources\BTLIMPAR_16.png",            LargeIcon = @"Resources\BTLIMPAR_32.png", DisplayName = "Limpar", ToolTip = "Limpa as informação do último processamento de dados")]
    [Command("MedabilButton1",      Icon = @"Resources\BT1_16.png",                 LargeIcon = @"Resources\BT1_32.png", DisplayName = "Processar", ToolTip = "Processa as informações dos arquivos anexados e preenche os dados da aba Medabil de propriedades")]
    [Command("MedabilButton2",      Icon = @"Resources\BT2_16.png",                 LargeIcon = @"Resources\BT2_32.png", DisplayName = "Importar SKIDs", ToolTip = "Carregar informações de SKID do relatório SAP")]
    [Command("MedabilButton3",      Icon = @"Resources\calendar_16_16.png",         LargeIcon = @"Resources\calendar_32_32.png", DisplayName = "Definir Data Execução", ToolTip = "Define a data de execução dos elementos selecionados")]
    [Command("MedabilButton4",      Icon = @"Resources\calendarRemove_16_16.png",   LargeIcon = @"Resources\calendarRemove_32_32.png", DisplayName = "Remove Data Execução", ToolTip = "Remove a data de execução dos elementos selecionados")]
    [Command("MedabilButton5",      Icon = @"Resources\setsVps_16.png",             LargeIcon = @"Resources\setsVps_32.png", DisplayName = "Sets e Viewpoints", ToolTip = "Gera os Sets e Viewpoints de forma organizada para os elementos executados")]
    [Command("MedabilButton6",      Icon = @"Resources\CalcSelection_16.png",       LargeIcon = @"Resources\CalcSelection_32.png", DisplayName = "Medabil/Tipo", ToolTip = "Apresenta o somatório das propriedades dos elementos selecionados separados por Medabil/Tipo")]
    [Command("MedabilButton7",      Icon = @"Resources\CalcSelection_16.png",       LargeIcon = @"Resources\CalcSelection_32.png", DisplayName = "IFC/OBJECTTYPE", ToolTip = "Apresenta o somatório das propriedades dos elementos selecionados separados por IFC/OBJECTTYPE")]
    [Command("MedabilButton8",      Icon = @"Resources\excelExport_16.png",         LargeIcon = @"Resources\excelExport_32.png", DisplayName = "Exportar Sets", ToolTip = "Exporta os somatórios das propriedades dos elementos dos sets de execução")]
    [Command("Sobre",               Icon = @"Resources\projetabim_16.png",          LargeIcon = @"Resources\projetabim_32.png", DisplayName = "Medabil", ToolTip = "Sobre")]

    //[PluginAttribute("MyPlugin", "ADSK", ToolTip = "My Plugin", DisplayName = "My Plugin")]
    public class Main : CommandHandlerPlugin
    {

        //Dictionary<string, Dictionary<string, string>> sequences = new Dictionary<string, Dictionary<string, string>>();
        Dictionary<string, Sequence> sequences = new Dictionary<string, Sequence>();
        Color cinza = Color.FromByteRGB(171, 171, 171);
        Color Amarelo = Color.FromByteRGB(255, 255, 0);
        Color Verde = Color.FromByteRGB(0, 128, 0);
        DateTime? lastDate = null;
        ProcessingForm processando = new ProcessingForm();
        ModelItem lastMember = null;


        public override int ExecuteCommand(string name, params string[] parameters)
        {

            AppDomain.CurrentDomain.FirstChanceException += (sender, eventArgs) =>
            {
                Debug.WriteLine(eventArgs.Exception.ToString());
            };

            
            //StartProcessMessage();
            switch (name)
            {
                case "MedabilButtonLimpar":
                    Limpar();
                    break;
                case "MedabilButton1":
                    Processar();
                    break;
                case "MedabilButton2":
                    SKIDsImport();
                    break;
                case "MedabilButton3":
                    ExecucaoDateDefine();
                    break;
                case "MedabilButton4":
                    ExecucaoDateDelete();
                    break;
                case "MedabilButton5":
                    ExecucaoCalculate();
                    break;
                case "MedabilButton6":
                    PropertiesSelectionSum("Medabil", "Peso (kg)", "Medabil", "Tipo");
                    break;
                case "MedabilButton7":
                    PropertiesSelectionSum("SDS2_Unified", "Material_Net_Weight", "IFC", "OBJECTTYPE");
                    break;
                case "MedabilButton8":
                    PropertiesSetsSum();
                    break;
                case "Sobre":
                    Sobre();
                    break;
            }
            //StopProcessMessage();
            return 0;
        }

        private void Sobre()
        {
            System.Windows.Forms.MessageBox.Show("Medabil 2020 - ₢\nSuporte: Daniel Lins Maciel\ndaniel.maciel@medabil.com.br");
        }

        private void StartProcessMessage()
        {
            
        }

        private void StopProcessMessage()
        {
            processando.Close(); ;
        }

        private void PropertiesSetsSum()
        {

            IList<string> setsFolders = SETFolderList();
            if (setsFolders.Count == 0)
            {
                MessageBox.Show("Nenhuma pasta de SETs encontrada!");
                return;
            }
            string setFolder = SETsListForm.Wait(setsFolders);

            if (setFolder == "") return;




            Dictionary<string, ModelItemCollection> setsList = SETListCollection(setFolder);

            int counter = 0;
            foreach (KeyValuePair<string, ModelItemCollection> set in setsList)
            {
                counter += set.Value.Count;
            }


            object[,] arrayExport = new object[counter, 6];
            int i = 0;
            foreach (KeyValuePair<string, ModelItemCollection> set in setsList)
            {


                if (set.Value.Count == 0) continue;

                
                //dynamic setSum = PropertiesSum(set.Value, "Medabil", "Peso (kg)", "Medabil", "Tipo");
                //dynamic setToExport = new ExpandoObject();
                foreach(ModelItem item in set.Value)
                {
                    if (item.PropertyCategories.FindCategoryByDisplayName("Medabil") == null)
                    {
                        DataProperty guidProp = item.PropertyCategories.FindPropertyByDisplayName("Item", "GUID");
                        string guid = guidProp != null ? guidProp.Value.ToDisplayString() : "SEM GUID";
                        Debug.Print("------ERRO------- " + guid);
                        continue;
                    }
                    
                    DateTime dataResult;
                    if (DateTime.TryParseExact(set.Key, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out dataResult))
                    {
                        arrayExport[i, 0] = dataResult;
                        
                    }
                    else
                    {
                        arrayExport[i, 0] = set.Key;
                        
                    }
                    
                    //arrayExport[i, 0] = set.Key;
                    arrayExport[i, 1] = item.PropertyCategories.FindPropertyByDisplayName("Medabil", "Etapa").Value.ToDisplayString() ?? "NA";
                    arrayExport[i, 2] = item.PropertyCategories.FindPropertyByDisplayName("Medabil", "Piecemark").Value.ToDisplayString() ?? "NA";
                    arrayExport[i, 3] = item.PropertyCategories.FindPropertyByDisplayName("Medabil", "Numero").Value.ToDisplayString() ?? "NA";
                    arrayExport[i, 4] = item.PropertyCategories.FindPropertyByDisplayName("Medabil", "Tipo").Value.ToDisplayString() ?? "NA";
                    arrayExport[i, 5] = Convert.ToDouble(item.PropertyCategories.FindPropertyByDisplayName("Medabil", "Peso (kg)").Value.ToDisplayString() ?? "0") ;
                   
                    i++;
                }
                /*
                arrayExport[i, 0] = set.Key;
                arrayExport[i, 1] = setSum.peso;
                arrayExport[i, 2] = setSum.contagem;

                //setsSum.Add(setToExport);
                */
            }

            PropertiesSetsExport(arrayExport);



        }




        //private void Teste1()
        //{

        //    Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
        //    //doc.CurrentSelection.SelectAll();
        //    //MessageBox.Show(doc.CurrentSelection.SelectedItems.Count.ToString());
        //    //MessageBox.Show(doc.CurrentSelection.SelectedItems.Count.ToString());

        //    Search s = new Search();

        //    s.Selection.SelectAll();
        //    SearchCondition oSearchCondition = SearchCondition.HasCategoryByDisplayName("SDS2_Unified");
        //    s.SearchConditions.Add(oSearchCondition);
        //    ModelItemCollection groupo1 = s.FindAll(doc, false);

        //    s.Selection.SelectAll();
        //    SearchCondition oSearchCondition2 = SearchCondition.HasCategoryByDisplayName("SDS2_General");
        //    s.SearchConditions.Clear();
        //    s.SearchConditions.Add(oSearchCondition2);
        //    ModelItemCollection groupo2 = s.FindAll(doc, false);

        //    ModelItemCollection items = new ModelItemCollection();
        //    items.AddRange(groupo1);
        //    items.AddRange(groupo2);

        //}

        private void Teste2()
        {
            ExecucaoClear();
        }

        private void SequencesVPs()
        {
            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;

            Search s = new Search();

            s.Selection.SelectAll();
            SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("Medabil", "Hierarquia").EqualValue(VariantData.FromDisplayString("Etapa"));
            s.SearchConditions.Add(oSearchCondition);
            ModelItemCollection items = s.FindAll(doc, false);

        }


        private void Processar()
        {
            ModelItemCollection items = new ModelItemCollection();

            Document activeDoc = Autodesk.Navisworks.Api.Application.ActiveDocument;
            //doc.CurrentSelection.SelectAll();
            //MessageBox.Show(doc.CurrentSelection.SelectedItems.Count.ToString());
            //MessageBox.Show(doc.CurrentSelection.SelectedItems.Count.ToString());

        

           

            Search s = new Search();
            s.Selection.SelectAll();
            SearchCondition oSearchCondition = SearchCondition.HasCategoryByDisplayName("SDS2_Unified");
            s.SearchConditions.Add(oSearchCondition);
            ModelItemCollection sds = s.FindAll(activeDoc, false);

            s.Selection.SelectAll();
            SearchCondition oSearchCondition2 = SearchCondition.HasCategoryByDisplayName("SDS2_General");
            s.SearchConditions.Clear();
            s.SearchConditions.Add(oSearchCondition2);
            ModelItemCollection sds2 = s.FindAll(activeDoc, false);


            Search s1 = new Search();
            s1.PruneBelowMatch = false;
            s1.SearchConditions.Clear();
            ModelItemCollection micTmp = new ModelItemCollection();
            foreach (var mod in activeDoc.Models)
            {
                micTmp.Add(mod.RootItem);
            }
            s1.Selection.CopyFrom(micTmp);
            ////Daniel Maciel
            items.AddRange(sds);
            items.AddRange(sds2);



            //se não acha nada, é porque talvez o arquivo não tenha marcas
            //mapeia via força bruta
            if (items.Count() == 0)
            {
              
                ModelItemCollection searchResults = new ModelItemCollection();
                s1.SearchConditions.Add(SearchCondition.HasCategoryByName(PropertyCategoryNames.Item));
                searchResults = s1.FindAll(Autodesk.Navisworks.Api.Application.ActiveDocument, false); 
                s1.PruneBelowMatch = false;
                s1.SearchConditions.Clear();
                var oss = searchResults.ToList().FindAll(x => !x.IsHidden).ToList();
                //ICON 4 = PROCURA TUDO QUE TEM UM GRUPO
                List<string> objetos = new List<string>();
                objetos.Add("IFCELEMTASSEMBLY");
                objetos.Add("IFCBEAM");
                objetos.Add("IFCCOLUMN");
                objetos.Add("IFCPLATE");
                objetos.Add("IFCFOOTING");
                objetos.Add("IFCMEMBER");
                objetos.Add("IFCELEMENT");

                var OBJS = GetObjetos(searchResults, "Item", "Type", objetos).ToList();
                items.AddRange(OBJS);
            }


            var itens = items.GroupBy(x => x).Select(x => x.First()).ToList();
            List<ModelItem> pp = new List<ModelItem>();
            items.Clear();

            var itss = itens.FindAll(x => !x.IsHidden).ToList();
            foreach (var it in itens)
            {

                ModelItem member, etapa;
                GetMembroPrincipal(it, out member, out etapa);
                if (member == null)
                { continue; }

                if(member!=null)
                {
                    pp.Add(member);
                }
                //if(it.Parent!=null)
                //{
                //    if(it.Children.Count()==0)
                //    {
                //        pp.Add(it.Parent);
                //    }
                //    else
                //    {

                //        pp.Add(it);
                //    }
                //}
            }
            pp = pp.FindAll(x => !x.IsHidden).ToList();
            pp = pp.GroupBy(x => x.GetHashCode()).Select(x => x.First()).ToList();
            //pp = pp.OrderBy(x => x).ToList().GroupBy(x => x).Select(x => x.First()).ToList();
            //MessageBox.Show(itens.Count.ToString());
            Loading mm = new Loading();
            mm.progressBar1.Maximum = pp.Count;
            mm.label1.Text = "Mapeando... " + pp.Count + " encontrados...";
            mm.Show();
            foreach (ModelItem item in pp)
            {
                if (item.Parent != null)
                {
                    Mapear(item);
                    mm.progressBar1.Value = mm.progressBar1.Value + 1;
                }

            }
            mm.Close();

            PropertiesSequencesProcess();
            sequences = new Dictionary<string, Sequence>();

            System.Windows.Forms.MessageBox.Show("Finalizado.");
        }

        private static List<DataProperty> GetPropriedades(ModelItemCollection searchResults, string categoria, string propriedade)
        {
            return searchResults.Select(x => x.PropertyCategories.ToList()
            .Find(y => y.DisplayName.ToUpper().Contains(categoria)))
                .SelectMany(x => x.Properties).ToList().FindAll(y => y.Name.ToUpper().Contains(propriedade)| y.DisplayName.ToUpper().Contains(propriedade)).ToList();
        }

        private static ModelItemCollection GetObjetos(ModelItemCollection searchResults, string categoria, string propriedade, List<string> valores)
        {
            if(searchResults.Count==0)
            {
                return new ModelItemCollection();
            }
            Loading mm = new Loading();
            mm.label1.Text = "Encontrando peças";
            mm.progressBar1.Maximum = searchResults.Count;

            mm.Show();
            mm.progressBar1.Value = mm.progressBar1.Value + 1;
            mm.progressBar1.Value = mm.progressBar1.Value + 1;
            //var pcs = searchResults.ToList().FindAll(x=>x.HasGeometry).ToList();
            List<ModelItem> pp = new List<ModelItem>();
            mm.progressBar1.Value = 0;
            foreach (var p in searchResults)
            {
                mm.progressBar1.Value = mm.progressBar1.Value + 1;

                if(!p.IsHidden)
                {
                    if(p.Parent!=null)
                    {
                        if(!p.Parent.IsHidden)
                        {

                        }
                    }
                }



                ModelItem member= null;

                if (TemPropriedade(categoria, propriedade, valores, p))
                {
                    member = p;
                }
            

                if(member ==null)
                {
                    foreach (var s in p.Descendants)
                    {
                        member = Validar(categoria, propriedade, valores,  s);
                        if(member!=null)
                        {
                            break;
                        }
                    }
                }





                if(member == null)
                {

                    foreach (var s in p.Ancestors)
                    {
                  member = Validar(categoria, propriedade, valores, s);
                        if (member != null)
                        {
                            break;
                        }
                    }
                }





                if(member!=null)
                {
                    pp.Add(member);
                }



                //ModelItem member, etapa;
                //GetMembroPrincipal(p, out member, out etapa);
                //if (member == null)
                //{ continue; }


                //if (member.Children.Count() == 0)
                //{
                //    continue;
                //}

                 




            }
            mm.Close();
            ModelItemCollection retorno = new ModelItemCollection();

            pp = pp.GroupBy(x => x).Select(x => x.First()).ToList();
            retorno.AddRange(pp);
            return retorno;
        }

        private static ModelItem Validar(string categoria, string propriedade, List<string> valores, ModelItem nivel0)
        {
            ModelItem member = null;
            var cod = TemPropriedade(categoria, propriedade, valores, nivel0);
            if (cod)
            {
                member = nivel0;
                var nivel1 = member.Parent;
                if (nivel1 != null)
                {
                    if(TemPropriedade(categoria, propriedade, valores, nivel1))
                    {
                        member = nivel1;
                    }
                    var nivel2 = nivel1.Parent;
                    if (nivel2 != null)
                    {
                        if (TemPropriedade(categoria, propriedade, valores, nivel2))
                        {
                            member = nivel2;
                        }
                        var nivel3 = nivel2.Parent;
                        if (nivel3 != null)
                        {
                            if (TemPropriedade(categoria, propriedade, valores, nivel3))
                            {
                                member = nivel3;
                            }
                            var nivel4 = nivel3.Parent;
                            if (nivel4 != null)
                            {
                                if (TemPropriedade(categoria, propriedade, valores, nivel4))
                                {
                                    member = nivel4;
                                }
                                var nivel5 = nivel4.Parent;
                                if (nivel5 != null)
                                {
                                    if (TemPropriedade(categoria, propriedade, valores, nivel5))
                                    {
                                        member = nivel5;
                                    }
                                }
                            }
                        }
                    }
                  
                   
                }
            }

            return member;
        }

        private static bool TemPropriedade(string categoria, string propriedade, List<string> valores, ModelItem member)
        {
            var s = member.PropertyCategories.FindPropertyByDisplayName(categoria, propriedade);
            if (s != null)
            {
                if (s.Value.IsDisplayString)
                {
                    var sst = s.Value.ToDisplayString();
                    foreach (var valor in valores)
                    {
                        if (sst.ToUpper().Contains(valor.ToUpper()))
                        {
                            return true;
                        }
                    }

                }
                else if (s.Value.IsInt32)
                {
                    var sst = s.Value.ToInt32();
                    foreach (var valor in valores)
                    {
                        if (sst.ToString().ToUpper().Contains(valor.ToUpper()))
                        {
                            return true;
                        }
                    }
                }
                else if (s.Value.IsNamedConstant)
                {
                    int sst = s.Value.ToNamedConstant().Value;
                    foreach (var valor in valores)
                    {
                        if (sst.ToString().ToUpper().Contains(valor.ToUpper()))
                        {
                            return true;
                        }
                    }
                }



            }
            else
            {

            }
            return false;
        }

        private void Limpar()
        {


            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
            //doc.CurrentSelection.SelectAll();
            //MessageBox.Show(doc.CurrentSelection.SelectedItems.Count.ToString());
            //MessageBox.Show(doc.CurrentSelection.SelectedItems.Count.ToString());

            ModelItemCollection items = new ModelItemCollection();
            items.AddRange(doc.Models.RootItemDescendantsAndSelf);

            
            //MessageBox.Show(items.Count.ToString());

            foreach (ModelItem item in items)
            {
                PropertyDelete(item, "Medabil", "Hierarquia");
                PropertyDelete(item, "Medabil", "Nome");
                PropertyDelete(item, "Medabil", "Etapa");
                PropertyDelete(item, "Medabil", "Piecemark");
                PropertyDelete(item, "Medabil", "Numero");
                PropertyDelete(item, "Medabil", "Tipo");
                PropertyDelete(item, "Medabil", "Peso");
            }


        }

        private void Mapear(ModelItem item)
        {

            List<PropertyCategory> propriedades = new List<PropertyCategory>();

            List<PropertyCategory> pcs_identification = new List<PropertyCategory>();
            //17/03/2020
            pcs_identification.AddRange(item.PropertyCategories.Where(p => { return p.DisplayName.ToUpper().Contains("SDS"); }));
            pcs_identification.AddRange(item.PropertyCategories.Where(p => { return p.DisplayName.ToUpper().Contains("TEKLA"); }));
            pcs_identification.AddRange(item.PropertyCategories.Where(p => { return p.DisplayName.ToUpper().Contains("PSET"); }));
            pcs_identification.AddRange(item.PropertyCategories.Where(p => { return p.DisplayName.ToUpper().Contains("STEEL GRAPHICS COMMON"); }));
            pcs_identification.AddRange(item.PropertyCategories.Where(p => { return p.DisplayName.ToUpper().Contains("IFC"); }));
            pcs_identification.AddRange(item.PropertyCategories.Where(p => { return p.DisplayName.ToUpper().Contains("ITEM"); }));
            //MessageBox.Show(pcs_identification.Count.ToString());
            propriedades.AddRange(pcs_identification);
            string etapa_str = "";
            string Member_Number_String = "";
            string marca_string = "";
            double peso=0;
            string Member_Type_String = "";

            //vai procurando pelas propriedades e se encontrar, seta o valor.
            foreach (PropertyCategory pc in propriedades)
            {
                try
                {
                
                    DataProperty Sequencia = item.PropertyCategories.FindPropertyByDisplayName(pc.DisplayName, "Sequence");
                    DataProperty Numero = item.PropertyCategories.FindPropertyByDisplayName(pc.DisplayName, "Member_Number");
                    DataProperty Marca = item.PropertyCategories.FindPropertyByDisplayName(pc.DisplayName, "Member_Piecemark");
                    DataProperty Peso = item.PropertyCategories.FindPropertyByDisplayName(pc.DisplayName, "Member_Net_Weight");
                    DataProperty Tipo = item.PropertyCategories.FindPropertyByDisplayName(pc.DisplayName, "Member_Type");

                    if (Marca == null)
                    {
                        Marca = item.PropertyCategories.FindPropertyByDisplayName(pc.DisplayName, "Assembly/Cast unit Mark");
                    }
                    if (Marca == null)
                    {
                        Marca = item.PropertyCategories.FindPropertyByDisplayName(pc.DisplayName, "TAG");
                    }

                    if (Marca == null)
                    {
                        Marca = item.PropertyCategories.FindPropertyByDisplayName(pc.DisplayName, "Assembly Mark");
                    }
                    if (Marca==null)
                    {
                        Marca = item.PropertyCategories.FindPropertyByDisplayName(pc.DisplayName, "DESCRIPTION");
                    }


                    if (Sequencia == null)
                    {
                        Sequencia = item.PropertyCategories.FindPropertyByDisplayName(pc.DisplayName, "Phase");
                    }
                    if (Numero == null)
                    {
                        Numero = item.PropertyCategories.FindPropertyByDisplayName(pc.DisplayName, "GUID");
                    }
                    
                    if (Marca == null)
                    {
                        Marca = item.PropertyCategories.FindPropertyByDisplayName(pc.DisplayName, "Reference");

                    }

                    if (Peso == null)
                    {
                        Peso = item.PropertyCategories.FindPropertyByDisplayName(pc.DisplayName, "Assembly/Cast unit weight");
                    }
                    if (Tipo == null)
                    {
                        Tipo = item.PropertyCategories.FindPropertyByDisplayName(pc.DisplayName, "Type");
                    }

                    //17/03/2020 - para marcas simples do tekla
                    if (Marca == null)
                    {
                        Marca = item.PropertyCategories.FindPropertyByDisplayName(pc.DisplayName, "Type");
                    }
                    if (Peso == null)
                    {
                        Peso = item.PropertyCategories.FindPropertyByDisplayName(pc.DisplayName, "Weight");
                    }



   
                    if (Numero == null)
                    {
                        Numero = item.PropertyCategories.FindPropertyByDisplayName(pc.DisplayName, "Member_Number");
                    }
                    if (Marca == null)
                    {
                        Marca = item.PropertyCategories.FindPropertyByDisplayName(pc.DisplayName, "BATID");
                    }




                    if (Sequencia != null && etapa_str == "") etapa_str = Sequencia.Value.ToDisplayString();
                    if (Marca != null && marca_string == "") marca_string = Marca.Value.ToDisplayString();
                    if (Peso != null && peso == 0) peso = Peso.Value.ToAnyDouble();
                    if (Tipo != null && Member_Type_String == "") Member_Type_String = Tipo.Value.ToDisplayString();
                    if (Numero != null && Member_Number_String == "") Member_Number_String = Numero.Value.ToDisplayString();


                    //15/04/2020 - para ler os inputs de TecnoMetal
                    if(marca_string.ToUpper().Contains("MARK") && marca_string.ToUpper().Contains("POS"))
                    {
                        var m = marca_string.Split(' ').ToList();
                        marca_string = m[0].ToUpper().Replace(" ", "").Replace("MARK", "").Replace(":", "");
                    }

                    if(marca_string.Contains(" "))
                    {
                        marca_string = item.DisplayName;
                    }
                    else if(marca_string  == Member_Type_String)
                    {
                        marca_string = item.DisplayName;
                    }
                }
                catch (Exception ex)
                {
                   // MessageBox.Show(ex.Message);
                }
             
            }

            if (marca_string == "") { return; }

            try
            {
                ModelItem member, etapa;
                GetMembroPrincipal(item, out member, out etapa);
                if(member == null) { return; }
                if (lastMember != null && member == lastMember) return;


                lastMember = member;
                //adiciona a medabil tab na sequence
                if (!sequences.ContainsKey(etapa_str))
                {
                    Sequence newSequence = new Sequence("sequence", etapa_str);
                    sequences.Add(etapa_str, newSequence);
                    PropertiesMedabilTabCreate(etapa, "Medabil", "Medabil");
                    PropertyCreateOrEdit(etapa, "Medabil", "Hierarquia", "Etapa", "Hierarquia");
                    PropertyCreateOrEdit(etapa, "Medabil", "Nome", etapa_str, "Nome");
                }

                //adiciona a medabil tab nos membros
                Sequence sequence = sequences[etapa_str];


                Member newMember = new Member(
                    "member",
                    Member_Number_String,
                    peso,
                    marca_string,
                    Member_Type_String
                    );
                PropertiesMedabilTabCreate(member, "Medabil", "Medabil");
                PropertyCreateOrEdit(member, "Medabil", "Hierarquia", newMember.TipoObjeto, "Hierarquia");
                PropertyCreateOrEdit(member, "Medabil", "Etapa", sequence.Name, "Etapa");
                PropertyCreateOrEdit(member, "Medabil", "Piecemark", newMember.Piecemark, "Piecemark");
                PropertyCreateOrEdit(member, "Medabil", "Numero", newMember.Number, "Numero");
                PropertyCreateOrEdit(member, "Medabil", "Tipo", newMember.Type, "Tipo");
                PropertyCreateOrEdit(member, "Medabil", "Peso", newMember.NetWeight.ToString(), "Peso (kg)");

                //if (newMember.Type == "MISCELANEA") return;

                sequence.Members.Add(newMember);
                if (!sequence.TypesCounter.ContainsKey(newMember.Type))
                {
                    sequence.TypesCounter.Add(newMember.Type, 1);
                    sequence.TypesNetWeight.Add(newMember.Type, newMember.NetWeight);

                }
                else
                {
                    sequence.TypesCounter[newMember.Type]++;
                    sequence.TypesNetWeight[newMember.Type] += newMember.NetWeight;

                }

                sequence.NetWeight += newMember.NetWeight;
            }
            catch (Exception)
            {

            }

   

 
            
            
            
        }
        public static double Double(object comp, int Decimais = 4)
        {
                    System.Globalization.CultureInfo US = new System.Globalization.CultureInfo("en-US");
       System.Globalization.CultureInfo BR = new System.Globalization.CultureInfo("pt-BR");
            //string comps = comp.ToString();
            //if (comps == "") { comps = "0"; }
            //try
            //{
            //    return Math.Round(double.Parse(comps.Replace(",", "."), System.Globalization.CultureInfo.InvariantCulture),Decimais);
            //    //return Math.Round(Convert.ToDouble(comps.Replace(".",",")),Decimais);
            //}
            //catch (Exception)
            //{

            //    return 0;
            //}
            try
            {

                double val;
                if (double.TryParse(comp.ToString().Replace(" ", ""), System.Globalization.NumberStyles.Float, BR, out val))
                {
                    try
                    {
                        return Math.Round(val, Decimais);

                    }
                    catch (Exception)
                    {

                        return val;
                    }
                }

                else if (double.TryParse(comp.ToString(), System.Globalization.NumberStyles.Float, US, out val))
                {
                    try
                    {

                        return Math.Round(val, Decimais);
                    }
                    catch (Exception)
                    {
                        return val;
                    }
                }
                else return 0;
            }
            catch (Exception)
            {

                return 0;
            }


        }
        private static void GetMembroPrincipal(ModelItem item, out ModelItem member, out ModelItem etapa)
        {

            int cod1 = item.PropertyCategories.FindPropertyByDisplayName("Item", "Icon").Value.ToNamedConstant().Value;
            if (cod1 == 4)
            {
                member = item;
                etapa = item.Parent;
                return;
            }
            //if(item.Parent==null)
            //{ etapa = null;
            //    member = item;
            //    return;
            //}

            //var itens = item.AncestorsAndSelf.ToList();
            //itens.AddRange(item.Descendants.ToList());
            //Lê todos os itens e procura pelo objeto pai.
            foreach (var s in item.Descendants)
            {
                int cod = GetTipo(s);
                if (cod == 4)
                {
                    member = s;
                    etapa = s.Parent;

                    if (member.Parent != null)
                    {
                        var tp = GetTipo(member.Parent);
                        if(member.Parent.Parent!=null)
                        {
                          var  tp2 = GetTipo(member.Parent.Parent);
                            if(tp2== 4)
                            {
                                member = member.Parent.Parent;
                                etapa = member.Parent.Parent.Parent;
                                return;
                            }
                        }
                        if (tp == 4)
                        {
                            member = member.Parent;
                            etapa = member.Parent.Parent;
                            return;
                        }
                    }
                    return;
                }
            }

           


            foreach (var s in item.Ancestors)
            {
                int cod = GetTipo(s);
                if (cod == 4)
                {
                    member = s;
                    etapa = s.Parent;

                    if (member.Parent != null)
                    {
                        var tp = GetTipo(member.Parent);
                        if (member.Parent.Parent != null)
                        {
                            var tp2 = GetTipo(member.Parent.Parent);
                            if (tp2 == 4)
                            {
                                member = member.Parent.Parent;
                                etapa = member.Parent.Parent.Parent;
                                return;
                            }
                        }
                        if (tp == 4)
                        {
                            member = member.Parent;
                            etapa = member.Parent.Parent;
                            return;
                        }
                    }
                    return;
                }
            }
          
            member = null;
            etapa = null;
            return;

            int Parent_Icon = GetTipo(item);
            if (Parent_Icon == 4)
            {

                member = item.Parent;
                etapa = item.Parent.Parent;
            }
            else
            {
                member = item;
                etapa = item.Parent;
            }
        }

        private static int GetTipo(ModelItem s)
        {
            return s.PropertyCategories.FindPropertyByDisplayName("Item", "Icon").Value.ToNamedConstant().Value;
        }

        private void PropertiesSequencesProcess()
        {

            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
            //doc.CurrentSelection.SelectAll();
            //MessageBox.Show(doc.CurrentSelection.SelectedItems.Count.ToString());
            //MessageBox.Show(doc.CurrentSelection.SelectedItems.Count.ToString());
            foreach (KeyValuePair<string, Sequence> seq in sequences)
            {
                Search s = new Search();

                s.Selection.SelectAll();
                SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("Medabil", "Nome").EqualValue(VariantData.FromDisplayString( seq.Key ));

                s.SearchConditions.Add(oSearchCondition);
                ModelItem item = s.FindFirst(doc, false);
                PropertyCreateOrEdit(item, "Medabil", "Hierarquia", "Etapa", "Hierarquia");
                PropertyCreateOrEdit(item, "Medabil", "Nome", seq.Value.Name, "Nome");
                PropertyCreateOrEdit(item, "Medabil", "MembersCount", seq.Value.Members.Count.ToString(), "Elementos (Quantidade)");
                PropertyCreateOrEdit(item, "Medabil", "MemebersMass", seq.Value.NetWeight.ToString() , "Elementos (Peso - kg)");
                foreach(KeyValuePair<string,int> typeCount in seq.Value.TypesCounter)
                {
                    if (typeCount.Key == "MISCELANEA") continue;
                    PropertyCreateOrEdit(item, "Medabil", typeCount.Key+"Count" , typeCount.Value.ToString(), typeCount.Key + " (Quantidade)");
                    PropertyCreateOrEdit(item, "Medabil", typeCount.Key+"Peso", seq.Value.TypesNetWeight[typeCount.Key].ToString() , typeCount.Key + " (Peso - kg)");
                    
                }

            }





        }

        private void PropertiesMedabilTabCreate(ModelItem item, string user_name, string internal_name) //string tipoObjeto, object objeto)
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
                if (nwAtt.UserDefined && nwAtt.ClassUserName == "Medabil") hasTab = true;
            }

            if(!hasTab) propn.SetUserDefined(0, "Medabil", "Medabil", newPvec);
        }

        private InwOaProperty PropertyCreate(InwOpState10 state, string name, string username, string value)
        {
            // create new property

            InwOaProperty newP =

                (InwOaProperty)state.ObjectFactory(

                nwEObjectType.eObjectType_nwOaProperty, null, null);



            // set the name, username and value of the new property

            newP.name = name;

            newP.UserName = username;

            newP.value = value;

            return newP;

        }
        //private static List<string> filtro { get; set; } = new List<string>();
        //private static List<string> Carregar()
        //{
        //    var arq = Application.StartupPath + @"\filtro.txt";
        //    string[] lines = File.ReadAllLines(arq);
        //    List<string> linhas = new List<string>();
        //    foreach (string line in lines)
        //        linhas.Add(line);

        //    return linhas;
        //}

        private void PropertyCreateOrEdit(ModelItem item, string tabName, string propertyName, string propertyValue, string propertyUserName = "NewUserProperty")
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
                

                //MessageBox.Show("newAtt => " + nwAtt.ClassName + " ___ " + nwAtt.ClassUserName + " ___ " + nwAtt.name);

                //adiciona as propriedades existentes, já modificando a solicitada
                foreach (InwOaProperty nwProp in nwAtt.Properties())
                {
                    InwOaProperty nwNewProp = state.ObjectFactory(nwEObjectType.eObjectType_nwOaProperty);

                    nwNewProp.UserName = nwProp.UserName;
                    nwNewProp.name = nwProp.name;
                    nwNewProp.value = nwProp.value;
                    //modifica a proprieade existente solicitada
                    
                    if (nwNewProp.name == propertyName )
                    {
                        
                        foundProperty = true;

                        nwNewProp.value = propertyValue;
                    }
                    if(nwNewProp.value != "_remover") newPvec.Properties().Add(nwNewProp);
                }

                //if (propertyValue == "_remover") return;
               
                //caso não tenha achado a propriedade, cria a propriedade
                if (!foundProperty && propertyValue != "_remover")
                {
                    InwOaProperty nwNewProp = state.ObjectFactory(nwEObjectType.eObjectType_nwOaProperty);

                    nwNewProp.UserName = propertyUserName;
                    nwNewProp.name = propertyName;
                    nwNewProp.value = propertyValue;
                    newPvec.Properties().Add(nwNewProp);
                }

                newPvec.Properties().Sort();
                propn.SetUserDefined(indexTab, nwAtt.ClassUserName, nwAtt.ClassName, newPvec);
                

            }
        }

        private void SKIDsImport()
        {
            
            object[,] retorno = ExcelImportToArray();
            
            SKIDsClear();
            Dictionary<string, string> skids = new Dictionary<string, string>();
            ModelItemCollection multipleSkid = new ModelItemCollection();
            List<string> skidList = new List<string>();

            

            
            for (int i = 1; i <= retorno.GetLength(0); i++)
            {
                object pmObj = retorno[i, 12];
                object skidObj = retorno[i, 3];

                if (pmObj == null || skidObj == null) continue;
                string pm = pmObj.ToString();
                string skid = skidObj.ToString();

                if (pm.Replace(" ", "") == "" || skid.Replace(" ", "") == "") continue;
                
                if (!skidList.Contains(skid)) skidList.Add(skid);

                if (!skids.ContainsKey(pm)) skids.Add(pm, skid);
                else if (!skids[pm].Contains(skid)) skids[pm] = skids[pm] + ", " + skid;
                
            }

            Dictionary<string, string> sequencesSkids = new Dictionary<string, string>();
                
            
            foreach (KeyValuePair<string, string> skid in skids)
            {

                Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;

                Search s = new Search();

                s.Selection.SelectAll();
                SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("Medabil", "Piecemark").EqualValue(VariantData.FromDisplayString(skid.Key));
                s.SearchConditions.Add(oSearchCondition);
                ModelItemCollection items = s.FindAll(doc, false);

                
                
                foreach (ModelItem item in items)
                {
                    PropertyCreateOrEdit(item, "Medabil", "SKID", skid.Value, "SKID");
                    bool multiSkids = false;
                    if (skid.Value.Contains(",")) multiSkids = true;
                    if (multiSkids)
                    {
                        PropertyCreateOrEdit(item, "Medabil", "SKID_Multiple", "Yes", "SKID_Multiple");
                        continue;
                    }
                    //insere o skid na sequenceList
                    string sequence = item.PropertyCategories.FindPropertyByDisplayName("Medabil", "Etapa").Value.ToDisplayString();
                    string sk = skid.Value;
                    if (!sequencesSkids.ContainsKey(sequence)) sequencesSkids.Add(sequence, sk);
                    else if (!sequencesSkids[sequence].Contains(sk)) sequencesSkids[sequence] = sequencesSkids[sequence] + ", " + sk;


                    /*NÃO DELETAR --> código para quando for multiskids
                    string sequence = item.PropertyCategories.FindPropertyByDisplayName("Medabil", "Etapa").Value.ToDisplayString();
                    string[] skidsArray = skid.Value.Replace(" ","").Split(","[0]);
                    foreach(string sk in skidsArray)
                    {
                        if(!sequencesSkids.ContainsKey(sequence)) sequencesSkids.Add(sequence, skidsArray[0]);
                        else if(!sequencesSkids[sequence].Contains(sk)) sequencesSkids[sequence] = sequencesSkids[sequence] + ", " + sk;
                    }
                    */
                }

                




            }
            
            SKIDsSequences(sequencesSkids);
            SKIDsSetsCreate(skidList);
            SKIDsSetsCreateForMultiSkid();
            SKIDsClearMultSkids();
            



            
            

            

        }

        private void SKIDsSequences(Dictionary<string, string> sequencesSkids)
        {
            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
            foreach (KeyValuePair<string, string> sq in sequencesSkids)
            {
                
                Search s = new Search();

                s.Selection.SelectAll();
                SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("Medabil", "Nome").EqualValue(VariantData.FromDisplayString(sq.Key ));
                s.SearchConditions.Add(oSearchCondition);
                SearchCondition oSearchCondition2 = SearchCondition.HasPropertyByDisplayName("Medabil", "Hierarquia").EqualValue(VariantData.FromDisplayString("Etapa"));
                s.SearchConditions.Add(oSearchCondition2);
                ModelItem sequenceItem = s.FindFirst(doc, false);
                string skidsText = string.Format("{0} SKID(s): {1}",
                    sq.Value.Split(","[0]).Length.ToString(),
                    sq.Value
                    );
                PropertyCreateOrEdit(sequenceItem, "Medabil", "SKIDs", skidsText, "SKIDs");
                
            }
        }

        private void SKIDsSetsCreate(List<string> skids)
        {
            
            SETFolderDelete("SKIDs");
            
            skids.OrderByDescending(x => x);
            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
            foreach(string skid in skids)
            {

                Search s = new Search();

                s.Selection.SelectAll();
                SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("Medabil", "SKID").EqualValue(VariantData.FromDisplayString(skid));
                s.SearchConditions.Add(oSearchCondition);
                ModelItemCollection items = s.FindAll(doc, false);

                if(items.Count > 0) SETCreateOrEdit("SKIDs", skid, items);


            }

            

            
        }


        private void SKIDsSetsCreateForMultiSkid()
        {
            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
            Search s = new Search();

            s.Selection.SelectAll();
            SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("Medabil", "SKID_Multiple");
            s.SearchConditions.Add(oSearchCondition);
            ModelItemCollection items = s.FindAll(doc, false);

            if (items.Count > 0) SETCreateOrEdit("SKIDs", "Multiplo", items);
        }

        private void SETCreateOrEdit(string folderName, string setName, ModelItemCollection items)
        {
            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
            var cs = items;
            var ss = doc.SelectionSets;
            

            var fn = folderName;
            var sn = setName;

            try
            {
                var set = new SelectionSet(cs)
                {
                    DisplayName = sn
                };

                var fi = ss.Value.IndexOfDisplayName(fn);

                if (fi == -1)
                {
                    var sf = new FolderItem() { DisplayName = fn };
                    sf.Children.Add(set);
                    ss.AddCopy(sf);
                }
                else
                {
                    ss.AddCopy(set);

                    fi = ss.Value.IndexOfDisplayName(fn);
                    var fo = ss.Value[fi] as FolderItem;

                    var si = ss.Value.IndexOfDisplayName(set.DisplayName);
                    var se = ss.Value[si] as SavedItem;
                    
                    ss.Move(se.Parent, si, fo, 0);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private Dictionary<string, ModelItemCollection> SETListCollection(string folderName)
        {
            Dictionary<string, ModelItemCollection> setsList = new Dictionary<string, ModelItemCollection>();
            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
            
            var ss = doc.SelectionSets;
            var fn = folderName;
            var fi = ss.Value.IndexOfDisplayName(fn);
            

            if (fi == -1) return setsList;


            var fo = ss.Value[fi] as FolderItem;
            foreach(SavedItem sv in fo.Children)
            {
                SelectionSet svSet = sv as SelectionSet;
                setsList.Add(svSet.DisplayName, SetDeepLook(svSet) );
            }

            return setsList;
                           
        }


        private void SetFolderToDeep(string folderName, bool deepLook = false)
        {
            Dictionary<string, ModelItemCollection> setsList = new Dictionary<string, ModelItemCollection>();
            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;

            var ss = doc.SelectionSets;
            var fn = folderName;
            var fi = ss.Value.IndexOfDisplayName(fn);


            if (fi == -1) return;


            var fo = ss.Value[fi] as FolderItem;
            foreach (SavedItem sv in fo.Children)
            {
                SelectionSet svSet = sv as SelectionSet;
                SETCreateOrEdit(fn + "Deep", svSet.DisplayName, SetDeepLook(svSet));
            }

        }

        private ModelItemCollection SetDeepLook(SelectionSet Set)
        {
            ModelItemCollection items = new ModelItemCollection();
            foreach (ModelItem item in Set.GetSelectedItems())
            {
                DataProperty hierarquia = item.PropertyCategories.FindPropertyByDisplayName("Medabil", "Hierarquia");
                try
                {

                    if (hierarquia == null)
                    {
                        IEnumerable<ModelItem> membrosPais = from x in item.Ancestors where x.PropertyCategories.FindPropertyByDisplayName("Medabil", "Hierarquia") != null && x.PropertyCategories.FindPropertyByDisplayName("Medabil", "Hierarquia").Value.ToDisplayString() == "member" select x;
                        items.AddRange(membrosPais);

                        IEnumerable<ModelItem> membrosFilhos = from x in item.Descendants where x.PropertyCategories.FindPropertyByDisplayName("Medabil", "Hierarquia") != null && x.PropertyCategories.FindPropertyByDisplayName("Medabil", "Hierarquia").Value.ToDisplayString() == "member" select x;
                        items.AddRange(membrosFilhos);


                    }
                    else if (hierarquia.Value.ToDisplayString() == "etapa")
                    {
                        IEnumerable<ModelItem> membros = from x in item.Descendants where x.PropertyCategories.FindPropertyByDisplayName("Medabil", "Hierarquia") != null select x;
                        items.AddRange(membros);
                    }
                    else if (hierarquia.Value.ToDisplayString() == "member")
                    {
                        items.Add(item);
                    }
                }
                catch (InvalidCastException e)
                {
                    Debug.Print(e.Message);
                    continue;
                }

            }

            return items;

            
        }

        private IList<string> SETFolderList()
        {
            IList<string> setFoldesList = new List<string>();
            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;

            var ss = doc.SelectionSets ;
            foreach(SavedItem set in doc.SelectionSets.Value )
            {
                if (set.IsGroup) setFoldesList.Add(set.DisplayName);
            }

            return setFoldesList;

        }


        

        private void ExecucaoCreateSheet()
        {

            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
           
            Search s = new Search();

            s.Selection.SelectAll();
            SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("Medabil", "Hierarquia").EqualValue(VariantData.FromDisplayString("member"));
            SearchCondition oSearchCondition2 = SearchCondition.HasPropertyByDisplayName("Medabil", "Tipo").EqualValue(VariantData.FromDisplayString("MISCELANEA")).Negate();

            s.SearchConditions.Add(oSearchCondition);
            s.SearchConditions.Add(oSearchCondition2);


            ModelItemCollection items = s.FindAll(doc, false);


            OpenFileDialog ofd = new OpenFileDialog();
            DialogResult dr = ofd.ShowDialog();
            if (dr == DialogResult.OK)
            {
                //FileInfo fi = new FileInfo(ofd.FileName);
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(ofd.FileName);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];

                int i = 1;

                foreach (ModelItem item in items)
                {

                    i++;
                   
                    xlWorksheet.Cells[i, 1] = item.PropertyCategories.FindPropertyByDisplayName("Medabil", "Etapa").Value.ToDisplayString();
                    xlWorksheet.Cells[i, 2] = item.PropertyCategories.FindPropertyByDisplayName("Medabil", "Numero").Value.ToDisplayString();
                }

                xlWorkbook.Save();
                xlWorkbook.Close();


                
            }




        }

        private void ExecucaoImportDates()
        {
            ExecucaoClear();
            Dictionary<string, Sequence> sequencesExecutados = new Dictionary<string,Sequence>();


            OpenFileDialog ofd = new OpenFileDialog();
            DialogResult dr = ofd.ShowDialog();
            if (dr == DialogResult.OK )
            {
                //FileInfo fi = new FileInfo(ofd.FileName);
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(ofd.FileName);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                Dictionary<int, string> datasExecucao = new Dictionary<int, string>();
                

                for (int i = 2; i <= xlRange.Rows.Count; i++)
                {

                    if (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null)
                    {
                       
                        string etapa = xlRange.Cells[i, 1].Text;
                        string numero = xlRange.Cells[i, 2].Text;
                        string dataExecução = xlRange.Cells[i, 3].Text;
                        int dataExecucaoMS = Convert.ToInt32(xlRange.Cells[i, 3].Value2);

                        if (!datasExecucao.ContainsKey(dataExecucaoMS)) datasExecucao.Add(dataExecucaoMS, dataExecução);


                        Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;

                        Search s = new Search();

                                               

                        s.Selection.SelectAll();
                        SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("Medabil", "Etapa").EqualValue(VariantData.FromDisplayString(etapa));
                        SearchCondition oSearchCondition2 = SearchCondition.HasPropertyByDisplayName("Medabil", "Numero").EqualValue(VariantData.FromDisplayString(numero));
                        s.SearchConditions.Add(oSearchCondition);
                        s.SearchConditions.Add(oSearchCondition2);
                        ModelItem item = s.FindFirst(doc, false);
                        if(item != null) PropertyCreateOrEdit(item, "Medabil", "DataExecucao", dataExecução , "Data de Execução");

                        if(dataExecucaoMS != 0)
                        {

                            string itemSequence = item.PropertyCategories.FindPropertyByDisplayName("Medabil", "Etapa").Value.ToDisplayString();
                            string itemType = item.PropertyCategories.FindPropertyByDisplayName("Medabil", "Tipo").Value.ToDisplayString();
                            double itemWeight = Convert.ToDouble(item.PropertyCategories.FindPropertyByDisplayName("Medabil", "Peso (kg)").Value.ToDisplayString());

                            if (!sequencesExecutados.ContainsKey(itemSequence)) sequencesExecutados.Add(itemSequence, new Sequence());

                            if (!sequencesExecutados[itemSequence].TypesCounter.ContainsKey(itemType) )
                            {
                                sequencesExecutados[itemSequence].TypesCounter.Add(itemType, 1);
                                sequencesExecutados[itemSequence].TypesNetWeight.Add(itemType, itemWeight);

                            }
                            else
                            {
                                sequencesExecutados[itemSequence].TypesCounter[itemType]++;
                                sequencesExecutados[itemSequence].TypesNetWeight[itemType] += itemWeight;

                            }
                        }







                    }


                    
                }

                xlWorkbook.Close();
                SETFolderDelete("Execução");
                VpFolderDelete("Execução");
                ExecucaoSets(datasExecucao);
                ExecucaoVPs(datasExecucao);
                ExecucaoSequencesProcess(sequencesExecutados);

            }
                

            

        }

        private void ExecucaoSets(Dictionary<int, string> datasExecucao)
        {
            if(datasExecucao.ContainsKey(0)) datasExecucao.Remove(0);
            List<int> sortedData = datasExecucao.Keys.ToList().OrderByDescending(x => x).ToList() ;
            


            foreach (int data in sortedData)
            {
                string dataString = datasExecucao[data];

                Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;

                Search s = new Search();

                s.Selection.SelectAll();
                SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("Medabil", "Data de Execução").EqualValue(VariantData.FromDisplayString(dataString));
                s.SearchConditions.Add(oSearchCondition);
                ModelItemCollection items = s.FindAll(doc, false);

                if (items.Count > 0) SETCreateOrEdit("Execução", dataString, items);
            }


            
        }

        private void ExecucaoVPs(Dictionary<int, string> datasExecucao)
        {
            if (datasExecucao.ContainsKey(0)) datasExecucao.Remove(0);
            List<int> sortedData = datasExecucao.Keys.ToList();
            sortedData.Sort();


            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
            var all = doc.Models.CreateCollectionFromRootItems().SelectMany(i => i.DescendantsAndSelf);

            //doc.Models.SetHidden(all, false);
            //doc.ActiveView.LookFromFrontRightTop();

            doc.Models.ResetAllPermanentMaterials();

            doc.Models.ResetAllTemporaryMaterials();
            
            VpCreateOrEdit("Execução", "Inicial");

            
            doc.Models.OverridePermanentColor(all, cinza);
            doc.Models.OverridePermanentTransparency(all, 0.5);

            ModelItemCollection itemsAnteriores = null;

            foreach (int data in sortedData)
            {
                string dataString = datasExecucao[data];

               

                Search s = new Search();

                s.Selection.SelectAll();
                SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("Medabil", "Data de Execução").EqualValue(VariantData.FromDisplayString(dataString));
                s.SearchConditions.Add(oSearchCondition);

                if( itemsAnteriores !=null)
                {
                    doc.Models.OverridePermanentColor(itemsAnteriores, Verde);
                    doc.Models.OverridePermanentTransparency(itemsAnteriores, 0);
                }

                ModelItemCollection items = s.FindAll(doc, false);
                doc.Models.OverridePermanentColor(items, Amarelo);
                doc.Models.OverridePermanentTransparency(items, 0);
                itemsAnteriores = items;

                VpCreateOrEdit("Execução", dataString);




            }
            


        }
        
        /*
        private void CreateOrEditVp(string folderName, string VpName)
        {
            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
            SavedViewpoint vp = new SavedViewpoint(doc.CurrentViewpoint.ToViewpoint());
            vp.DisplayName = VpName;
       
            doc.SavedViewpoints.AddCopy(vp);
          
        }
        */
        

        private void VpCreateOrEdit(string folder, string name)
        {
            var state = ComApiBridge.State;
            InwOpFolderView folderView = null;

            foreach (InwOpSavedView savedview in state.SavedViews())
            {
                if( savedview.Type == nwESavedViewType.eSavedViewType_Folder && savedview.name == folder) folderView = (InwOpFolderView)savedview;

            }

            if(folderView == null)
            {
                folderView = state.ObjectFactory(nwEObjectType.eObjectType_nwOpFolderView);
                folderView.name = folder;
                state.SavedViews().Add(folderView);
            }


            var cv = state.CurrentView.Copy();

            InwOpView vp = state.ObjectFactory(nwEObjectType.eObjectType_nwOpView);

            vp.ApplyHideAttribs = true;
            vp.ApplyMaterialAttribs = true;
            vp.name = name;
            vp.anonview = cv;

            folderView.SavedViews().Add(vp);           

            

        }

        private void VpFolderDelete(string folder)
        {
            var state = ComApiBridge.State;

            if (state.SavedViews().Count > 0)
            {
                InwSavedViewsColl savedViews = state.SavedViews();
                for (int i = 1; i <= state.SavedViews().Count; i++)
                {

                    InwOpSavedView savedview = savedViews[i] as InwOpSavedView;
                    if (savedview.Type == nwESavedViewType.eSavedViewType_Folder && savedview.name == folder)
                    {
                        state.SavedViews().Remove(i);
                        break;
                    }
                }

            }
            

        }

        private void SETFolderDelete(string folderName)
        {
            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
            var ss = doc.SelectionSets;


            var fn = folderName;
            

            try
            {
               

                var fi = ss.Value.IndexOfDisplayName(fn);

                if (fi != -1)
                {
                    ss.Remove(ss.Value[fi]);
                }
               

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ExecucaoSequencesProcess(Dictionary<string, Sequence> sequencesExecucao)
        {
            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;

            foreach ( KeyValuePair<string, Sequence> sequence in sequencesExecucao)
            {
                Search s = new Search();

                s.Selection.SelectAll();
                SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("Medabil", "Nome").EqualValue(VariantData.FromDisplayString(sequence.Key));

                s.SearchConditions.Add(oSearchCondition);
                ModelItem item = s.FindFirst(doc, false);

                foreach(KeyValuePair<string,int> tipo in sequence.Value.TypesCounter)
                {
                    PropertyCreateOrEdit(item, "Medabil", tipo.Key + "QtdExec", tipo.Value.ToString(), tipo.Key + " (Quantidade - Executada)");
                    PropertyCreateOrEdit(item, "Medabil", tipo.Key + "PesoExec", sequence.Value.TypesNetWeight[tipo.Key].ToString() , tipo.Key + " (Peso - kg - Executada)");
                }


            }
        }

        private void SKIDsClear()
        {

            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
         
            Search s = new Search();

            s.Selection.SelectAll();
            SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("Medabil", "SKID");

            s.SearchConditions.Add(oSearchCondition);
            ModelItemCollection items = s.FindAll(doc, false);


            foreach (ModelItem item in items)
            {
                PropertyDelete(item, "Medabil", "SKID");
            }

        }


        private void SKIDsClearMultSkids()
        {

            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;

            Search s = new Search();

            s.Selection.SelectAll();
            SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("Medabil", "SKID_Multiple");

            s.SearchConditions.Add(oSearchCondition);
            ModelItemCollection items = s.FindAll(doc, false);


            foreach (ModelItem item in items)
            {
                PropertyDelete(item, "Medabil", "SKID_Multiple");
            }

        }

        private void ExecucaoClear()
        {

            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;

            Search s = new Search();

            s.Selection.SelectAll();
            SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("Medabil", "Data de Execução");

            s.SearchConditions.Add(oSearchCondition);
            ModelItemCollection items = s.FindAll(doc, false);


            foreach (ModelItem item in items)
            {
                PropertyDelete(item, "Medabil", "DataExecucao");
            }

            ExecucaoSequencesClear();

        }

        private void ExecucaoSequencesClear()
        {

            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;

            Search s = new Search();

            s.Selection.SelectAll();
            SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("Medabil", "Hierarquia").EqualValue(VariantData.FromDisplayString("Etapa")); 

            s.SearchConditions.Add(oSearchCondition);
            ModelItemCollection items = s.FindAll(doc, false);


            foreach (ModelItem item in items)
            {
                ExecucaoPropertyDelete(item);
            }



        }

        private void ExecucaoPropertyDelete(ModelItem item)
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
            



            foreach (InwGUIAttribute2 nwAtt in propn.GUIAttributes())
            {
                if (!nwAtt.UserDefined) continue;
                if (nwAtt.ClassUserName != "Medabil")
                {
                    indexTab++;
                    continue;
                }

                //MessageBox.Show("newAtt => " + nwAtt.ClassName + " ___ " + nwAtt.ClassUserName + " ___ " + nwAtt.name);

                //adiciona as propriedades existentes, já modificando a solicitada
                foreach (InwOaProperty nwProp in nwAtt.Properties())
                {
                    InwOaProperty nwNewProp = state.ObjectFactory(nwEObjectType.eObjectType_nwOaProperty);

                    nwNewProp.UserName = nwProp.UserName;
                    nwNewProp.name = nwProp.name;
                    nwNewProp.value = nwProp.value;
                    //modifica a proprieade existente solicitada
                    if (nwNewProp.UserName.Contains("Executada)")) continue;
                    newPvec.Properties().Add(nwNewProp);
                }

                

               
                propn.SetUserDefined(indexTab, nwAtt.ClassUserName, nwAtt.ClassName, newPvec);


            }
        }

        private void PropertyDelete(ModelItem item, string tabName, string propertyName)
        {
            PropertyCreateOrEdit(item, tabName, propertyName, "_remover");
        }

        private void ExecucaoDateDefine()
        {
            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;

            if(doc.CurrentSelection.SelectedItems.Count == 0)
            {
                MessageBox.Show("Nenhum elemento selecionado");
                return;
            }

           

            Search s = new Search();

            s.Selection.CopyFrom(doc.CurrentSelection );
            SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("Medabil", "Hierarquia").EqualValue(VariantData.FromDisplayString("member")) ;

            s.SearchConditions.Add(oSearchCondition);
            ModelItemCollection items = s.FindAll(doc, false);

            if(items.Count == 0)
            {
                MessageBox.Show("Os itens selecionado não correspondem ao padrão para controle de execução");
                return;
            }


            lastDate = ExecucaoDateForm.Wait(lastDate);

            if (lastDate == null) return;
            
            string date = ((DateTime)lastDate).ToString("dd/MM/yyyy");

            foreach (ModelItem item in items)
            {
                PropertyCreateOrEdit(item, "Medabil", "DataExecucao", date, "Data de Execução");
            }

            

        }

        private void ExecucaoDateDelete()
        {
            DialogResult confirmacao = MessageBox.Show("Confirmar remoção das datas de execução dos elementos selecionados?", "Confirmação", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);

            if (confirmacao == DialogResult.No) return;

            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;

            if (doc.CurrentSelection.SelectedItems.Count == 0)
            {
                MessageBox.Show("Nenhum elemento selecionado");
                return;
            }

            Search s = new Search();

            s.Selection.CopyFrom(doc.CurrentSelection);
            SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("Medabil", "Hierarquia").EqualValue(VariantData.FromDisplayString("member"));

            s.SearchConditions.Add(oSearchCondition);
            ModelItemCollection items = s.FindAll(doc, false);

            if (items.Count == 0)
            {
                MessageBox.Show("Os itens selecionados não correspondem ao padrão para controle de execução");
                return;
            }
                     

            foreach (ModelItem item in items)
            {
                PropertyDelete(item, "Medabil", "DataExecucao");
            }



        }

        private void ExecucaoCalculate()
        {
            Dictionary<int, string> datasExecucao = new Dictionary<int, string>();
            Dictionary<string, Sequence> sequencesExecutados = new Dictionary<string, Sequence>();

            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;

            Search s = new Search();

            s.Selection.SelectAll();
            SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("Medabil", "Data de Execução");

            s.SearchConditions.Add(oSearchCondition);
            ModelItemCollection items = s.FindAll(doc, false);

            foreach(ModelItem item in items)
            {
                string dateString = item.PropertyCategories.FindPropertyByDisplayName("Medabil", "Data de Execução").Value.ToDisplayString();
                
                if (!DateTime.TryParseExact(dateString, "dd/MM/yyyy", null, System.Globalization.DateTimeStyles.None, out DateTime date)) continue;
                int dateInt = int.Parse(date.ToString("yyyyMMdd"));
                if (!datasExecucao.ContainsKey(dateInt)) datasExecucao.Add(dateInt, dateString);


                string itemSequence = item.PropertyCategories.FindPropertyByDisplayName("Medabil", "Etapa").Value.ToDisplayString();
                string itemType = item.PropertyCategories.FindPropertyByDisplayName("Medabil", "Tipo").Value.ToDisplayString();
                double itemWeight = Convert.ToDouble(item.PropertyCategories.FindPropertyByDisplayName("Medabil", "Peso (kg)").Value.ToDisplayString());

                if (!sequencesExecutados.ContainsKey(itemSequence)) sequencesExecutados.Add(itemSequence, new Sequence());

                if (!sequencesExecutados[itemSequence].TypesCounter.ContainsKey(itemType))
                {
                    sequencesExecutados[itemSequence].TypesCounter.Add(itemType, 1);
                    sequencesExecutados[itemSequence].TypesNetWeight.Add(itemType, itemWeight);

                }
                else
                {
                    sequencesExecutados[itemSequence].TypesCounter[itemType]++;
                    sequencesExecutados[itemSequence].TypesNetWeight[itemType] += itemWeight;

                }

            }





            SETFolderDelete("Execução");
            VpFolderDelete("Execução");
            ExecucaoSets(datasExecucao);
            ExecucaoVPs(datasExecucao);
            ExecucaoSequencesClear();
            ExecucaoSequencesProcess(sequencesExecutados);
        }

        private dynamic PropertiesSum(ModelItemCollection elementos, String pesoCategory, string pesoProperty, string tipoCategory, string tipoProperty)
        {
            dynamic somatorio = new ExpandoObject();
            somatorio.peso = Convert.ToDouble(0);
            somatorio.contagem = 0;
            Dictionary<string, dynamic> segregated = new Dictionary<string, dynamic>();
            
            foreach (ModelItem item in elementos)
            {
                DataProperty pesoPropriedade;
                pesoPropriedade = item.PropertyCategories.FindPropertyByDisplayName(pesoCategory, pesoProperty);
                if(pesoPropriedade == null && pesoProperty == "SDS2_Unified") pesoPropriedade = item.PropertyCategories.FindPropertyByDisplayName("SDS2_General", pesoProperty);
                

                double peso = 0;
                if (pesoPropriedade != null && pesoPropriedade.Value.IsDisplayString) peso = Convert.ToDouble(pesoPropriedade.Value.ToDisplayString());
                else if(pesoPropriedade != null) peso = pesoPropriedade.Value.ToDouble(); 


                somatorio.peso += peso;
                somatorio.contagem++;
                string tipo = item.PropertyCategories.FindPropertyByDisplayName(tipoCategory, tipoProperty).Value.ToDisplayString();
                if (!segregated.ContainsKey(tipo))
                {
                    dynamic segregatedData = new ExpandoObject();
                    segregatedData.peso = Convert.ToDouble(0);
                    segregatedData.contagem = 0;
                    segregated.Add(tipo, segregatedData);
                }
                segregated[tipo].contagem++;
                segregated[tipo].peso += peso;
            }

            somatorio.segregated = segregated;
            
            return somatorio;
        }

        private void PropertiesSelectionSum(String pesoCategory, string pesoProperty, string tipoCategory, string tipoProperty)
        {
            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;

            if (doc.CurrentSelection.SelectedItems.Count == 0)
            {
                MessageBox.Show("Nenhum elemento selecionado");
                return;
            }

           

            Search s = new Search();

            s.Selection.CopyFrom(doc.CurrentSelection);
            //SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName(pesoCategory, pesoProperty).EqualValue(VariantData.FromDisplayString("member"));
            
            s.SearchConditions.Add(SearchCondition.HasPropertyByDisplayName(tipoCategory, tipoProperty));

            ModelItemCollection items = s.FindAll(doc, false);



            if (items.Count == 0)
            {
                MessageBox.Show("Nenhum dos elementos selecionados possuí propriedades que possam ser somadas");
                return;
            }

            dynamic retorno = PropertiesSum(items, pesoCategory, pesoProperty, tipoCategory, tipoProperty);


            string mensagem = "";
            
            mensagem += "TOTAIS";
            mensagem += "\nQuantidade de elementos: " + retorno.contagem;
            mensagem += "\nPeso Total (kg): " + Math.Round(retorno.peso, 2);
            foreach(KeyValuePair<string, dynamic> tipo in (Dictionary<string, dynamic>)retorno.segregated)
            {
                mensagem += "\n---------------";
                mensagem += "\n" + tipo.Key;
                mensagem += "\n     Quantidade: " + tipo.Value.contagem ;
                mensagem += "\n     Peso (kg): " + Math.Round(tipo.Value.peso,2);
            }
            MessageBox.Show(mensagem);
            
            
            
        }

        

        private void PropertiesSetsExport(object[,] arrayExport)
        {
            
            OpenFileDialog ofd = new OpenFileDialog();
            DialogResult dr = ofd.ShowDialog();
            if (dr == DialogResult.OK)
            {
                //FileInfo fi = new FileInfo(ofd.FileName);
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(ofd.FileName);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                Excel.Range range;

                range = xlWorksheet.Range[xlWorksheet.Cells[2, 1], xlWorksheet.Cells[xlRange.Rows.Count, xlRange.Columns.Count]];
                range.Clear();

                //range = xlWorksheet.Range[xlWorksheet.Cells[2, 1], xlWorksheet.Cells[arrayExport.GetLength(0) + 1, 1]];
                //range.Cells.NumberFormat = "dd/mm/yyyy";fcdxfdxscfdsfgdsdfscdx

                range = xlWorksheet.Range[xlWorksheet.Cells[2, 1], xlWorksheet.Cells[arrayExport.GetLength(0) + 1, arrayExport.GetLength(1)]];
                range.Cells.Value = arrayExport;

                //range = xlWorksheet.Range[xlWorksheet.Cells[2, 1], xlWorksheet.Cells[arrayExport.GetLength(0) + 1, 1]];
                //range.Cells.NumberFormat = "dd/mm/yyyy";

                xlWorkbook.Save();
                xlWorkbook.Close();
                
            }

        }

        private object[,] ExcelImportToArray() 
        {
            OpenFileDialog ofd = new OpenFileDialog();
            DialogResult dr = ofd.ShowDialog();
            if (dr != DialogResult.OK) return null;
            
            //FileInfo fi = new FileInfo(ofd.FileName);
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(ofd.FileName);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            object[,] myvalues = xlRange.Cells.Value2;

            xlWorkbook.Close();
            return myvalues;



            
        }

    }
}