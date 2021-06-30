using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.ComApi;
using Autodesk.Navisworks.Api.Interop.ComApi;
using Autodesk.Navisworks.Api.Plugins;
using Conexoes;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace MedabilNavisworks
{
    [Plugin("MedabilRibbon", "DLM", DisplayName = "Medabil - DLM")]
    [RibbonLayout("MedabilRibbon.xaml")]
    [RibbonTab("MedabilRibbonTab1", DisplayName = "Medabil")]
    [Command("MedabilButtonLimpar", LargeIcon = @"Resources\BTLIMPAR_32.ico", DisplayName = "Limpar", ToolTip = "Limpa as informação do último processamento de dados")]
    [Command("Mapeia", LargeIcon = @"Resources\BT1_32.ico", DisplayName = "Processar", ToolTip = "Processa as informações dos arquivos anexados e preenche os dados da aba Medabil de propriedades")]
    [Command("ImportaPlanilha", LargeIcon = @"Resources\BT2_32.ico", DisplayName = "Importar Status", ToolTip = "Carregar informações de Status do Painel de Obras")]
    [Command("DefineData", LargeIcon = @"Resources\calendar_32_32.ico", DisplayName = "Definir Data Execução", ToolTip = "Define a data de execução dos elementos selecionados")]
    [Command("RemoveData", LargeIcon = @"Resources\calendarRemove_32_32.ico", DisplayName = "Remove Data Execução", ToolTip = "Remove a data de execução dos elementos selecionados")]
    [Command("SetsViews", LargeIcon = @"Resources\setsVps_32.ico", DisplayName = "Sets e Viewpoints", ToolTip = "Gera os Sets e Viewpoints de forma organizada para os elementos executados")]
    [Command("MedabilButton6", LargeIcon = @"Resources\CalcSelection_32.ico", DisplayName = "Medabil/Tipo", ToolTip = "Apresenta o somatório das propriedades dos elementos selecionados separados por Medabil/Tipo")]
    [Command("MedabilButton7", LargeIcon = @"Resources\CalcSelection_32.ico", DisplayName = "IFC/OBJECTTYPE", ToolTip = "Apresenta o somatório das propriedades dos elementos selecionados separados por IFC/OBJECTTYPE")]
    [Command("PropertiesSetsSum", LargeIcon = @"Resources\excelExport_32.ico", DisplayName = "Exportar", ToolTip = "Exporta os somatórios das propriedades dos elementos dos sets de execução")]
    [Command("SetAtributos", LargeIcon = @"Resources\Estrela_32.ico", DisplayName = "Atributos Custom", ToolTip = "Cria / Edita qualquer tipo de atributo")]
    [Command("Sobre", LargeIcon = @"Resources\projetabim_32.ico", DisplayName = "Medabil", ToolTip = "Sobre")]

    public class Main : CommandHandlerPlugin
    {
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
                case "Mapeia":
                    Mapeia();
                    break;
                case "ImportaPlanilha":
                    ImportaPlanilha();
                    break;
                case "DefineData":
                    DefineData();
                    break;
                case "RemoveData":
                    RemoveData();
                    break;
                case "SetsViews":
                    SetsViews();
                    break;
                case "MedabilButton6":
                    PropertiesSelectionSum(Constantes.Tab, Constantes.PesoDesc, Constantes.Tab, Constantes.Tipo);
                    break;
                case "SetAtributos":
                    this.Propriedade_Custom_Edita();
                    break;
                case "MedabilButton7":
                    PropertiesSelectionSum("SDS2_Unified", "Material_Net_Weight", "IFC", "OBJECTTYPE");
                    break;
                case "PropertiesSetsSum":
                    PropertiesSetsSum();
                    break;
                case "Sobre":
                    Sobre();
                    break;
            }
            //StopProcessMessage();
            return 0;
        }

        private List<Etapa> Etapas = new List<Etapa>();
        public Color cinza { get; set; } = Color.FromByteRGB(171, 171, 171);
        public Color Amarelo { get; set; } = Color.FromByteRGB(255, 255, 0);
        public Color Verde { get; set; } = Color.FromByteRGB(0, 128, 0);
        public DateTime? lastDate { get; set; } = null;
        public ModelItem lastMember { get; set; } = null;
        
        private void Sobre()
        {
            System.Windows.Forms.MessageBox.Show("Medabil 2021 - ₢\nSuporte: Daniel Lins Maciel\ndaniel.maciel@medabil.com.br");
        }


        private void PropertiesSetsSum()
        {

            IList<string> setsFolders = SETFolderList();
            if (setsFolders.Count == 0)
            {
                Utilz.Alerta("Nenhuma pasta de SETs encontrada!");
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


                foreach (ModelItem item in set.Value)
                {
                    if (item.PropertyCategories.FindCategoryByDisplayName(Constantes.Tab) == null)
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
                    arrayExport[i, 1] = item.PropertyCategories.FindPropertyByDisplayName(Constantes.Tab, Constantes.Etapa).Value.ToDisplayString() ?? "NA";
                    arrayExport[i, 2] = item.PropertyCategories.FindPropertyByDisplayName(Constantes.Tab, Constantes.Piecemark).Value.ToDisplayString() ?? "NA";
                    arrayExport[i, 3] = item.PropertyCategories.FindPropertyByDisplayName(Constantes.Tab, Constantes.Numero).Value.ToDisplayString() ?? "NA";
                    arrayExport[i, 4] = item.PropertyCategories.FindPropertyByDisplayName(Constantes.Tab, Constantes.Tipo).Value.ToDisplayString() ?? "NA";
                    arrayExport[i, 5] = Convert.ToDouble(item.PropertyCategories.FindPropertyByDisplayName(Constantes.Tab, Constantes.PesoDesc).Value.ToDisplayString() ?? "0");

                    i++;
                }
            }


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

                range = xlWorksheet.Range[xlWorksheet.Cells[2, 1], xlWorksheet.Cells[arrayExport.GetLength(0) + 1, arrayExport.GetLength(1)]];
                range.Cells.Value = arrayExport;


                xlWorkbook.Save();
                xlWorkbook.Close();

            }

        }
        private void Mapeia()
        {
            Wait w = new Wait(10, "Mapeando peças...");
            w.Show();

            ModelItemCollection items = new ModelItemCollection();

            Document activeDoc = Autodesk.Navisworks.Api.Application.ActiveDocument;


           
            Search s = new Search();
            s.Selection.SelectAll();
            SearchCondition oSearchCondition = SearchCondition.HasCategoryByDisplayName("SDS2_Unified");
            s.SearchConditions.Add(oSearchCondition);
            ModelItemCollection sds = s.FindAll(activeDoc, false);
            w.somaProgresso();
            s.Selection.SelectAll();
            SearchCondition oSearchCondition2 = SearchCondition.HasCategoryByDisplayName("SDS2_General");
            s.SearchConditions.Clear();
            s.SearchConditions.Add(oSearchCondition2);
            ModelItemCollection sds2 = s.FindAll(activeDoc, false);
            w.somaProgresso();

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
            w.somaProgresso();


            //se não acha nada, é porque talvez o arquivo não tenha marcas
            //essa bosta fica toda hora se perdendo com entidades do Tekla
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
            w.somaProgresso();

            var itens = items.GroupBy(x => x).Select(x => x.First()).ToList();
            List<ModelItem> pp = new List<ModelItem>();
            items.Clear();

            var itss = itens.FindAll(x => !x.IsHidden).ToList();
            itens = itens.GroupBy(x => x.GetHashCode()).Select(x => x.First()).ToList();
            itens = itens.OrderBy(x => x.GetHashCode()).ToList();
            foreach (var it in itens)
            {

                ModelItem member, etapa;
                GetMembroPrincipal(it, out member, out etapa);

                if (member != null)
                {
                    pp.Add(member);
                }
            }
            pp = pp.FindAll(x => !x.IsHidden).ToList();
            pp = pp.GroupBy(x => x.GetHashCode()).Select(x => x.First()).ToList();

            w.SetProgresso(1, pp.Count, $"Mapeando etapas nas peças... {pp.Count}");
  
            foreach (ModelItem item in pp)
            {
                if (item.Parent != null)
                {
                    Mapear(item);
                    w.somaProgresso();
                }

            }
            w.SetProgresso(1, 5, "Setando propriedades");
            w.somaProgresso();
            PropertiesSequencesProcess();
            Etapas = new List<Etapa>();
            w.Close();

            Utilz.Alerta("Finalizado.");
        }
        private static ModelItemCollection GetObjetos(ModelItemCollection searchResults, string categoria, string propriedade, List<string> valores)
        {
            if (searchResults.Count == 0)
            {
                return new ModelItemCollection();
            }
            
            //var pcs = searchResults.ToList().FindAll(x=>x.HasGeometry).ToList();
            List<ModelItem> pp = new List<ModelItem>();
            foreach (var p in searchResults)
            {

                if (!p.IsHidden)
                {
                    if (p.Parent != null)
                    {
                        if (!p.Parent.IsHidden)
                        {

                        }
                    }
                }



                ModelItem member = null;

                if (TemPropriedade(categoria, propriedade, valores, p))
                {
                    member = p;
                }


                if (member == null)
                {
                    foreach (var s in p.Descendants)
                    {
                        member = Validar(categoria, propriedade, valores, s);
                        if (member != null)
                        {
                            break;
                        }
                    }
                }





                if (member == null)
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





                if (member != null)
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
                    if (TemPropriedade(categoria, propriedade, valores, nivel1))
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


            ModelItemCollection items = new ModelItemCollection();
            items.AddRange(doc.Models.RootItemDescendantsAndSelf);



            foreach (ModelItem item in items)
            {
                
                Apagar_Propriedade(item, Constantes.Tab, Constantes.Hierarquia);
                Apagar_Propriedade(item, Constantes.Tab, Constantes.Nome);
                Apagar_Propriedade(item, Constantes.Tab, Constantes.Etapa);
                Apagar_Propriedade(item, Constantes.Tab, Constantes.Piecemark);
                Apagar_Propriedade(item, Constantes.Tab, Constantes.Numero);
                Apagar_Propriedade(item, Constantes.Tab, Constantes.Tipo);
                Apagar_Propriedade(item, Constantes.Tab, Constantes.Peso);
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
            
            propriedades.AddRange(pcs_identification);
            string nome_etapa = "";
            string Member_Number_String = "";
            string marca_string = "";
            double peso = 0;
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
                    if (Marca == null)
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




                    if (Sequencia != null && nome_etapa == "") nome_etapa = Sequencia.Value.ToDisplayString();
                    if (Marca != null && marca_string == "") marca_string = Marca.Value.ToDisplayString();
                    if (Peso != null && peso == 0) peso = Peso.Value.ToAnyDouble();
                    if (Tipo != null && Member_Type_String == "") Member_Type_String = Tipo.Value.ToDisplayString();
                    if (Numero != null && Member_Number_String == "") Member_Number_String = Numero.Value.ToDisplayString();


                    //15/04/2020 - para ler os inputs de TecnoMetal
                    if (marca_string.ToUpper().Contains("MARK") && marca_string.ToUpper().Contains("POS"))
                    {
                        var m = marca_string.Split(' ').ToList();
                        marca_string = m[0].ToUpper().Replace(" ", "").Replace("MARK", "").Replace(":", "");
                    }

                    if (marca_string.Contains(" "))
                    {
                        marca_string = item.DisplayName;
                    }
                    else if (marca_string == Member_Type_String)
                    {
                        marca_string = item.DisplayName;
                    }
                }
                catch (Exception ex)
                {

                }

            }

            if (marca_string == "") { return; }

            try
            {
                ModelItem member, etapa;
                GetMembroPrincipal(item, out member, out etapa);
                if (member == null) { return; }
                if (lastMember != null && member == lastMember) return;


                lastMember = member;
                //adiciona a medabil tab na sequence
                var nova_etapa = Etapas.Find(x => x.Nome == nome_etapa);
                if (nova_etapa == null)
                {
                    nova_etapa = new Etapa("sequence", nome_etapa);
                    Etapas.Add(nova_etapa);
                    CriaTabDePropriedades(etapa, Constantes.Tab, Constantes.Tab);
                    Propriedade_Edita_Cria(etapa,Constantes.Tab, Constantes.Hierarquia, Constantes.Etapa, Constantes.Hierarquia);
                    Propriedade_Edita_Cria(etapa,Constantes.Tab, Constantes.Nome, nome_etapa, Constantes.Nome);
                }

                //adiciona a medabil tab nos membros


                Peca newMember = new Peca(
                    "member",
                    Member_Number_String,
                    peso,
                    marca_string,
                    Member_Type_String
                    );
                CriaTabDePropriedades(member,  Constantes.Tab, Constantes.Tab);
                Propriedade_Edita_Cria(member, Constantes.Tab, Constantes.Hierarquia, newMember.TipoObjeto, Constantes.Hierarquia);
                Propriedade_Edita_Cria(member, Constantes.Tab, Constantes.Etapa, nova_etapa.Nome, Constantes.Etapa);
                Propriedade_Edita_Cria(member, Constantes.Tab, Constantes.Piecemark, newMember.Marca, Constantes.Piecemark);
                Propriedade_Edita_Cria(member, Constantes.Tab, Constantes.Numero, newMember.Numero, Constantes.Numero);
                Propriedade_Edita_Cria(member, Constantes.Tab, Constantes.Tipo, newMember.Tipo, Constantes.Tipo);
                Propriedade_Edita_Cria(member, Constantes.Tab, Constantes.Peso, newMember.PesoLiquido.ToString(), Constantes.PesoDesc);

                //if (newMember.Type == "MISCELANEA") return;

                nova_etapa.Pecas.Add(newMember);
                if (!nova_etapa.TypesCounter.ContainsKey(newMember.Tipo))
                {
                    nova_etapa.TypesCounter.Add(newMember.Tipo, 1);
                    nova_etapa.TypesNetWeight.Add(newMember.Tipo, newMember.PesoLiquido);

                }
                else
                {
                    nova_etapa.TypesCounter[newMember.Tipo]++;
                    nova_etapa.TypesNetWeight[newMember.Tipo] += newMember.PesoLiquido;

                }

                nova_etapa.PesoLiquido += newMember.PesoLiquido;
            }
            catch (Exception)
            {

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

            foreach (var etapa in Etapas)
            {
                Search s = new Search();

                s.Selection.SelectAll();
                SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName(Constantes.Tab, Constantes.Nome).EqualValue(VariantData.FromDisplayString(etapa.Nome));

                s.SearchConditions.Add(oSearchCondition);
                ModelItem item = s.FindFirst(doc, false);
                Propriedade_Edita_Cria(item, Constantes.Tab, Constantes.Hierarquia, Constantes.Etapa, Constantes.Hierarquia);
                Propriedade_Edita_Cria(item, Constantes.Tab, Constantes.Nome, etapa.Nome, Constantes.Nome);
                Propriedade_Edita_Cria(item, Constantes.Tab, "MembersCount", etapa.Pecas.Count.ToString(), "Elementos (Quantidade)");
                Propriedade_Edita_Cria(item, Constantes.Tab, "MemebersMass", etapa.PesoLiquido.ToString(), "Elementos (Peso - kg)");
                foreach (KeyValuePair<string, int> typeCount in etapa.TypesCounter)
                {
                    if (typeCount.Key == "MISCELANEA") continue;
                    Propriedade_Edita_Cria(item, Constantes.Tab, typeCount.Key + "Count", typeCount.Value.ToString(), typeCount.Key + " (Quantidade)");
                    Propriedade_Edita_Cria(item, Constantes.Tab, typeCount.Key + Constantes.Peso, etapa.TypesNetWeight[typeCount.Key].ToString(), typeCount.Key + " (Peso - kg)");

                }

            }





        }
        private void CriaTabDePropriedades(ModelItem item, string user_name, string internal_name) //string tipoObjeto, object objeto)
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
                if (nwAtt.UserDefined && nwAtt.ClassUserName == Constantes.Tab) hasTab = true;
            }

            if (!hasTab) propn.SetUserDefined(0, Constantes.Tab, Constantes.Tab, newPvec);
        }
        private void Propriedade_Edita_Cria(ModelItem item, string tabName, string propertyName, string propertyValue, string propertyUserName = "NewUserProperty")
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

                    nwNewProp.UserName = propertyUserName;
                    nwNewProp.name = propertyName;
                    nwNewProp.value = propertyValue;
                    newPvec.Properties().Add(nwNewProp);
                }

                newPvec.Properties().Sort();
                propn.SetUserDefined(indexTab, nwAtt.ClassUserName, nwAtt.ClassName, newPvec);
            }
        }
        private void ImportaPlanilha()
        {
       

            string arquivo = Conexoes.Utilz.Abrir_String("xlsx","Selecione o arquivo de report");
            if(!File.Exists(arquivo))
            {
                return;
            }

            Dictionary<string, string> status = new Dictionary<string, string>();
            ModelItemCollection status_multiplo = new ModelItemCollection();
            List<string> lista_status = new List<string>();

            try
            {
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(arquivo);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;
                object[,] retorno = xlRange.Cells.Value2;
                xlWorkbook.Close();

                var linhas = retorno.GetLength(0);

                for (int i = 1; i <= linhas; i++)
                {
                    object pmObj = retorno[i, 12];
                    object statusObj = retorno[i, 3];

                    if (pmObj == null || statusObj == null) continue;
                    string pm = pmObj.ToString();
                    string st = statusObj.ToString();

                    if (pm.Replace(" ", "") == "" || st.Replace(" ", "") == "") continue;

                    if (!lista_status.Contains(st)) lista_status.Add(st);

                    if (!status.ContainsKey(pm)) status.Add(pm, st);
                    else if (!status[pm].Contains(st)) status[pm] = status[pm] + ", " + st;

                }
            }
            catch (Exception ex)
            {
                Conexoes.Utilz.Alerta(ex.Message + "\n" + ex.StackTrace);
                return;
            }



            var atuais = GetPecas(Constantes.Tab, Constantes.Status);
            if(atuais.Count>0)
            {
                if (Conexoes.Utilz.Pergunta($"Deseja limpar o romaneio atual das {atuais.Count} peças existentes?"))
                {
                    Apagar_Propriedade(Constantes.Tab, Constantes.Status,atuais);
                }
            }









            Dictionary<string, string> status_sequencia = new Dictionary<string, string>();


            foreach (KeyValuePair<string, string> st in status)
            {

                string marca = st.Key;
                ModelItemCollection items = GetPecas(marca);

                foreach (ModelItem item in items)
                {
                    //grava o romaneio na peça
                    Propriedade_Edita_Cria(item, Constantes.Tab, Constantes.Status, st.Value, Constantes.Status);
                    //insere o status na lista
                    string etapa = GetPropriedade(item,Constantes.Tab,Constantes.Etapa).Value.ToDisplayString();
                    string sk = st.Value;
                    if (!status_sequencia.ContainsKey(etapa)) status_sequencia.Add(etapa, sk);
                    else if (!status_sequencia[etapa].Contains(sk)) status_sequencia[etapa] = status_sequencia[etapa] + ", " + sk;
                }
            }

            SetStatusNivelEtapa(status_sequencia);
            CriaEditaStatus(lista_status);




        }

        private static DataProperty GetPropriedade(ModelItem item, string tab, string propriedade)
        {
            var s = item.PropertyCategories.FindPropertyByDisplayName(tab, propriedade);
            if (s == null)
            {
                s = new DataProperty(propriedade, propriedade, new VariantData(""));
            }
            return s;
        }
        public static List<DataProperty> GetPropriedades(ModelItem item)
        {
            List<DataProperty> retorno = new List<DataProperty>();
            var s =item.PropertyCategories.ToList();
            foreach(var t in s)
            {
                retorno.AddRange(t.Properties.ToList());
            }
            return retorno;
        }


        private void Apagar_Propriedade(string tab, string propriedade, ModelItemCollection pcs = null)
        {
            if (pcs == null)
            {
                pcs = GetPecas(tab, propriedade);
            }


            Conexoes.Wait w = new Wait(pcs.Count,$"Apagando propriedade {tab} - {propriedade} de {pcs.Count} Peças");
            w.Show();

            foreach (ModelItem item in pcs)
            {
                w.somaProgresso();
                Apagar_Propriedade(item, tab,propriedade);
            }
            w.Close();
        }

        private static ModelItemCollection GetPecas(string marca)
        {
            Search s = new Search();
            s.Selection.SelectAll();

            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
            SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName(Constantes.Tab, Constantes.Piecemark).EqualValue(VariantData.FromDisplayString(marca));
            s.SearchConditions.Add(oSearchCondition);
            ModelItemCollection items = s.FindAll(doc, false);
            return items;
        }

        private void SetStatusNivelEtapa(Dictionary<string, string> sequencia_status)
        {
            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
            foreach (KeyValuePair<string, string> sq in sequencia_status)
            {

                Search s = new Search();

                s.Selection.SelectAll();
                SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName(Constantes.Tab, Constantes.Nome).EqualValue(VariantData.FromDisplayString(sq.Key));
                s.SearchConditions.Add(oSearchCondition);
                SearchCondition oSearchCondition2 = SearchCondition.HasPropertyByDisplayName(Constantes.Tab, Constantes.Hierarquia).EqualValue(VariantData.FromDisplayString(Constantes.Etapa));
                s.SearchConditions.Add(oSearchCondition2);

                ModelItem Nivel_Etapa = s.FindFirst(doc, false);

                string statusTexto = string.Format("{0} "+ Constantes.Status + " {1}",
                    sq.Value.Split(","[0]).Length.ToString(),
                    sq.Value
                    );
                Propriedade_Edita_Cria(Nivel_Etapa, Constantes.Tab, Constantes.Status, statusTexto, Constantes.Status);

            }
        }
        private void CriaEditaStatus(List<string> lista_status)
        {
            SETFolderDelete(Constantes.Status);
            lista_status.OrderByDescending(x => x);
            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
            foreach (string status in lista_status)
            {

                Search s = new Search();

                s.Selection.SelectAll();
                SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName(Constantes.Tab, Constantes.Status).EqualValue(VariantData.FromDisplayString(status));
                s.SearchConditions.Add(oSearchCondition);
                ModelItemCollection items = s.FindAll(doc, false);

                if (items.Count > 0) SetarCriarEditar(Constantes.Status, status, items);
            }




        }

        private void SetarCriarEditar(string folderName, string setName, ModelItemCollection items)
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
            foreach (SavedItem sv in fo.Children)
            {
                SelectionSet svSet = sv as SelectionSet;
                setsList.Add(svSet.DisplayName, SetDeepLook(svSet));
            }

            return setsList;

        }
        private ModelItemCollection SetDeepLook(SelectionSet Set)
        {
            ModelItemCollection items = new ModelItemCollection();
            foreach (ModelItem item in Set.GetSelectedItems())
            {
                DataProperty hierarquia = item.PropertyCategories.FindPropertyByDisplayName(Constantes.Tab, Constantes.Hierarquia);
                try
                {

                    if (hierarquia == null)
                    {
                        IEnumerable<ModelItem> membrosPais = from x in item.Ancestors where x.PropertyCategories.FindPropertyByDisplayName(Constantes.Tab, Constantes.Hierarquia) != null && x.PropertyCategories.FindPropertyByDisplayName("Medabil", "Hierarquia").Value.ToDisplayString() == "member" select x;
                        items.AddRange(membrosPais);

                        IEnumerable<ModelItem> membrosFilhos = from x in item.Descendants where x.PropertyCategories.FindPropertyByDisplayName(Constantes.Tab, Constantes.Hierarquia) != null && x.PropertyCategories.FindPropertyByDisplayName("Medabil", "Hierarquia").Value.ToDisplayString() == "member" select x;
                        items.AddRange(membrosFilhos);


                    }
                    else if (hierarquia.Value.ToDisplayString() == "etapa")
                    {
                        IEnumerable<ModelItem> membros = from x in item.Descendants where x.PropertyCategories.FindPropertyByDisplayName(Constantes.Tab, Constantes.Hierarquia) != null select x;
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

            var ss = doc.SelectionSets;
            foreach (SavedItem set in doc.SelectionSets.Value)
            {
                if (set.IsGroup) setFoldesList.Add(set.DisplayName);
            }

            return setFoldesList;

        }
        private void ExecucaoSets(Dictionary<int, string> datasExecucao)
        {
            if (datasExecucao.ContainsKey(0)) datasExecucao.Remove(0);
            List<int> sortedData = datasExecucao.Keys.ToList().OrderByDescending(x => x).ToList();



            foreach (int data in sortedData)
            {
                string dataString = datasExecucao[data];

                Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;

                Search s = new Search();

                s.Selection.SelectAll();
                SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName(Constantes.Tab, Constantes.DataExecucaoDesc).EqualValue(VariantData.FromDisplayString(dataString));
                s.SearchConditions.Add(oSearchCondition);
                ModelItemCollection items = s.FindAll(doc, false);

                if (items.Count > 0) SetarCriarEditar("Execução", dataString, items);
            }



        }
        private void ExecucaoVPs(Dictionary<int, string> datasExecucao)
        {
            try
            {
                if (datasExecucao.ContainsKey(0)) datasExecucao.Remove(0);
                List<int> sortedData = datasExecucao.Keys.ToList();
                sortedData.Sort();


                Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
                var all = doc.Models.CreateCollectionFromRootItems().SelectMany(i => i.DescendantsAndSelf);


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
                    SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName(Constantes.Tab, Constantes.DataExecucaoDesc).EqualValue(VariantData.FromDisplayString(dataString));
                    s.SearchConditions.Add(oSearchCondition);

                    if (itemsAnteriores != null)
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
            catch (Exception)
            {
            }
           
        }
        private void VpCreateOrEdit(string folder, string name)
        {
            var state = ComApiBridge.State;
            InwOpFolderView folderView = null;

            foreach (InwOpSavedView savedview in state.SavedViews())
            {
                if (savedview.Type == nwESavedViewType.eSavedViewType_Folder && savedview.name == folder) folderView = (InwOpFolderView)savedview;

            }

            if (folderView == null)
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

        private void LimpaStatus()
        {
           

        }

        private static ModelItemCollection GetPecas(string tab, string propriedade)
        {
            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;

            Search s = new Search();

            s.Selection.SelectAll();
            SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName(tab,propriedade);

            s.SearchConditions.Add(oSearchCondition);
            ModelItemCollection items = s.FindAll(doc, false);
            return items;
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
                if (nwAtt.ClassUserName != Constantes.Tab)
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
                    if (nwNewProp.UserName.Contains("Executada)")) continue;
                    newPvec.Properties().Add(nwNewProp);
                }




                propn.SetUserDefined(indexTab, nwAtt.ClassUserName, nwAtt.ClassName, newPvec);


            }
        }
        private void Apagar_Propriedade(ModelItem item, string tabName, string propertyName)
        {
            Propriedade_Edita_Cria(item, tabName, propertyName, "_remover");
        }
        private void Propriedade_Custom_Edita()
        {
            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;

            if (doc.CurrentSelection.SelectedItems.Count == 0)
            {
                Utilz.Alerta("Nenhum elemento selecionado");
                return;
            }



            Search s = new Search();

            s.Selection.CopyFrom(doc.CurrentSelection);
            SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName(Constantes.Tab, Constantes.Hierarquia).EqualValue(VariantData.FromDisplayString("member"));

            s.SearchConditions.Add(oSearchCondition);
            ModelItemCollection items = s.FindAll(doc, false);

            if (items.Count == 0)
            {
                Utilz.Alerta("Os itens selecionado não correspondem ao padrão para controle de execução");
                return;
            }



            var mm = new Menus.SetarAtributo();
            mm.ShowDialog();


            if(mm.DialogResult== DialogResult.OK)
            {
                Conexoes.Wait w = new Wait(items.Count, $"Setando atributo [{mm.txt_propriedade.Text}] = [{mm.txt_valor.Text}] em  {items.Count} itens");
                w.Show();
                foreach (ModelItem item in items)
                {
                    Propriedade_Edita_Cria(item, mm.txt_grupo.Text, Utilz.RemoverCaracteresEspeciais(mm.txt_propriedade.Text),mm.txt_valor.Text, mm.txt_propriedade.Text);
                    w.somaProgresso();
                }
                w.Close();
                Conexoes.Utilz.Alerta("Finalizado", "", System.Windows.MessageBoxImage.Information);
            }





        }
        private void DefineData()
        {
            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;

            if (doc.CurrentSelection.SelectedItems.Count == 0)
            {
                Utilz.Alerta("Nenhum elemento selecionado");
                return;
            }



            Search s = new Search();

            s.Selection.CopyFrom(doc.CurrentSelection);
            SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName(Constantes.Tab, Constantes.Hierarquia).EqualValue(VariantData.FromDisplayString("member"));

            s.SearchConditions.Add(oSearchCondition);
            ModelItemCollection items = s.FindAll(doc, false);

            if (items.Count == 0)
            {
                Utilz.Alerta("Os itens selecionado não correspondem ao padrão para controle de execução");
                return;
            }


            lastDate = ExecucaoDateForm.PedirData(lastDate);

            if (lastDate == null) return;

            string date = ((DateTime)lastDate).ToString("dd/MM/yyyy");

            foreach (ModelItem item in items)
            {
                Propriedade_Edita_Cria(item, Constantes.Tab, Constantes.DataExecucao, date, Constantes.DataExecucaoDesc);
            }



        }
        private void RemoveData()
        {
            var confirmacao = Utilz.Pergunta("Confirmar remoção das datas de execução dos elementos selecionados?", "Confirmação");

            if (!confirmacao) return;

            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;

            if (doc.CurrentSelection.SelectedItems.Count == 0)
            {
                Utilz.Alerta("Nenhum elemento selecionado");
                return;
            }

            Search s = new Search();

            s.Selection.CopyFrom(doc.CurrentSelection);
            SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName(Constantes.Tab, Constantes.Hierarquia).EqualValue(VariantData.FromDisplayString("member"));

            s.SearchConditions.Add(oSearchCondition);
            ModelItemCollection items = s.FindAll(doc, false);

            if (items.Count == 0)
            {
                Utilz.Alerta("Os itens selecionados não correspondem ao padrão para controle de execução");
                return;
            }


            foreach (ModelItem item in items)
            {
                Apagar_Propriedade(item, Constantes.Tab, Constantes.DataExecucao);
            }



        }
        private void SetsViews()
        {
            Dictionary<int, string> datasExecucao = new Dictionary<int, string>();
            Dictionary<string, Etapa> sequencesExecutados = new Dictionary<string, Etapa>();

            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;

            Search s = new Search();
            s.Selection.SelectAll();
            SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName(Constantes.Tab, Constantes.DataExecucaoDesc);

            s.SearchConditions.Add(oSearchCondition);
            ModelItemCollection items = s.FindAll(doc, false);

            foreach (ModelItem item in items)
            {
                string dateString = item.PropertyCategories.FindPropertyByDisplayName(Constantes.Tab, Constantes.DataExecucaoDesc).Value.ToDisplayString();

                if (!DateTime.TryParseExact(dateString, "dd/MM/yyyy", null, System.Globalization.DateTimeStyles.None, out DateTime date)) continue;
                int dateInt = int.Parse(date.ToString("yyyyMMdd"));
                if (!datasExecucao.ContainsKey(dateInt)) datasExecucao.Add(dateInt, dateString);


                string nome_etapa = item.PropertyCategories.FindPropertyByDisplayName(Constantes.Tab, Constantes.Etapa).Value.ToDisplayString();
                string tipo = item.PropertyCategories.FindPropertyByDisplayName(Constantes.Tab, Constantes.Tipo).Value.ToDisplayString();
                double peso = Convert.ToDouble(item.PropertyCategories.FindPropertyByDisplayName(Constantes.Tab, Constantes.PesoDesc).Value.ToDisplayString());

                if (!sequencesExecutados.ContainsKey(nome_etapa)) sequencesExecutados.Add(nome_etapa, new Etapa());

                if (!sequencesExecutados[nome_etapa].TypesCounter.ContainsKey(tipo))
                {
                    sequencesExecutados[nome_etapa].TypesCounter.Add(tipo, 1);
                    sequencesExecutados[nome_etapa].TypesNetWeight.Add(tipo, peso);
                }
                else
                {
                    sequencesExecutados[nome_etapa].TypesCounter[tipo]++;
                    sequencesExecutados[nome_etapa].TypesNetWeight[tipo] += peso;
                }

            }





            SETFolderDelete("Execução");
            VpFolderDelete("Execução");
            ExecucaoSets(datasExecucao);
            ExecucaoVPs(datasExecucao);




            s = new Search();
            s.Selection.SelectAll();
            oSearchCondition = SearchCondition.HasPropertyByDisplayName(Constantes.Tab, Constantes.Hierarquia).EqualValue(VariantData.FromDisplayString(Constantes.Etapa));
            s.SearchConditions.Add(oSearchCondition);
            items = s.FindAll(doc, false);

            foreach (ModelItem item in items)
            {
                ExecucaoPropertyDelete(item);
            }






            //seta o peso total na raiz do modelo
            foreach (var etapa in sequencesExecutados.Select(x => x.Value).ToList())
            {
                s = new Search();

                s.Selection.SelectAll();
                oSearchCondition = SearchCondition.HasPropertyByDisplayName(Constantes.Tab, Constantes.Nome).EqualValue(VariantData.FromDisplayString(etapa.Nome));

                s.SearchConditions.Add(oSearchCondition);
                ModelItem item = s.FindFirst(doc, false);

                foreach (KeyValuePair<string, int> tipo in etapa.TypesCounter)
                {
                    Propriedade_Edita_Cria(item, Constantes.Tab, tipo.Key + Constantes.QtdExec, tipo.Value.ToString(), tipo.Key + " (Quantidade - Executada)");
                    Propriedade_Edita_Cria(item, Constantes.Tab, tipo.Key + Constantes.PesoExec, etapa.TypesNetWeight[tipo.Key].ToString(), tipo.Key + " (Peso - kg - Executada)");
                }


            }


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
                if (pesoPropriedade == null && pesoProperty == "SDS2_Unified") pesoPropriedade = item.PropertyCategories.FindPropertyByDisplayName("SDS2_General", pesoProperty);


                double peso = 0;
                if (pesoPropriedade != null && pesoPropriedade.Value.IsDisplayString) peso = Convert.ToDouble(pesoPropriedade.Value.ToDisplayString());
                else if (pesoPropriedade != null) peso = pesoPropriedade.Value.ToDouble();


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
                Utilz.Alerta("Nenhum elemento selecionado");
                return;
            }



            Search s = new Search();

            s.Selection.CopyFrom(doc.CurrentSelection);
            //SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName(pesoCategory, pesoProperty).EqualValue(VariantData.FromDisplayString("member"));

            s.SearchConditions.Add(SearchCondition.HasPropertyByDisplayName(tipoCategory, tipoProperty));

            ModelItemCollection items = s.FindAll(doc, false);



            if (items.Count == 0)
            {
                Utilz.Alerta("Nenhum dos elementos selecionados possuí propriedades que possam ser somadas");
                return;
            }

            dynamic retorno = PropertiesSum(items, pesoCategory, pesoProperty, tipoCategory, tipoProperty);


            string mensagem = "";

            mensagem += "TOTAIS";
            mensagem += "\nQuantidade de elementos: " + retorno.contagem;
            mensagem += "\nPeso Total (kg): " + Math.Round(retorno.peso, 2);
            foreach (KeyValuePair<string, dynamic> tipo in (Dictionary<string, dynamic>)retorno.segregated)
            {
                mensagem += "\n---------------";
                mensagem += "\n" + tipo.Key;
                mensagem += "\n     Quantidade: " + tipo.Value.contagem;
                mensagem += "\n     Peso (kg): " + Math.Round(tipo.Value.peso, 2);
            }
            Conexoes.Utilz.Alerta(mensagem,"Finalizado", System.Windows.MessageBoxImage.Information);



        }
    }
}