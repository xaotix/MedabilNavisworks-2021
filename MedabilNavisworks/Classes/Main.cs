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
    [Command("ImportaPlanilhaCustom", LargeIcon = @"Resources\BT2_32.ico", DisplayName = "Importar planilha custom", ToolTip = "Carregar informações de Status de uma planilha externa")]
    [Command("SetsViews", LargeIcon = @"Resources\setsVps_32.ico", DisplayName = "Sets e Viewpoints", ToolTip = "Gera os Sets e Viewpoints de forma organizada para os elementos executados")]
    [Command("ApagaAtributo", LargeIcon = @"Resources\BTLIMPAR_32.ico", DisplayName = "Apaga atributo selecionado", ToolTip = "Apaga atributo selecionado")]
    [Command("MedabilButton6", LargeIcon = @"Resources\CalcSelection_32.ico", DisplayName = "Medabil/Tipo", ToolTip = "Apresenta o somatório das propriedades dos elementos selecionados separados por Medabil/Tipo")]
    [Command("MedabilButton7", LargeIcon = @"Resources\CalcSelection_32.ico", DisplayName = "IFC/OBJECTTYPE", ToolTip = "Apresenta o somatório das propriedades dos elementos selecionados separados por IFC/OBJECTTYPE")]
    [Command("PropertiesSetsSum", LargeIcon = @"Resources\excelExport_32.ico", DisplayName = "Exportar Somatórios", ToolTip = "Exporta os somatórios das propriedades dos elementos dos sets de execução")]
    [Command("ExtrairAtributos", LargeIcon = @"Resources\excelExport_32.ico", DisplayName = "Exportar Atributos", ToolTip = "Exporta os atributos dos elementos")]
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
                    Funcoes.Limpar();
                    break;
                case "Mapeia":
                    Mapeia();
                    break;
                case "ImportaPlanilha":
                    ImportaPlanilha();
                    break;
                case "ImportaPlanilhaCustom":
                    ImportaPlanilhaCustom();
                    break;
                case "DefineData":
                    DefineData();
                    break;
                case "RemoveData":
                    RemoveData();
                    break;
                case "ExtrairAtributos":
                    ExtrairAtributos();
                    break;
                case "ApagaAtributo":
                    ApagaAtributo();
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

       

        public void ExtrairAtributos()
        {
            var destino = Conexoes.Utilz.SalvarArquivo("csv");

            if (destino != null && destino != "")
            {
              
                var lista = Funcoes.GetPecas();
              
                Conexoes.Wait w = new Wait(lista.Count +2, $"Lendo atributos das Pcs...{lista.Count}");
                w.Show();


                if (lista.Count > 0)
                {
                    w.Visibility = System.Windows.Visibility.Collapsed;
                    var colunas = Conexoes.Utilz.SelecionarObjetos(lista.SelectMany(x=>x.GetTabs()).Distinct().ToList());
                    w.Visibility = System.Windows.Visibility.Visible;
                    if (colunas.Count == 0) {
                        w.Close();
                        return; }




                    DB.Tabela tb = new DB.Tabela();
                    foreach (var p in lista)
                    {
                        w.somaProgresso();
                        DB.Linha n = new DB.Linha();
                        n.Tabela = p.Marca;
                        n.Add("Marca", p.Marca);
                        foreach (var att in p.GetPropriedades(colunas))
                        {
                            n.Celulas.Add(att);
                        }
                        n.Celulas = n.Celulas.OrderBy(x => x.Coluna).ToList();
                        n.Celulas.Insert(0, new DB.Celula("Marca", p.Marca));
                        n.Celulas.Insert(1, new DB.Celula("Tipo", p.Tipo));

                        tb.Linhas.Add(n);
                    }
                    w.Close();
                    Conexoes.Utilz.Arquivo.Gravar(destino, tb.GetLista().Select(x => string.Join(";", x)).ToList());

                    if (File.Exists(destino))
                    {
                        Conexoes.Utilz.Abrir(destino);
                    }

                }
            }
        }

        private void Mapeia()
        {
            int max = 4;
         


            Document activeDoc = Autodesk.Navisworks.Api.Application.ActiveDocument;
            var pcs = Funcoes.GetPecas();

            Wait w = new Wait(10, $"1/{max} Procurando peças...");
            w.Show();
            w.somaProgresso();
            



            w.SetProgresso(1, pcs.Count+2, $"3/{max} Setando atributos...{pcs.Count} itens");
            w.somaProgresso();
            w.somaProgresso();
            foreach(var peca in pcs)
            {
                try
                {
                   

                    string nome_etapa = peca.GetEtapa();
                    var etapa = peca.ModelItem.Parent;
                    
                    //adiciona a medabil tab na sequence
                    var nova_etapa = Etapas.Find(x => x.Nome == nome_etapa);
                    if (nova_etapa == null)
                    {
                        nova_etapa = new Etapa("sequence", nome_etapa, etapa);
                        Etapas.Add(nova_etapa);
                    }

                    //adiciona a medabil tab nos membros
                    Funcoes.CriaTabDePropriedades(peca.ModelItem, Constantes.Tab);
                    Funcoes.Propriedade_Edita_Cria(peca.ModelItem, Constantes.Tab, Constantes.Hierarquia, peca.TipoObjeto);
                    Funcoes.Propriedade_Edita_Cria(peca.ModelItem, Constantes.Tab, Constantes.Etapa, nova_etapa.Nome);
                    Funcoes.Propriedade_Edita_Cria(peca.ModelItem, Constantes.Tab, Constantes.Piecemark, peca.Marca);
                    Funcoes.Propriedade_Edita_Cria(peca.ModelItem, Constantes.Tab, Constantes.Numero, peca.GetNumero());
                    Funcoes.Propriedade_Edita_Cria(peca.ModelItem, Constantes.Tab, Constantes.Tipo, peca.Tipo);
                    Funcoes.Propriedade_Edita_Cria(peca.ModelItem, Constantes.Tab, Constantes.Descricao, peca.GetDescricao());
                    Funcoes.Propriedade_Edita_Cria(peca.ModelItem, Constantes.Tab, Constantes.Peso, peca.GetPesoLiquido().ToString());

                    /*novos*/
                    Funcoes.Propriedade_Edita_Cria(peca.ModelItem, Constantes.Tab, Constantes.Area, peca.GetAtributo("Area"));
                    Funcoes.Propriedade_Edita_Cria(peca.ModelItem, Constantes.Tab, Constantes.Volume, peca.GetAtributo("Volume"));
                    Funcoes.Propriedade_Edita_Cria(peca.ModelItem, Constantes.Tab, Constantes.Comprimento,peca.GetComprimento().ToString() );
                    Funcoes.Propriedade_Edita_Cria(peca.ModelItem, Constantes.Tab, Constantes.Largura,peca.GetLargura().ToString() );
                    Funcoes.Propriedade_Edita_Cria(peca.ModelItem, Constantes.Tab, Constantes.Espessura,peca.GetEspessura().ToString() );
                    Funcoes.Propriedade_Edita_Cria(peca.ModelItem, Constantes.Tab, "Altura Bruta", peca.GetAtributo("Altura Bruta"));
                    Funcoes.Propriedade_Edita_Cria(peca.ModelItem, Constantes.Tab, "Largura Bruta", peca.GetAtributo("Largura Bruta"));
                    Funcoes.Propriedade_Edita_Cria(peca.ModelItem, Constantes.Tab, "ID", peca.GetAtributo("Global ID"));

                    var pesopc = peca.GetPesoLiquido();

                    nova_etapa.Pecas.Add(peca);
                    if (!nova_etapa.TypesCounter.ContainsKey(peca.Tipo))
                    {
                        nova_etapa.TypesCounter.Add(peca.Tipo, 1);
                        nova_etapa.TypesNetWeight.Add(peca.Tipo, pesopc);
                    }
                    else
                    {
                        nova_etapa.TypesCounter[peca.Tipo]++;
                        nova_etapa.TypesNetWeight[peca.Tipo] += pesopc;
                    }

                    nova_etapa.PesoLiquido += pesopc;
                }
                catch (Exception)
                {
                }
                w.somaProgresso();
            }


            w.SetProgresso(1, Etapas.Count, $"4/{max} Setando propriedades Etapas {Etapas.Count}");

            foreach (var etapa in Etapas)
            {
                //var s = new Search();
                //s.Selection.SelectAll();
                //s.SearchConditions.Add(SearchCondition.HasPropertyByDisplayName(Constantes.Tab, Constantes.Nome).EqualValue(VariantData.FromDisplayString(etapa.Nome)));
                //ModelItem item = s.FindFirst(doc, false);
                Funcoes.CriaTabDePropriedades(etapa.modelItem, Constantes.Tab);
                Funcoes.Propriedade_Edita_Cria(etapa.modelItem, Constantes.Tab, Constantes.Hierarquia, Constantes.Etapa);
                Funcoes.Propriedade_Edita_Cria(etapa.modelItem, Constantes.Tab, Constantes.Nome, etapa.Nome);
                Funcoes.Propriedade_Edita_Cria(etapa.modelItem, Constantes.Tab, "MembersCount", etapa.Pecas.Count.ToString());
                Funcoes.Propriedade_Edita_Cria(etapa.modelItem, Constantes.Tab, "MemebersMass", etapa.PesoLiquido.ToString());
                foreach (KeyValuePair<string, int> typeCount in etapa.TypesCounter)
                {
                    if (typeCount.Key != "MISCELANEA")
                    {
                        Funcoes.Propriedade_Edita_Cria(etapa.modelItem, Constantes.Tab, typeCount.Key + "Count", typeCount.Value.ToString());
                        Funcoes.Propriedade_Edita_Cria(etapa.modelItem, Constantes.Tab, typeCount.Key + Constantes.Peso, etapa.TypesNetWeight[typeCount.Key].ToString());
                    }

                }
                w.somaProgresso();
            }
            Etapas = new List<Etapa>();
            w.Close();

            Utilz.Alerta("Finalizado.", "", System.Windows.MessageBoxImage.Information);
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







        private void ImportaPlanilha()
        {


            string arquivo = Conexoes.Utilz.Abrir_String("xlsx", "Selecione o arquivo de report");
            if (!File.Exists(arquivo))
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



            var atuais = Funcoes.GetPecas();
            if (atuais.Count > 0)
            {
                if (Conexoes.Utilz.Pergunta($"Deseja limpar o romaneio atual das {atuais.Count} peças existentes?"))
                {
                    Funcoes.Apagar_Propriedade(Constantes.Tab, Constantes.Status, atuais);
                }
            }









            Dictionary<string, string> status_sequencia = new Dictionary<string, string>();


            foreach (KeyValuePair<string, string> st in status)
            {

                string marca = st.Key;
                ModelItemCollection items = Funcoes.GetPecasPorMarca(marca);

                foreach (ModelItem item in items)
                {
                    //grava o romaneio na peça
                    Funcoes.Propriedade_Edita_Cria(item, Constantes.Tab, Constantes.Status, st.Value);
                    //insere o status na lista
                    string etapa = Funcoes.GetPropriedade(item, Constantes.Tab, Constantes.Etapa).Value.ToDisplayString();
                    string sk = st.Value;
                    if (!status_sequencia.ContainsKey(etapa)) status_sequencia.Add(etapa, sk);
                    else if (!status_sequencia[etapa].Contains(sk)) status_sequencia[etapa] = status_sequencia[etapa] + ", " + sk;
                }
            }

            SetStatusNivelEtapa(status_sequencia);
            CriaEditaStatus(lista_status);




        }


        private void ImportaPlanilhaCustom()
        {


            string arquivo = Conexoes.Utilz.Abrir_String("xlsx", "Selecione o arquivo de report");
            if (!File.Exists(arquivo))
            {
                return;
            }



            Dictionary<string, string> status = new Dictionary<string, string>();
            ModelItemCollection status_multiplo = new ModelItemCollection();
            List<string> lista_status = new List<string>();

            try
            {

             var tbs =   Conexoes.Utilz.GetTabelas(arquivo);
                if(tbs.Count==0)
                {
                    return;
                }

               DB.Tabela sel = tbs[0];

                if(tbs.Count>1)
                {
                    sel = Conexoes.Utilz.SelecionarObjeto(tbs, null, "Selecione a aba do excel");
                }


                if (sel == null) { return; }

                var coluna_marca = Conexoes.Utilz.SelecionarObjeto(sel.GetColunas(), null, "Selecione a coluna que corresponde ao nome da marca");
                List<Report> retorno = new List<Report>();

                if (coluna_marca!=null)
                {
                    var colunas_importar = Conexoes.Utilz.SelecionarObjetos(sel.GetColunas().FindAll(x => x != coluna_marca), true, "Selecione as colunas que deseja importar os atributos");

                    if(colunas_importar.Count>0)
                    {
                        var linhas = sel.Linhas.GroupBy(x => x.Get(coluna_marca).valor).ToList();
                       if(linhas.Count>0)
                        {
                            var pecas_modelo = Funcoes.GetPecas();
                            if(pecas_modelo.Count==0)
                            {
                                Conexoes.Utilz.Alerta("Nenhuma peça encontrada no modelo.");
                                return;
                            }
                            Conexoes.Wait w = new Wait(pecas_modelo.Count, "Setando valores...");
                            w.Show();

                            foreach(var pc in pecas_modelo)
                            {
                                var igual = linhas.Find(x => x.Key.ToUpper().Replace(" ", "_") == pc.Marca.ToUpper().Replace(" ", "_"));

                                if(igual!=null)
                                {
                                   foreach(var col in colunas_importar)
                                    {
                                        var valores = string.Join(",",igual.Select(x => x.Get(col).valor).Distinct().ToList());

                                        if(valores!="")
                                        {
                                            Funcoes.Propriedade_Edita_Cria(pc.ModelItem, Constantes.Tab, col, valores);
                                        }
                                    }
                                }
                                else
                                {
                                    retorno.Add(new Report("Nenhum atributo encontrado para ser setado.", pc.Marca ));
                                }
                                w.somaProgresso();
                            }
                            w.Close();
                            Conexoes.Utilz.Alerta("Finalizado", "", System.Windows.MessageBoxImage.Information);

                            Conexoes.Utilz.ShowReports(retorno);
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                Conexoes.Utilz.Alerta(ex.Message + "\n" + ex.StackTrace);
                return;
            }

            return;

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

                string statusTexto = string.Format("{0} " + Constantes.Status + " {1}",
                    sq.Value.Split(","[0]).Length.ToString(),
                    sq.Value
                    );
                Funcoes.Propriedade_Edita_Cria(Nivel_Etapa, Constantes.Tab, Constantes.Status, statusTexto);

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


        private void ApagaAtributo()
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





            var pcs = items.Select(x => new Peca(x)).ToList();

            var atributos = pcs.SelectMany(x => x.GetPropriedades(new List<string> { Constantes.Tab })).Select(x=>x.Coluna).Distinct().ToList().FindAll(x=>x!=Constantes.Hierarquia);

            if(atributos.Count==0)
            {
                Conexoes.Utilz.Alerta($"Nenhum atributo do tipo [{Constantes.Tab}] encontrado.");
                return;
            }

            var sel = Conexoes.Utilz.SelecionarObjetos(atributos, null, "Selecione");

            if(sel.Count>0)
            {
                if(Conexoes.Utilz.Pergunta("Tem certeza q deseja apagar os atributos selecionados das peças selecionadas?"))
                {
                    Conexoes.Wait w = new Wait(pcs.Count, "Apagando...");
                    w.Show();

                    foreach(var pc in pcs)
                    {
                        foreach(var c in  sel)
                        {
                            Funcoes.Apagar_Propriedade(pc.ModelItem, Constantes.Tab, c);
                            
                        }
                        w.somaProgresso();
                    }
                    w.Close();
                }
            }
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
                Utilz.Alerta("Nenhuma peça do tipo Medabil (member) encontrada na seleção.");
                return;
            }

            var props = items.ToList().Select(x => new Peca(x));
            var cols = props.SelectMany(x => x.GetPropriedades(new List<string> { Constantes.Tab }).Select(y=> y.Coluna).ToList().FindAll(y=>y!= Constantes.Hierarquia)).Distinct().ToList();

       

            var mm = new Menus.SetarAtributo(cols);
            mm.ShowDialog();


            if (mm.DialogResult == DialogResult.OK)
            {
                Conexoes.Wait w = new Wait(items.Count, $"Setando atributo [{mm.txt_propriedade.Text}] = [{mm.txt_valor.Text}] em  {items.Count} itens");
                w.Show();
                foreach (ModelItem item in items)
                {
                    Funcoes.Propriedade_Edita_Cria(item, mm.txt_grupo.Text, Utilz.RemoverCaracteresEspeciais(mm.txt_propriedade.Text,true), mm.txt_valor.Text);
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
                Funcoes.Propriedade_Edita_Cria(item, Constantes.Tab, Constantes.DataExecucao, date);
            }



        }
        private void RemoveData()
        {
            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
            if (doc.CurrentSelection.SelectedItems.Count == 0)
            {
                Utilz.Alerta("Nenhum elemento selecionado");
                return;
            }
            var confirmacao = Utilz.Pergunta("Confirmar remoção das datas de execução dos elementos selecionados?", "Confirmação");

            if (!confirmacao) return;





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
                Funcoes.Apagar_Propriedade(item, Constantes.Tab, Constantes.DataExecucao);
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
                    Funcoes.Propriedade_Edita_Cria(item, Constantes.Tab, tipo.Key + Constantes.QtdExec, tipo.Value.ToString());
                    Funcoes.Propriedade_Edita_Cria(item, Constantes.Tab, tipo.Key + Constantes.PesoExec, etapa.TypesNetWeight[tipo.Key].ToString());
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
            Conexoes.Utilz.Alerta(mensagem, "Finalizado", System.Windows.MessageBoxImage.Information);



        }
    }
}
