using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MedabilNavisworks.Menus
{
    public partial class SetarAtributo : Form
    {
        public List<string> atributos_existentes { get; set; } = new List<string>();
        public SetarAtributo(List<string> atributos)
        {
            this.atributos_existentes = atributos;
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Hide();
            if(this.atributos_existentes.Count==0)
            {
                return;
            }
            var txt = Conexoes.Utilz.SelecionarObjeto(this.atributos_existentes,null,"Selecione");
            if(txt!=null)
            {
                this.txt_propriedade.Text = txt;
            }
            this.Show();
        }
    }
}
