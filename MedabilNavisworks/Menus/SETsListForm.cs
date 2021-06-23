using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MedabilNavisworks
{
    public partial class SETsListForm : Form
    {
        private string retorno = "";
        public SETsListForm()
        {
            InitializeComponent();
        }

        private void SETsListForm_Load(object sender, EventArgs e)
        {

        }

        public static string Wait(IList<string> SETsFolders)
        {
            SETsListForm dialog = new SETsListForm();
            dialog.comboBox1.Items.AddRange(SETsFolders.ToArray());
            dialog.comboBox1.SelectedIndex = 0;
            dialog.ShowDialog();

            return dialog.retorno;
        }



        private void button2_Click_1(object sender, EventArgs e)
        {
            retorno = comboBox1.SelectedItem.ToString();
            this.Close();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            retorno = "";
            this.Close();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
