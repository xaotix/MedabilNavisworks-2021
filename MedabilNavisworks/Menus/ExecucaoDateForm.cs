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
    public partial class ExecucaoDateForm : Form
    {
        private DateTime? retorno = null;
        
        public ExecucaoDateForm()
        {
            InitializeComponent();
        }

        private void ExecucaoDateForm_Load(object sender, EventArgs e)
        {

        }

        public static DateTime? Wait(DateTime? initialDate = null)
        {
            ExecucaoDateForm dialog = new ExecucaoDateForm();
            if (initialDate != null) dialog.dateTimePicker1.Value = (DateTime)initialDate;
            dialog.ShowDialog();

            return dialog.retorno;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            retorno = dateTimePicker1.Value;
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
