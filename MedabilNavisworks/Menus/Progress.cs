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
{ //lógica da barra de progresso.
    // é para ela executar via button após a lógica, e só parar quando a lógica for totalmente executada/carregada
    // ela só libera o software depois de a barra fechar

    public partial class Progress : Form
    {
        public Action Loader { get; set; }


        public Progress(Action loader)
        {
            InitializeComponent();
            if (loader == null)
                throw new ArgumentNullException();
            Loader = loader;
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            Task.Factory.StartNew(Loader).ContinueWith(t => { this.Close(); }, TaskScheduler.FromCurrentSynchronizationContext());
        }
    }
}
