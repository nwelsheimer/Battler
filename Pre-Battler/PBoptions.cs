using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pre_Battler
{
    public partial class PBOptions : Form
    {
        public PBOptions()
        {
            InitializeComponent();
            txtDB.Text = Properties.Settings.Default.dbName;
            txtServer.Text = Properties.Settings.Default.dbServer;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.dbName = txtDB.Text;
            Properties.Settings.Default.dbServer = txtServer.Text;
            Properties.Settings.Default.Save();
            this.Close();
        }
    }
}
