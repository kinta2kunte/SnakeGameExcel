using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SnakeGameExcel
{
    public partial class frmSettings : Form
    {
        public int gameMode = 0;
        public frmSettings()
        {
            InitializeComponent();
        }

        private void frmSettings_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = gameMode;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            gameMode = comboBox1.SelectedIndex;
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
