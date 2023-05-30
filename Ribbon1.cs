using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SnakeGameExcel
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnInit_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.btnInit();
        }

        private void btnStart_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.btnStart();
        }

        private void btnLeft_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.nNextDirectionMove = 1;

        }

        private void btnRight_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.nNextDirectionMove = 2;

        }

        private void btnUp_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.nNextDirectionMove = 3;

        }

        private void btnDown_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.nNextDirectionMove = 4;

        }

        private void btnHelp_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.btnHelp();

        }

        private void btnSetup_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.btnSettings();

        }
    }
}
