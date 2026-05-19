using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

namespace sign_templater
{
    public partial class ThisAddIn
    {
        private SidebarControl myUserControl;
        private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new TuregRibbon();
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            myUserControl = new SidebarControl();
            myCustomTaskPane = this.CustomTaskPanes.Add(myUserControl, "TUREG Template Tool");
            myCustomTaskPane.Width = 350;
            myCustomTaskPane.Visible = true;
        }

        public void ToggleSidebar()
        {
            if (myCustomTaskPane != null)
            {
                myCustomTaskPane.Visible = !myCustomTaskPane.Visible;
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
