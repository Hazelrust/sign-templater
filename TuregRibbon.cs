using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace sign_templater
{
    [ComVisible(true)]
    public class TuregRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public TuregRibbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
                        <ribbon>
                            <tabs>
                                <tab id='TUREGTab' label='TUREG Add-in'>
                                    <group id='TUREGGroup' label='Tools'>
                                        <button id='btnShowSidebar' label='Show Task Pane' 
                                                size='large' onAction='OnShowSidebar' 
                                                imageMso='AdpDiagramShowTable' />
                                    </group>
                                </tab>
                            </tabs>
                        </ribbon>
                    </customUI>";
        }

        #endregion

        #region Ribbon Callbacks

        public void OnRibbonLoad(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnShowSidebar(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.ToggleSidebar();
        }

        #endregion
    }
}
