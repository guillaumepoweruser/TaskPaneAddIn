using System;
using Microsoft.Office.Tools;
using PPTTaskPaneAddIn.Views;

namespace PPTTaskPaneAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            TaskPaneUserControlPpt taskPaneUserControlPpt = new TaskPaneUserControlPpt();
            CustomTaskPane customTaskPane = CustomTaskPanes.Add(taskPaneUserControlPpt, "Task pane issue with WPF");
            customTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            customTaskPane.Width = 800;
            customTaskPane.Visible = true;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += new EventHandler(ThisAddIn_Startup);
            Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
