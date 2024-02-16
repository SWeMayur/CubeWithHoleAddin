using InvAddIn;
using Inventor;
using Microsoft.Win32;
using System;
using System.Runtime.InteropServices;

namespace CubeWithHoleAddin
{
    /// <summary>
    /// This is the primary AddIn Server class that implements the ApplicationAddInServer interface
    /// that all Inventor AddIns are required to implement. The communication between Inventor and
    /// the AddIn is via the methods on this interface.
    /// </summary>
    [GuidAttribute("6a5b08ec-fa48-46de-b9ba-f81de901d384")]
    public class StandardAddInServer : Inventor.ApplicationAddInServer
    {
        public static Application m_inventorApplication;
        public ButtonDefinition m_helloWorldButton;

        public StandardAddInServer()
        {
        }

        #region ApplicationAddInServer Members

        public void Activate(Inventor.ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            m_inventorApplication = addInSiteObject.Application;
            AddButton();
        }

        public void Deactivate()
        {
            ReleaseObjects();
        }

        public void ExecuteCommand(int commandID)
        {
            // Note: this method is now obsolete; you should use the 
            // ControlDefinition functionality for implementing commands.
        }

        public object Automation
        {
            get
            {
                // TODO: Add ApplicationAddInServer.Automation getter implementation
                return null;
            }
        }

        #endregion

        private void AddButton()
        {
            UserInterfaceManager uiMgr = m_inventorApplication.UserInterfaceManager;

            // Create a button definition
            m_helloWorldButton = m_inventorApplication.CommandManager.ControlDefinitions.AddButtonDefinition(
                "Cube With Hole", "HelloWorldCmd", CommandTypesEnum.kNonShapeEditCmdType, "{6a5b08ec-fa48-46de-b9ba-f81de901d384}", "Display Hello World Message");

            // Get the tools tab
            Ribbon toolsRibbon = uiMgr.Ribbons["Part"];

            //Get the tools tab
            RibbonTab toolsTab = toolsRibbon.RibbonTabs["id_TabTools"];

            // Get the tools panel within the tools tab
            RibbonPanel toolsPanel = toolsTab.RibbonPanels["id_PanelP_ShowPanels"];

            // Add the button to the panel
            CommandControl commandControl = toolsPanel.CommandControls.AddButton(
                m_helloWorldButton,
                false,
                true,
                "",
                false);
            // Wire up the event handler
            m_helloWorldButton.OnExecute += ButtonHandler.OnExecute;
        }

        private void ReleaseObjects()
        {
            m_inventorApplication = null;
            m_helloWorldButton = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}