using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace VisioPlugin
{
    [ComVisible(true)]
    public class RibbonExtension : Office.IRibbonExtensibility
    {
        private readonly ThisAddIn addIn;

        public RibbonExtension(ThisAddIn addIn)
        {
            this.addIn = addIn ?? throw new ArgumentNullException(nameof(addIn));
        }

        public string GetCustomUI(string ribbonID)
        {
            return @"
            <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='Ribbon_Load'>
              <ribbon>
                <tabs>
                  <tab id='ScdCreatorTab' label='Scd_Creator'>
                    <group id='LibraryManagementGroup' label='Library Management'>
                      <button id='RefreshLibrariesButton' label='Refresh Libraries' onAction='OnRefreshLibrariesButtonClick' />
                      <dropDown id='CategorySelectionDropDown' label='Select Category' 
                                onAction='OnCategorySelectionChange' 
                                getItemCount='GetCategoryCount' 
                                getItemID='GetCategoryId' 
                                getItemLabel='GetCategoryLabel' 
                                getSelectedItemID='GetSelectedCategoryID' />
                      <button id='AddTestShapeButton' label='Add Test Shape' onAction='OnAddTestShapeClick' />
                    </group>
                    <group id='AIInteractionGroup' label='AI Interaction'>
                      <editBox id='APIEndpointTextBox' label='API Endpoint' onChange='OnAPIEndpointChange' sizeString='http://localhost:11434/v1' />
                      <dropDown id='ModelSelectionDropDown' label='Select Model' 
                                getItemCount='GetModelCount' 
                                getItemLabel='GetModelLabel' 
                                onAction='OnModelSelectionChange' />
                      <button id='ConnectButton' label='Connect to AI' onAction='OnConnectButtonClick' />
                      <labelControl id='ConnectionStatus' getLabel='GetConnectionStatus'/>
                    </group>
                  </tab>
                </tabs>
              </ribbon>
            </customUI>";
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            addIn.Ribbon = ribbonUI ?? throw new ArgumentNullException(nameof(ribbonUI));
        }

        public void OnRefreshLibrariesButtonClick(Office.IRibbonControl control)
        {
            addIn.OnRefreshLibrariesButtonClick(control);
        }

        public void OnCategorySelectionChange(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            Debug.WriteLine($"RibbonExtension: Selected ID: {selectedId}, Selected Index: {selectedIndex}");
            addIn.OnCategorySelectionChange(control, selectedId, selectedIndex);
        }

        public void OnAddTestShapeClick(Office.IRibbonControl control)
        {
            addIn.OnAddTestShapeClick(control);
        }

        public int GetCategoryCount(Office.IRibbonControl control)
        {
            return addIn.GetCategoryCount(control);
        }

        public string GetCategoryId(Office.IRibbonControl control, int index)
        {
            var categories = addIn.GetCategories();
            if (index < 0 || index >= categories.Length)
            {
                throw new ArgumentOutOfRangeException(nameof(index), "Index is out of range.");
            }
            string categoryId = categories[index];
            Debug.WriteLine($"GetCategoryId called for index {index}, returning: {categoryId}");
            return categoryId;
        }

        public string GetCategoryLabel(Office.IRibbonControl control, int index)
        {
            var categories = addIn.GetCategories();
            if (index < 0 || index >= categories.Length)
            {
                throw new ArgumentOutOfRangeException(nameof(index), "Index is out of range.");
            }
            string categoryLabel = categories[index];
            Debug.WriteLine($"GetCategoryLabel called for index {index}, returning: {categoryLabel}");
            return categoryLabel;
        }

        public string GetSelectedCategoryID(Office.IRibbonControl control)
        {
            if (!string.IsNullOrEmpty(addIn.CurrentCategory))
            {
                Debug.WriteLine($"GetSelectedCategoryID called, returning: {addIn.CurrentCategory}");
                return addIn.CurrentCategory;
            }
            else
            {
                var categories = addIn.GetCategories();
                if (categories.Length > 0)
                {
                    addIn.CurrentCategory = categories[0];
                    Debug.WriteLine($"GetSelectedCategoryID called, setting CurrentCategory to default: {addIn.CurrentCategory}");
                    return addIn.CurrentCategory;
                }
                else
                {
                    return null;
                }
            }
        }

        // AI Interaction Controls

        public void OnAPIEndpointChange(Office.IRibbonControl control, string text)
        {
            addIn.OnAPIEndpointChange(control, text);
        }

        public void OnConnectButtonClick(Office.IRibbonControl control)
        {
            addIn.OnConnectButtonClick(control);
        }


        public int GetModelCount(Office.IRibbonControl control)
        {
            return addIn.GetModelCount(control);
        }

        public string GetModelLabel(Office.IRibbonControl control, int index)
        {
            return addIn.GetModelLabel(control, index);
        }

        public void OnModelSelectionChange(Office.IRibbonControl control, string selectedItemId)
        {
            addIn.OnModelSelectionChange(control, selectedItemId);
        }

        // Connection status label control update
        public string GetConnectionStatus(Office.IRibbonControl control)
        {
            return addIn.isConnected ? "Connected" : "Not Connected";
        }
    }
}
