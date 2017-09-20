using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using SolidWorks.Interop.swconst;

namespace TestAddIn
{
    public class UserPMPage
    {
        //Local Objects
        IPropertyManagerPage2 swPropertyPage = null;
        PMPHandler handler = null;
        ISldWorks iSwApp = null;
        SwAddin userAddin = null;

        #region Property Manager Page Controls
        IPropertyManagerPageTab tab1;
        IPropertyManagerPageTab tab2;

        //Groups
        IPropertyManagerPageGroup group1;
        IPropertyManagerPageGroup group2;

        //Controls
        IPropertyManagerPageSelectionbox selection1;
        IPropertyManagerPageNumberbox num1;
        IPropertyManagerPageButton button1;

        //Control IDs
        public const int group1ID = 0;
        public const int group2ID = 1;

        public const int selection1ID = 8;
        public const int num1ID = 9;
        public const int button1ID = 11;

        public const int tab1ID = 12;
        public const int tab2ID = 13;
        #endregion

        public UserPMPage(SwAddin addin)
        {
            userAddin = addin;
            if (userAddin != null)
            {
                iSwApp = (ISldWorks)userAddin.SwApp;
                CreatePropertyManagerPage();
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("SwAddin not set.");
            }
        }


        protected void CreatePropertyManagerPage()
        {
            int errors = -1;
            int options = (int)swPropertyManagerPageOptions_e.swPropertyManagerOptions_OkayButton |
                (int)swPropertyManagerPageOptions_e.swPropertyManagerOptions_CancelButton;

            handler = new PMPHandler(userAddin);
            swPropertyPage = (IPropertyManagerPage2)iSwApp.CreatePropertyManagerPage("Estrusion PMP", options, handler, ref errors);
            if (swPropertyPage != null && errors == (int)swPropertyManagerPageStatus_e.swPropertyManagerPage_Okay)
            {
                try
                {
                    AddControls();
                }
                catch (Exception e)
                {
                    iSwApp.SendMsgToUser2(e.Message, 0, 0);
                }
            }
        }


        //Controls are displayed on the page top to bottom in the order 
        //in which they are added to the object.
        protected void AddControls()
        {
            short controlType = -1;
            short align = -1;
            int options = -1;

            //Add tabs
            tab1 = (PropertyManagerPageTab)swPropertyPage.AddTab(tab1ID, "Step 1", "", 0);
            tab2 = (PropertyManagerPageTab)swPropertyPage.AddTab(tab2ID, "Step 2", "", 0);

            //Add the groups
            options = (int)swAddGroupBoxOptions_e.swGroupBoxOptions_Expanded |
                      (int)swAddGroupBoxOptions_e.swGroupBoxOptions_Visible;

            group1 = (IPropertyManagerPageGroup)swPropertyPage.AddGroupBox(group1ID, "Modify Extrusion", options);

            //options = (int)swAddGroupBoxOptions_e.swGroupBoxOptions_Checkbox |
            //          (int)swAddGroupBoxOptions_e.swGroupBoxOptions_Visible;

            group2 = (IPropertyManagerPageGroup)swPropertyPage.AddGroupBox(group2ID, "Rebuild the Model", options);

            //Add the controls to group1
            
            //button1
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Button;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                      (int)swAddControlOptions_e.swControlOptions_Visible;

            button1 = (IPropertyManagerPageButton)group2.AddControl(button1ID, controlType, "Redraw", align, options, "Click to redraw");

            //Add controls to group2
            //selection1
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Selectionbox;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                      (int)swAddControlOptions_e.swControlOptions_Visible;

            selection1 = (IPropertyManagerPageSelectionbox)group1.AddControl(selection1ID, controlType, "Sample Selection", align, options, "Displays features selected in main view");
            if (selection1 != null)
            {
                int[] filter = { (int)swSelectType_e.swSelEDGES, (int)swSelectType_e.swSelVERTICES };
                selection1.Height = 40;
                selection1.SetSelectionFilters(filter);
            }

            //num1
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Numberbox;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                      (int)swAddControlOptions_e.swControlOptions_Visible;

            num1 = (IPropertyManagerPageNumberbox)group1.AddControl(num1ID, controlType, "Sample Numberbox", align, options, "Allows for numerical input");
            if (num1 != null)
            {
                num1.Value = 50.0;
                num1.SetRange((int)swNumberboxUnitType_e.swNumberBox_UnitlessDouble, 0.0, 100.0, 0.01, true);
            }

            tab1.AddGroupBox(group1ID, "", 0);
            tab2.AddGroupBox(group2ID, "", 0);
            tab1.Activate();
            DisplayControls(1);
        }

        public void Show()
        {
            if (swPropertyPage != null)
            {
                swPropertyPage.Show();
            }
        }

        public void DisplayControls(int ID)
        {
            if(ID == 1)
            {
                //Tab1 selected; hide group2 controls
                group2.Visible = false;
                group1.Visible = true;
            }else
            {
                group1.Visible = false;
                group2.Visible = true;
            }
        }
    }
}
