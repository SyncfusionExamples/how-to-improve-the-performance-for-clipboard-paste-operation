#region Copyright Syncfusion Inc. 2001 - 2015
//
//  Copyright Syncfusion Inc. 2001 - 2015. All rights reserved.
//
//  Use of this code is subject to the terms of our license.
//  A copy of the current license can be obtained at any time by e-mailing
//  licensing@syncfusion.com. Any infringement will be prosecuted under
//  applicable laws. 
//
#endregion

using System;
using System.Drawing;
using System.Windows.Forms;
using System.Collections;
using System.ComponentModel;
using System.ComponentModel.Design;
using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Design;
using Syncfusion.Collections;
using Syncfusion.ComponentModel;
using System.Drawing.Design;
using System.Collections.Specialized;
using System.Diagnostics;
using System.Runtime.Serialization;
using System.Reflection;
using Syncfusion.Runtime.Serialization;

using Syncfusion.Windows.Forms.Grid;

using Syncfusion.Windows.Forms.Tools;
using Syncfusion.GridHelperClasses;

namespace ExcelLikeUI
{
    public partial class Form1 : RibbonForm
    {
        #region [private Variables ]
        private static int accCount = 1;
        Panel TabBarPane = new Panel();
        Panel newPanel = new Panel();
        #endregion

        #region [Constructor]
        public Form1()
        {
            InitializeComponent();
            this.excelRibbon.RibbonStyle = RibbonStyle.Office2013;
            try
            {
                System.Drawing.Icon ico = new System.Drawing.Icon(GetIconFile(@"common\Images\Grid\Icon\sficon.ico"));
                this.Icon = ico;
            }
            catch { }

            this.BackStageInitializeComponent();
            this.backStageView1.HostForm = this;
            this.excelRibbon.BackStageView = this.backStageView1;
            newPanel.BorderStyle = BorderStyle.FixedSingle;
            xpToolBar1.Dock = DockStyle.Fill;
            newPanel.Size = xpToolBar1.Size;

            this.Controls.Remove(this.xpToolBar1);
            newPanel.Controls.Add(xpToolBar1);
        }
        #endregion

        #region Form Icon
        private string GetIconFile(string bitmapName)
        {
            for (int n = 0; n < 10; n++)
            {
                if (System.IO.File.Exists(bitmapName))
                    return bitmapName;

                bitmapName = @"..\" + bitmapName;
            }

            return bitmapName;
        }
        #endregion

        #region [Methods]
        public void PopulateFonts()
        {
            foreach (FontFamily ff in FontFamily.Families)
            {
                if (ff.IsStyleAvailable(FontStyle.Regular))
                {
                    toolStripFontfaceComboBox.Items.Add(ff.Name);
                }
            }
        }
        #endregion

        #region[Form Load]
        private void Form1_Load(object sender, EventArgs e)
        {
            #region [Ribbon Related Changes]

           
            this.toolStripFontfaceComboBox.SelectedIndex = 0;
            this.toolStripFontSizeComboBox.SelectedIndex = 3;
            toolStripButton45.Checked = true;

            PopulateFonts();
            toolStripCheckBox1.Checked = true;
            toolStripCheckBox2.Checked = true;
            #endregion

            #region [WorkBook]
            // Create a new child
            WorkBook workBook = new WorkBook(this);
            this.workBook = workBook;
            Panel BackPanel = new Panel();
            Panel topPanel = new Panel();
            this.Controls.Add(this.xpToolBar1);
            topPanel.Dock = DockStyle.Top;
            TabBarPane.Size = new System.Drawing.Size(500, 500);

            BackPanel.Dock = DockStyle.Fill;
            TabBarPane.Controls.Add(workBook.tabBarSplitterControl);
            this.xpToolBar1.Bar.Items.AddRange(new Syncfusion.Windows.Forms.Tools.XPMenus.BarItem[] {           
            this.barItem1,
            this.barItem2,
         });
            System.ComponentModel.ComponentResourceManager resources1 = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.barItem1.Image = ((Syncfusion.Windows.Forms.Tools.XPMenus.ImageExt)(resources1.GetObject("barItem1.Image")));
            this.barItem2.Image = ((Syncfusion.Windows.Forms.Tools.XPMenus.ImageExt)(resources1.GetObject("barItem2.Image")));

            workBook.tabBarSplitterControl.BringToFront();
            workBook.tabBarSplitterControl.Dock = DockStyle.Fill;

            TabBarPane.BringToFront();
            TabBarPane.Dock = DockStyle.Fill;
            BackPanel.Dock = DockStyle.Fill;
            this.Controls.Add(BackPanel);
            BackPanel.Controls.Add(TabBarPane);
            newPanel.Dock = DockStyle.Top;
            BackPanel.Controls.Add(newPanel);

            this.workBook.tabBarSplitterControl.EnableOffice2013Style = true;
            this.workBook.tabBarSplitterControl.ShowAddNewTabBarPageOption = true;
            this.workBook.tabBarSplitterControl.Style = TabBarSplitterStyle.Metro;                     

            foreach (Control ctrl in this.Controls)
            {
                if (ctrl is MdiClient)
                {
                    MdiClient mdiClient = (MdiClient)ctrl;
                    mdiClient.BackColor = Color.FromArgb(165, 195, 239);
                }
            }
            #endregion
        }
        #endregion

        #region [BackStage]
        private void BackStageInitializeComponent()
        {
            this.backStageView1 = new Syncfusion.Windows.Forms.BackStageView(this.components);
            this.backStage1 = new Syncfusion.Windows.Forms.BackStage();
            this.saveBackStageButton = new Syncfusion.Windows.Forms.BackStageButton();
            this.openBackStageButton = new Syncfusion.Windows.Forms.BackStageButton();
            this.exitBackStageButton = new Syncfusion.Windows.Forms.BackStageButton();
            //newTabPanel = new NewTabPanel();
            ((System.ComponentModel.ISupportInitialize)(this.backStage1)).BeginInit();
            this.backStage1.SuspendLayout();
            this.SuspendLayout();
            // 
            // backStageView1
            // 
            this.backStageView1.BackStage = this.backStage1;
            this.backStageView1.HostControl = null;
            // 
            // backStage1
            // 
            this.backStage1.AllowDrop = true;
            this.backStage1.BeforeTouchSize = new System.Drawing.Size(715, 449);
            this.backStage1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.backStage1.Controls.Add(this.openBackStageButton);
            this.backStage1.Controls.Add(this.saveBackStageButton);
            this.backStage1.Controls.Add(this.exitBackStageButton);
            this.backStage1.ItemSize = new System.Drawing.Size(138, 40);
            this.backStage1.Location = new System.Drawing.Point(1, 51);
            this.backStage1.Name = "backStage1";
            this.backStage1.OfficeColorScheme = Syncfusion.Windows.Forms.Tools.ToolStripEx.ColorScheme.Blue;
            this.backStage1.Size = new System.Drawing.Size(715, 449);
            this.backStage1.TabIndex = 4;
            // 
            // saveBackStageButton
            // 
            this.saveBackStageButton.Accelerator = "";
            this.saveBackStageButton.BackColor = System.Drawing.Color.Transparent;
            this.saveBackStageButton.BeforeTouchSize = new System.Drawing.Size(75, 23);
            this.saveBackStageButton.IsBackStageButton = false;
            this.saveBackStageButton.Location = new System.Drawing.Point(0, 53);
            this.saveBackStageButton.Name = "saveBackStageButton";
            this.saveBackStageButton.Size = new System.Drawing.Size(137, 38);
            this.saveBackStageButton.TabIndex = 4;
            this.saveBackStageButton.Text = "Save";
            

            // 
            // openBackStageButton
            // 
            this.openBackStageButton.Accelerator = "";
            this.openBackStageButton.BackColor = System.Drawing.Color.Transparent;
            this.openBackStageButton.BeforeTouchSize = new System.Drawing.Size(75, 23);
            this.openBackStageButton.IsBackStageButton = false;
            this.openBackStageButton.Location = new System.Drawing.Point(0, 53);
            this.openBackStageButton.Name = "openBackStageButton";
            this.openBackStageButton.Size = new System.Drawing.Size(137, 38);
            this.openBackStageButton.TabIndex = 4;
            this.openBackStageButton.Text = "Open";
            
            // 
            // exitBackStageButton
            // 
            this.exitBackStageButton.Accelerator = "";
            this.exitBackStageButton.BackColor = System.Drawing.Color.Transparent;
            this.exitBackStageButton.BeforeTouchSize = new System.Drawing.Size(75, 23);
            this.exitBackStageButton.IsBackStageButton = false;
            this.exitBackStageButton.Location = new System.Drawing.Point(0, 53);
            this.exitBackStageButton.Name = "exitBackStageButton";
            this.exitBackStageButton.Size = new System.Drawing.Size(137, 38);
            this.exitBackStageButton.TabIndex = 4;
            this.MinimumSize = new Size(867, 679);
            this.exitBackStageButton.Text = "Exit";
           
            // 
            // 
            // 
            ((System.ComponentModel.ISupportInitialize)(this.backStage1)).EndInit();
            this.backStage1.ResumeLayout(false);
            this.ResumeLayout(false);
        }
        #endregion

      
    }
}
