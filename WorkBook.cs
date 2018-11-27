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
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;


using Syncfusion.XlsIO;
using Syncfusion.Windows.Forms;
using Syncfusion.GridExcelConverter;
using Syncfusion.Windows.Forms.Grid;
using Syncfusion.GridHelperClasses;


namespace ExcelLikeUI
{
    /// <summary>
    /// Summary description for Form1.
    /// </summary>
    public class WorkBook : Office2007Form
    {
        #region Private Variables
        public Syncfusion.Windows.Forms.TabBarSplitterControl tabBarSplitterControl;
        public Syncfusion.Windows.Forms.Grid.GridAwareTextBox gridAwareTextBox1;
        public Syncfusion.Windows.Forms.Grid.GridAwareTextBox gridAwareTextBox2;
        private Syncfusion.Windows.Forms.Tools.XPMenus.ParentBarItem parentBarItem;
        private Syncfusion.Windows.Forms.Tools.XPMenus.BarItem insertRowBarItem;
        private Syncfusion.Windows.Forms.Tools.ContextMenuStripEx gridCMStrip;
        private ToolStripMenuItem cutToolStripMenuItem;
        private ToolStripMenuItem copyToolStripMenuItem;
        private ToolStripMenuItem pasteToolStripMenuItem;
        private ToolStripMenuItem deleteToolStripMenuItem;
        private ToolStripSeparator toolStripSeparator1;
        private ToolStripMenuItem hyperlinkToolStripMenuItem;
        private System.ComponentModel.IContainer components = null;
        LayoutSupportHelper layoutHelper;
        internal GridControl _grid = null;
        Form1 form;
        #endregion

        #region Constructor
        public WorkBook(Form1 frm)
        {
            //
            // Required for Windows Form Designer support
            //
            InitializeComponent();
            this.MyInit();
            form = frm;
        }
        #endregion

        #region Override Methods
        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }
        #endregion

        #region Windows Form Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(WorkBook));
            this.tabBarSplitterControl = new Syncfusion.Windows.Forms.TabBarSplitterControl();
            this.gridAwareTextBox1 = new Syncfusion.Windows.Forms.Grid.GridAwareTextBox();
            this.gridAwareTextBox2 = new Syncfusion.Windows.Forms.Grid.GridAwareTextBox();
            this.parentBarItem = new Syncfusion.Windows.Forms.Tools.XPMenus.ParentBarItem();
            this.insertRowBarItem = new Syncfusion.Windows.Forms.Tools.XPMenus.BarItem();
            this.gridCMStrip = new Syncfusion.Windows.Forms.Tools.ContextMenuStripEx();
            this.cutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.copyToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.pasteToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.deleteToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.hyperlinkToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tabBarSplitterControl.SuspendLayout();
            this.gridCMStrip.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabBarSplitterControl
            // 
            this.tabBarSplitterControl.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(219)))), ((int)(((byte)(232)))), ((int)(((byte)(249)))));
            this.tabBarSplitterControl.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.tabBarSplitterControl.Controls.Add(this.gridAwareTextBox1);
            this.tabBarSplitterControl.Controls.Add(this.gridAwareTextBox2);
            this.tabBarSplitterControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabBarSplitterControl.EnabledColor = System.Drawing.SystemColors.WindowText;
            this.tabBarSplitterControl.Location = new System.Drawing.Point(0, 0);
            this.tabBarSplitterControl.Name = "tabBarSplitterControl";
            this.tabBarSplitterControl.Office2007ScrollBars = true;
            this.tabBarSplitterControl.Size = new System.Drawing.Size(776, 502);
            this.tabBarSplitterControl.SplitBars = ((Syncfusion.Windows.Forms.DynamicSplitBars)((Syncfusion.Windows.Forms.DynamicSplitBars.SplitRows | Syncfusion.Windows.Forms.DynamicSplitBars.SplitColumns)));
            this.tabBarSplitterControl.Style = Syncfusion.Windows.Forms.TabBarSplitterStyle.Office2007;
            this.tabBarSplitterControl.TabFolderDelta = 11;
            this.tabBarSplitterControl.TabIndex = 0;
            this.tabBarSplitterControl.Text = "tabBarSplitterControl1";
            this.tabBarSplitterControl.ActivePageChanging += new System.Windows.Forms.ControlEventHandler(this.tabBarSplitterControl_ActivePageChanging);
            
            // 
            // gridAwareTextBox1
            // 
            this.gridAwareTextBox1.DisabledBackColor = System.Drawing.SystemColors.Window;
            this.gridAwareTextBox1.EnabledBackColor = System.Drawing.SystemColors.Window;
            this.gridAwareTextBox1.Location = new System.Drawing.Point(-100, -100);
            this.gridAwareTextBox1.Name = "gridAwareTextBox1";
            this.gridAwareTextBox1.Size = new System.Drawing.Size(100, 20);
            this.gridAwareTextBox1.TabIndex = 1;
            this.gridAwareTextBox1.BorderStyle = BorderStyle.FixedSingle;
            // 
            // gridAwareTextBox2
            // 
            this.gridAwareTextBox2.DisabledBackColor = System.Drawing.SystemColors.Window;
            this.gridAwareTextBox2.EnabledBackColor = System.Drawing.SystemColors.Window;
            this.gridAwareTextBox2.Location = new System.Drawing.Point(-100, -100);
            this.gridAwareTextBox2.Name = "gridAwareTextBox2";
            this.gridAwareTextBox2.Size = new System.Drawing.Size(100, 20);
            this.gridAwareTextBox2.TabIndex = 2;
            this.gridAwareTextBox2.BorderStyle = BorderStyle.FixedSingle;
            // 
            // parentBarItem
            // 
            this.parentBarItem.BarName = "parentBarItem";
            this.parentBarItem.Items.AddRange(new Syncfusion.Windows.Forms.Tools.XPMenus.BarItem[] {
            this.insertRowBarItem});
            this.parentBarItem.MetroColor = System.Drawing.Color.LightSkyBlue;
            this.parentBarItem.ShowToolTipInPopUp = false;
            this.parentBarItem.SizeToFit = true;
            // 
            // insertRowBarItem
            // 
            this.insertRowBarItem.BarName = "insertRowBarItem";
            this.insertRowBarItem.ID = "insertRowBarItem";
            this.insertRowBarItem.ShowToolTipInPopUp = false;
            this.insertRowBarItem.SizeToFit = true;
            this.insertRowBarItem.Text = "Insert Row";
            // 
            // gridCMStrip
            // 
            this.gridCMStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.cutToolStripMenuItem,
            this.copyToolStripMenuItem,
            this.pasteToolStripMenuItem,
            this.deleteToolStripMenuItem,
            this.toolStripSeparator1,
            this.hyperlinkToolStripMenuItem});
            this.gridCMStrip.MetroColor = System.Drawing.Color.FromArgb(((int)(((byte)(17)))), ((int)(((byte)(158)))), ((int)(((byte)(218)))));
            this.gridCMStrip.Name = "gridCMStrip";
            this.gridCMStrip.Size = new System.Drawing.Size(126, 120);
            this.gridCMStrip.Style = Syncfusion.Windows.Forms.Tools.ContextMenuStripEx.ContextMenuStyle.Default;
            // 
            // cutToolStripMenuItem
            // 
            this.cutToolStripMenuItem.Name = "cutToolStripMenuItem";
            this.cutToolStripMenuItem.Size = new System.Drawing.Size(125, 22);
            this.cutToolStripMenuItem.Text = "Cut";
            this.cutToolStripMenuItem.Click += new System.EventHandler(this.cutMenuItem_Click);
            // 
            // copyToolStripMenuItem
            // 
            this.copyToolStripMenuItem.Name = "copyToolStripMenuItem";
            this.copyToolStripMenuItem.Size = new System.Drawing.Size(125, 22);
            this.copyToolStripMenuItem.Text = "Copy";
            this.copyToolStripMenuItem.Click += new System.EventHandler(this.copyMenuItem_Click);
            // 
            // pasteToolStripMenuItem
            // 
            this.pasteToolStripMenuItem.Name = "pasteToolStripMenuItem";
            this.pasteToolStripMenuItem.Size = new System.Drawing.Size(125, 22);
            this.pasteToolStripMenuItem.Text = "Paste";
            this.pasteToolStripMenuItem.Click += new System.EventHandler(this.pasteMenuItem_Click);
            // 
            // deleteToolStripMenuItem
            // 
            this.deleteToolStripMenuItem.Name = "deleteToolStripMenuItem";
            this.deleteToolStripMenuItem.Size = new System.Drawing.Size(125, 22);
            this.deleteToolStripMenuItem.Text = "Delete";
            this.deleteToolStripMenuItem.Click += new System.EventHandler(this.deleteMenuItem_Click_1);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(122, 6);
            // 
            // hyperlinkToolStripMenuItem
            // 
            this.hyperlinkToolStripMenuItem.Name = "hyperlinkToolStripMenuItem";
            this.hyperlinkToolStripMenuItem.Size = new System.Drawing.Size(125, 22);
            this.hyperlinkToolStripMenuItem.Text = "Hyperlink";
            this.hyperlinkToolStripMenuItem.Click += new System.EventHandler(this.HyperLinkMenuItem_Click);
            // 
            // WorkBook
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(776, 502);
            this.Controls.Add(this.tabBarSplitterControl);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(600, 400);
            this.ControlBox = false;
            this.Name = "WorkBook";
            this.Text = "Form2";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            
            this.tabBarSplitterControl.ResumeLayout(false);
            this.tabBarSplitterControl.PerformLayout();
            this.gridCMStrip.ResumeLayout(false);
            this.ResumeLayout(false);
        }

        internal ArrayList HiddenSheets = new ArrayList();

        //Add the sheet 
        int i = 0;
        private void MyInit()
        {
            # region Initial Settings
            this.tabBarSplitterControl.SuspendLayout();
            this.SuspendLayout();
            TabBarPage tabBarPage = new TabBarPage();
            tabBarPage.TabBackColor = Color.FromArgb(219, 232, 249);
            GridControl _grid;
            GridModel model = new GridModel();
            SampleGrid.SetupGridModel(model);
            _grid = new SampleGrid(model);
            // 
            // _grid
            // 
            _grid.ContextMenuStrip = gridCMStrip;
            _grid.Location = new System.Drawing.Point(0, 0);
            _grid.Name = string.Format("gridControl{0}", i + 1);
            _grid.SmartSizeBox = false;
            _grid.Text = string.Format("gridControl{0}", i + 1);

            // 
            // tabBarPage
            // 
            _grid.PersistAppearanceSettings = false;
            _grid.ThemesEnabled = false;
            _grid.GridVisualStyles = GridVisualStyles.Metro;
            model.Options.GridVisualStyles = GridVisualStyles.Metro;
            GridMetroColors grid = new GridMetroColors();
            grid.HeaderColor.NormalColor = Color.Red;
            _grid.SetMetroStyle(grid);
            tabBarPage.Controls.Add(_grid);
            tabBarPage.Location = new System.Drawing.Point(0, 0);
            tabBarPage.Name = string.Format("tabBarPage{0}", i + 1);
            tabBarPage.SplitBars = Syncfusion.Windows.Forms.DynamicSplitBars.Both;
            tabBarPage.Text = string.Format("Sheet{0}", i + 1);
            //tabBarPage.ThemesEnabled = true;
            this.tabBarSplitterControl.TabBarPages.Add(tabBarPage);

            this._grid.Model.Properties.MarkColHeader = true;
            this._grid.Model.Properties.MarkRowHeader = true;

            GridCellRendererBase renderer = this._grid.CellRenderers["FormulaCell"];
            if (renderer is GridFormulaCellRenderer)
            {
                GridFormulaCellRenderer textBoxRenderer = renderer as GridFormulaCellRenderer;
            }


            i++;
            this._grid.TableStyle.WrapText = false;
            this.tabBarSplitterControl.ResumeLayout(true);
            this.ResumeLayout(true);
            #endregion
        }

        void TextBox_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                if (sender is TextBoxBase)
                {
                    TextBoxBase textbox = sender as TextBoxBase;
                    gridCMStrip.Show(textbox, e.Location);
                }
            }
        }

        #endregion

        #region Create New Sheet

        /// <summary>
        /// Add a new worksheet for the SpreadsheetControl
        /// </summary>
        public void AddNewWorkheet()
        {
            TabBarPage tabBarPage = new TabBarPage();
            tabBarPage.TabBackColor = Color.FromArgb(219, 232, 249);
            GridControl _grid;
            GridModel model = new GridModel();
            SampleGrid.SetupGridModel(model);
            _grid = new SampleGrid(model);

            // 
            // _grid
            // 
            _grid.MarkColHeader = true;
            _grid.MarkRowHeader = true;
            _grid.ContextMenuStrip = gridCMStrip;
            _grid.Location = new System.Drawing.Point(0, 0);
            _grid.Name = string.Format("gridControl{0}", i + 1);
            _grid.SmartSizeBox = false;
            _grid.Text = string.Format("gridControl{0}", i + 1);
            _grid.ThemesEnabled = true;


            
            // 
            // tabBarPage
            // 
            tabBarPage.Controls.Add(_grid);
            tabBarPage.Location = new System.Drawing.Point(0, 0);
            tabBarPage.Name = string.Format("tabBarPage{0}", i + 1);
            tabBarPage.SplitBars = Syncfusion.Windows.Forms.DynamicSplitBars.Both;
            tabBarPage.ForeColor = Color.Green;
            tabBarPage.Text = string.Format("Sheet{0}", i + 1);
            this.tabBarSplitterControl.TabBarPages.Add(tabBarPage);
            GridCellRendererBase renderer = this._grid.CellRenderers["FormulaCell"];
            if (renderer is GridFormulaCellRenderer)
            {
                GridFormulaCellRenderer textBoxRenderer = renderer as GridFormulaCellRenderer;
            }
            i++;
        }
        #endregion
        # region Menu Handlers
        private void HyperLinkMenuItem_Click(object sender, System.EventArgs e)
        {
            GridCurrentCell cc = this._grid.CurrentCell;
            GridStyleInfo style = this._grid.Model[cc.RowIndex, cc.ColIndex];
            if (style.CellType == "LinkLabel")
                style.CellType = "FormulaCell";
            else
            {
                style.CellType = "LinkLabel";
                style.Tag = style.Text;
            }
        }

        private void cutMenuItem_Click(object sender, System.EventArgs e)
        {
            this._grid.Model.CutPaste.Cut();
        }

        private void copyMenuItem_Click(object sender, System.EventArgs e)
        {
            this._grid.Model.CutPaste.Copy();
        }

        private void pasteMenuItem_Click(object sender, System.EventArgs e)
        {
            this._grid.Model.CutPaste.Paste();
        }

        private void deleteMenuItem_Click_1(object sender, System.EventArgs e)
        {
            this._grid.Model.Clear(true);
        }
        # endregion

        #region Events

        private void tabBarSplitterControl_ActivePageChanging(object sender, System.Windows.Forms.ControlEventArgs e)
        {
            if (e.Control != null)
                foreach (Control control in e.Control.Controls)
                {
                    if (control is GridControl)
                    {
                        this._grid = control as GridControl;
                        break;
                    }
                }
        }

        #endregion
    }
}
