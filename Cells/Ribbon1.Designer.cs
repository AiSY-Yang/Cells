namespace Cells
{
	partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
	{
		/// <summary>
		/// 必需的设计器变量。
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		public Ribbon1()
			: base(Globals.Factory.GetRibbonFactory())
		{
			InitializeComponent();
		}

		/// <summary> 
		/// 清理所有正在使用的资源。
		/// </summary>
		/// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region 组件设计器生成的代码

		/// <summary>
		/// 设计器支持所需的方法 - 不要修改
		/// 使用代码编辑器修改此方法的内容。
		/// </summary>
		private void InitializeComponent()
		{
			Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
			Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
			Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
			Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
			this.tab1 = this.Factory.CreateRibbonTab();
			this.group1 = this.Factory.CreateRibbonGroup();
			this.Light = this.Factory.CreateRibbonCheckBox();
			this.group2 = this.Factory.CreateRibbonGroup();
			this.box1 = this.Factory.CreateRibbonBox();
			this.editBox1 = this.Factory.CreateRibbonEditBox();
			this.editBox2 = this.Factory.CreateRibbonEditBox();
			this.splitChar = this.Factory.CreateRibbonComboBox();
			this.separator1 = this.Factory.CreateRibbonSeparator();
			this.concatenation = this.Factory.CreateRibbonButton();
			this.group3 = this.Factory.CreateRibbonGroup();
			this.button1 = this.Factory.CreateRibbonButton();
			this.ctrlL = this.Factory.CreateRibbonButton();
			this.ctrlR = this.Factory.CreateRibbonButton();
			this.tab1.SuspendLayout();
			this.group1.SuspendLayout();
			this.group2.SuspendLayout();
			this.box1.SuspendLayout();
			this.group3.SuspendLayout();
			this.SuspendLayout();
			// 
			// tab1
			// 
			this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
			this.tab1.Groups.Add(this.group1);
			this.tab1.Groups.Add(this.group2);
			this.tab1.Groups.Add(this.group3);
			this.tab1.Label = "Cells";
			this.tab1.Name = "tab1";
			// 
			// group1
			// 
			this.group1.Items.Add(this.Light);
			this.group1.Label = "group1";
			this.group1.Name = "group1";
			// 
			// Light
			// 
			this.Light.Label = "聚光灯";
			this.Light.Name = "Light";
			this.Light.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Light_Click);
			// 
			// group2
			// 
			this.group2.Items.Add(this.box1);
			this.group2.Items.Add(this.separator1);
			this.group2.Items.Add(this.concatenation);
			this.group2.Label = "并列同类项";
			this.group2.Name = "group2";
			// 
			// box1
			// 
			this.box1.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
			this.box1.Items.Add(this.editBox1);
			this.box1.Items.Add(this.editBox2);
			this.box1.Items.Add(this.splitChar);
			this.box1.Name = "box1";
			// 
			// editBox1
			// 
			this.editBox1.Label = "同类项列";
			this.editBox1.Name = "editBox1";
			this.editBox1.Text = "E";
			// 
			// editBox2
			// 
			this.editBox2.Label = "并列项列";
			this.editBox2.Name = "editBox2";
			this.editBox2.Text = "F";
			// 
			// splitChar
			// 
			ribbonDropDownItemImpl1.Label = "Tab(并列)";
			ribbonDropDownItemImpl2.Label = "Space(空格)";
			ribbonDropDownItemImpl3.Label = "-";
			ribbonDropDownItemImpl4.Label = ",";
			this.splitChar.Items.Add(ribbonDropDownItemImpl1);
			this.splitChar.Items.Add(ribbonDropDownItemImpl2);
			this.splitChar.Items.Add(ribbonDropDownItemImpl3);
			this.splitChar.Items.Add(ribbonDropDownItemImpl4);
			this.splitChar.Label = "分隔符 ";
			this.splitChar.Name = "splitChar";
			this.splitChar.Text = "Tab(并列)";
			// 
			// separator1
			// 
			this.separator1.Name = "separator1";
			// 
			// concatenation
			// 
			this.concatenation.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
			this.concatenation.Image = global::Cells.Properties.Resources.Exec;
			this.concatenation.Label = "执行";
			this.concatenation.Name = "concatenation";
			this.concatenation.ShowImage = true;
			this.concatenation.SuperTip = "主要用于一个条件匹配多个结果";
			this.concatenation.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Concatenation_Click);
			// 
			// group3
			// 
			this.group3.Items.Add(this.button1);
			this.group3.Items.Add(this.ctrlL);
			this.group3.Items.Add(this.ctrlR);
			this.group3.Label = "单元格对齐";
			this.group3.Name = "group3";
			// 
			// button1
			// 
			this.button1.Label = "选中同色单元格";
			this.button1.Name = "button1";
			this.button1.SuperTip = "选中当前工作表内同一种文字颜色及底色的所有单元格";
			this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SameFormat);
			// 
			// ctrlL
			// 
			this.ctrlL.Label = "单元格左对齐";
			this.ctrlL.Name = "ctrlL";
			this.ctrlL.SuperTip = "向左对齐所有单元格";
			this.ctrlL.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Align_Click);
			// 
			// ctrlR
			// 
			this.ctrlR.Label = "单元格右对齐";
			this.ctrlR.Name = "ctrlR";
			this.ctrlR.SuperTip = "向右对其所有单元格";
			this.ctrlR.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Align_Click);
			// 
			// Ribbon1
			// 
			this.Name = "Ribbon1";
			this.RibbonType = "Microsoft.Excel.Workbook";
			this.Tabs.Add(this.tab1);
			this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Cells_Load);
			this.tab1.ResumeLayout(false);
			this.tab1.PerformLayout();
			this.group1.ResumeLayout(false);
			this.group1.PerformLayout();
			this.group2.ResumeLayout(false);
			this.group2.PerformLayout();
			this.box1.ResumeLayout(false);
			this.box1.PerformLayout();
			this.group3.ResumeLayout(false);
			this.group3.PerformLayout();
			this.ResumeLayout(false);

		}

		#endregion
		internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
		internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox Light;
		internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
		internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
		internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox1;
		internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox2;
		internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
		internal Microsoft.Office.Tools.Ribbon.RibbonComboBox splitChar;
		internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton concatenation;
		internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton ctrlL;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton ctrlR;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
	}

	partial class ThisRibbonCollection
	{
		internal Ribbon1 Cells
		{
			get { return this.GetRibbon<Ribbon1>(); }
		}
	}
}
