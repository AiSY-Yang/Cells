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
			this.tab1 = this.Factory.CreateRibbonTab();
			this.group1 = this.Factory.CreateRibbonGroup();
			this.Light = this.Factory.CreateRibbonCheckBox();
			this.group2 = this.Factory.CreateRibbonGroup();
			this.editBox1 = this.Factory.CreateRibbonEditBox();
			this.editBox2 = this.Factory.CreateRibbonEditBox();
			this.button1 = this.Factory.CreateRibbonButton();
			this.button2 = this.Factory.CreateRibbonButton();
			this.tab1.SuspendLayout();
			this.group1.SuspendLayout();
			this.group2.SuspendLayout();
			this.SuspendLayout();
			// 
			// tab1
			// 
			this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
			this.tab1.Groups.Add(this.group1);
			this.tab1.Groups.Add(this.group2);
			this.tab1.Label = "Cells";
			this.tab1.Name = "tab1";
			// 
			// group1
			// 
			this.group1.Items.Add(this.Light);
			this.group1.Items.Add(this.button2);
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
			this.group2.Items.Add(this.editBox1);
			this.group2.Items.Add(this.editBox2);
			this.group2.Items.Add(this.button1);
			this.group2.Label = "group2";
			this.group2.Name = "group2";
			// 
			// editBox1
			// 
			this.editBox1.Label = "同类项列";
			this.editBox1.Name = "editBox1";
			this.editBox1.Text = null;
			// 
			// editBox2
			// 
			this.editBox2.Label = "并列项列";
			this.editBox2.Name = "editBox2";
			this.editBox2.Text = null;
			// 
			// button1
			// 
			this.button1.Label = "并列同类项";
			this.button1.Name = "button1";
			this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button1_Click);
			// 
			// button2
			// 
			this.button2.Label = "button2";
			this.button2.Name = "button2";
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
			this.ResumeLayout(false);

		}

		#endregion
		internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
		internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox Light;
		internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
		internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
		internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox1;
		internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox2;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
	}

	partial class ThisRibbonCollection
	{
		internal Ribbon1 Cells
		{
			get { return this.GetRibbon<Ribbon1>(); }
		}
	}
}
