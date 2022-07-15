namespace LIBERO
{
	partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		public Ribbon()
			: base(Globals.Factory.GetRibbonFactory())
		{
			InitializeComponent();
		}

		/// <summary> 
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Component Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.tabLibero = this.Factory.CreateRibbonTab();
			this.grpSiemens = this.Factory.CreateRibbonGroup();
			this.btnImport = this.Factory.CreateRibbonButton();
			this.btnFormat = this.Factory.CreateRibbonButton();
			this.btnPie = this.Factory.CreateRibbonButton();
			this.tabLibero.SuspendLayout();
			this.grpSiemens.SuspendLayout();
			this.SuspendLayout();
			// 
			// tabLibero
			// 
			this.tabLibero.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
			this.tabLibero.Groups.Add(this.grpSiemens);
			this.tabLibero.Label = "LIBERO";
			this.tabLibero.Name = "tabLibero";
			// 
			// grpSiemens
			// 
			this.grpSiemens.Items.Add(this.btnImport);
			this.grpSiemens.Items.Add(this.btnFormat);
			this.grpSiemens.Items.Add(this.btnPie);
			this.grpSiemens.Label = "西门子";
			this.grpSiemens.Name = "grpSiemens";
			// 
			// btnImport
			// 
			this.btnImport.Label = "导入清单";
			this.btnImport.Name = "btnImport";
			this.btnImport.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnImport_Click);
			// 
			// btnFormat
			// 
			this.btnFormat.Label = "格式化";
			this.btnFormat.Name = "btnFormat";
			this.btnFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormat_Click);
			// 
			// btnPie
			// 
			this.btnPie.Label = "统计图表";
			this.btnPie.Name = "btnPie";
			this.btnPie.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPie_Click);
			// 
			// Ribbon
			// 
			this.Name = "Ribbon";
			this.RibbonType = "Microsoft.Excel.Workbook";
			this.Tabs.Add(this.tabLibero);
			this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
			this.tabLibero.ResumeLayout(false);
			this.tabLibero.PerformLayout();
			this.grpSiemens.ResumeLayout(false);
			this.grpSiemens.PerformLayout();
			this.ResumeLayout(false);

		}

		#endregion

		internal Microsoft.Office.Tools.Ribbon.RibbonTab tabLibero;
		internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpSiemens;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton btnImport;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormat;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPie;
	}

	partial class ThisRibbonCollection
	{
		internal Ribbon Ribbon1
		{
			get { return this.GetRibbon<Ribbon>(); }
		}
	}
}
