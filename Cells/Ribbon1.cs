using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Drawing;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace Cells
{
	public partial class Ribbon1
	{
		Excel.Application application;
		private void Cells_Load(object sender, RibbonUIEventArgs e)
		{
			application = Globals.ThisAddIn.Application;
		}

		private void Light_Click(object sender, RibbonControlEventArgs e)
		{
			if (Light.Checked)
			{
				if (MessageBox.Show("确认开启聚光灯功能么？这个功能将会影响撤销功能", "警告", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1) == DialogResult.OK)
				{
					Globals.ThisAddIn.Application.SheetSelectionChange += Application_SheetSelectionChange;
					Application_SheetSelectionChange(new object(), application.ActiveCell);
				}
				else
				{
					Light.Checked = false;
					Globals.ThisAddIn.Application.SheetSelectionChange -= Application_SheetSelectionChange;
				}
			}
			else
			{
				Excel.Worksheet worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
				worksheet.Cells.Interior.ColorIndex = -4142;
				Globals.ThisAddIn.Application.SheetSelectionChange -= Application_SheetSelectionChange;
			}
		}

		private void Application_SheetSelectionChange(object Sh, Excel.Range Target)
		{
			Excel.Worksheet worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
			worksheet.Cells.Interior.ColorIndex = -4142;
			Target.EntireRow.Interior.Color = Color.CadetBlue;
			Target.EntireColumn.Interior.Color = Color.CadetBlue;
			//application.OnUndo("聚光灯", "LightOnUnDo");
		}

		private void Button1_Click(object sender, RibbonControlEventArgs e)
		{
			int row = 2;
			int columns = 1;
			Excel.Worksheet worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
			Range eqrange = worksheet.Columns[editBox1.Text.ToUpper()];
			Range neqrange = worksheet.Columns[editBox2.Text.ToUpper()];
			Range eq0, eq1, neq0, neq1;
			eq0 = eqrange.Rows[1];
			neq0 = neqrange.Rows[1];
			eq1 = eqrange.Rows[2];
			neq1 = neqrange.Rows[2];
			while (eq1.Value != null)
			{
				if (eq1.Value == eq0.Value)
				{
					worksheet.Cells[neq0.Row, neq1.Column + columns].Value = neq1.Value;
					columns++;
					worksheet.Rows[row].Delete(XlDeleteShiftDirection.xlShiftUp);
					eq1 = eqrange.Rows[row];
					neq1 = neqrange.Rows[row];
					continue;
				}
				else
				{
					columns = 1;
					eq0 = eq1.Cells;
					neq0 = neq1.Cells;
				}
				row++;
				eq1 = eqrange.Rows[row];
				neq1 = neqrange.Rows[row];
			}
		}
	}
}
