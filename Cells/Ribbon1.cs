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
		public static Excel.Worksheet activeWorksheet;
		Excel.Worksheet lastActiveWorksheet;
		System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();

		private void Cells_Load(object sender, RibbonUIEventArgs e)
		{
			application = Globals.ThisAddIn.Application;
			Globals.ThisAddIn.Application.SheetSelectionChange += Application_SheetSelectionChange;
		}
		#region 聚光灯功能
		Color lightColor = Color.CadetBlue;
		private void Light_Click(object sender, RibbonControlEventArgs e)
		{
			if (Light.Checked)
			{
				if (MessageBox.Show("确认开启聚光灯功能么？这个功能将会影响撤销功能", "警告", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1) == DialogResult.OK)
				{
					activeWorksheet = application.ActiveSheet;
					application.Selection.EntireRow.Interior.Color = lightColor;
					application.Selection.EntireColumn.Interior.Color = lightColor;
				}
				else
				{
					Light.Checked = false;
				}
			}
			else
			{
				activeWorksheet.Cells.Interior.ColorIndex = -4142;
			}
		}


		private void Application_SheetSelectionChange(object Sh, Excel.Range Target)
		{
			if (activeWorksheet == application.ActiveSheet)
			{
				if (Light.Checked)
				{
					activeWorksheet.Cells.Interior.ColorIndex = -4142;
					Target.EntireRow.Interior.Color = lightColor;
					Target.EntireColumn.Interior.Color = lightColor;
				}
			}
			else
			{
				activeWorksheet.Cells.Interior.ColorIndex = -4142;
				activeWorksheet = application.ActiveSheet;
			}
			//以下命令为撤销时显示的命令 暂时无法完成撤销功能
			//application.OnUndo("聚光灯", "FunctionName");
		}

		#endregion
		#region 并列同类项功能
		private void Concatenation_Click(object sender, RibbonControlEventArgs e)
		{
			Excel.Worksheet worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
			Range eqrange = worksheet.Columns[editBox1.Text.ToUpper()];
			Range neqrange = worksheet.Columns[editBox2.Text.ToUpper()];
			int baseRow = 1;
			int row = 2;
			int compareColumn = eqrange.Column;
			int column = neqrange.Column;
			int deltaColumn = 1;
			Range baseCell = worksheet.Cells[baseRow, compareColumn];
			Range cell = worksheet.Cells[row, compareColumn];
			while (cell.Value != null)
			{
				try
				{
					if (cell.Value == baseCell.Value)
					{
						switch (splitChar.Text)
						{
							case "Tab(并列)":
								worksheet.Cells[baseRow, column + deltaColumn].Value = worksheet.Cells[row, column].Value;
								deltaColumn++;
								break;
							case "Space(空格)":
								worksheet.Cells[baseRow, column].Value = worksheet.Cells[baseRow, column].Value + ' ' + (worksheet.Cells[row, column].Value).ToString();
								break;
							default:
								worksheet.Cells[baseRow, column].Value = worksheet.Cells[baseRow, column].Value + splitChar.Text + (worksheet.Cells[row, column].Value).ToString();
								break;
						}
						cell.EntireRow.Delete();
						cell = worksheet.Cells[row, compareColumn];
						continue;
					}
					else
					{
						baseRow++;
						row++;
						deltaColumn = 1;
						baseCell = worksheet.Cells[baseRow, compareColumn];
						cell = worksheet.Cells[row, compareColumn];
					}
				}
				catch (Exception)
				{
					baseCell.Interior.Color = Color.Red;
					baseRow++;
					row++;
					baseCell = worksheet.Cells[baseRow, compareColumn];
					cell = worksheet.Cells[row, compareColumn];
				}
			}

		}
		#endregion

		private void Align_Click(object sender, RibbonControlEventArgs e)
		{
			activeWorksheet = application.ActiveSheet;

			Range sel = application.Selection;
			if (sel.Columns.Count > 123)
			{
				MessageBox.Show("列数过大,请重新选择区域");
				return;
			}

			application.ScreenUpdating = false;
			//application.ScreenUpdating = true;
			switch ((sender as RibbonButton).Name)
			{
				case "ctrlL":
					{
						for (int i = sel.Row; i < sel.Row + sel.Rows.Count; i++)
						{
							Range cell = activeWorksheet.Cells[i, sel.Column];
							for (int j = sel.Column; j < sel.Column + sel.Columns.Count; j++)
							{

								dynamic value = activeWorksheet.Cells[i, j].Value;
								if (value != null)
								{
									activeWorksheet.Cells[i, j].Value = null;
									cell.Value = value;
									cell = cell.Next;
								}
							}
						}
						break;
					}
				case "ctrlR":
					{
						for (int i = sel.Row; i < sel.Row + sel.Rows.Count; i++)
						{
							Range cell = activeWorksheet.Cells[i, sel.Column + sel.Columns.Count - 1];
							for (int j = sel.Column + sel.Columns.Count - 1; j >= sel.Column; j--)
							{
								dynamic value = activeWorksheet.Cells[i, j].Value;
								if (value != null)
								{
									activeWorksheet.Cells[i, j].Value = null;
									cell.Value = value;
									cell = cell.Offset[0, -1];

								}
							}
						}
						break;
					}
				default:
					break;
			}
			//application.ScreenUpdating = false;
			application.ScreenUpdating = true;
		}

		private void SameFormat(object sender, RibbonControlEventArgs e)
		{
			//application.ScreenUpdating = false;
			//application.ScreenUpdating = true;
			Range cell = application.ActiveCell;
			Range selectRange = cell.Cells;
			foreach (Range item in cell.CurrentRegion)
			{
				if (item.DisplayFormat.Font.Color == cell.DisplayFormat.Font.Color)
				{
					if (item.DisplayFormat.Interior.Color == cell.DisplayFormat.Interior.Color)
					{
						selectRange = application.Union(selectRange, item);
					}
				}
			}
			selectRange.Select();

		}
	}
}
