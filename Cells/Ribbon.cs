using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;

// TODO:   按照以下步骤启用功能区(XML)项: 

// 1. 将以下代码块复制到 ThisAddin、ThisWorkbook 或 ThisDocument 类中。

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon();
//  }

// 2. 在此类的“功能区回调”区域中创建回调方法，以处理用户
//    操作(如单击某个按钮)。注意: 如果已经从功能区设计器中导出此功能区，
//    则将事件处理程序中的代码移动到回调方法并修改该代码以用于
//    功能区扩展性(RibbonX)编程模型。

// 3. 向功能区 XML 文件中的控制标记分配特性，以标识代码中的相应回调方法。  

// 有关详细信息，请参见 Visual Studio Tools for Office 帮助中的功能区 XML 文档。


namespace Cells
{
	[ComVisible(true)]
	public class Ribbon : Office.IRibbonExtensibility
	{
		private Office.IRibbonUI ribbon;

		public Ribbon()
		{
		}

		#region IRibbonExtensibility 成员

		public string GetCustomUI(string ribbonID)
		{
			return GetResourceText("Cells.Ribbon.xml");
		}

		#endregion

		#region 功能区回调
		//在此处创建回叫方法。有关添加回叫方法的详细信息，请访问 https://go.microsoft.com/fwlink/?LinkID=271226

		public void Ribbon_Load(Office.IRibbonUI ribbonUI)
		{
			this.ribbon = ribbonUI;

		}

		#endregion

		#region 帮助器

		private static string GetResourceText(string resourceName)
		{
			Assembly asm = Assembly.GetExecutingAssembly();
			string[] resourceNames = asm.GetManifestResourceNames();
			for (int i = 0; i < resourceNames.Length; ++i)
			{
				if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
				{
					using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
					{
						if (resourceReader != null)
						{
							return resourceReader.ReadToEnd();
						}
					}
				}
			}
			return null;
		}
		void Light_Click(IRibbonControl control, bool pressed)
		{
			Color lightColor = Color.FromArgb(34, 116, 71);
			CheckBox checkBox =(CheckBox) control;
			if (pressed)
			{
				if (System.Windows.Forms.MessageBox.Show("确认开启聚光灯功能么？这个功能将会影响撤销功能", "警告", System.Windows.Forms.MessageBoxButtons.OKCancel, System.Windows.Forms.MessageBoxIcon.Exclamation, System.Windows.Forms.MessageBoxDefaultButton.Button1) == System.Windows.Forms.DialogResult.OK)
				{
					Globals.ThisAddIn.Application.SheetSelectionChange += Application_SheetSelectionChange;
					Globals.ThisAddIn.Application.Selection.EntireRow.Interior.Color = lightColor;
					Globals.ThisAddIn.Application.Selection.EntireColumn.Interior.Color = lightColor;
				}
				else
				{
					checkBox.Value = false;
				}
			}
			else
			{
				System.Windows.Forms.MessageBox.Show("0");
				Globals.ThisAddIn.Application.Cells.Interior.ColorIndex = -4142;
			}
		}

		private void Application_SheetSelectionChange(object Sh, Range Target)
		{
			throw new NotImplementedException();
		}
		#endregion
	}
}
