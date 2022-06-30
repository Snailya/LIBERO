﻿using LIBERO.Exceptions;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Linq;
using System.Windows.Forms;

namespace LIBERO
{
	public partial class Ribbon1
	{
		private Microsoft.Office.Interop.Excel.Application _excel;
		private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
		{
			_excel = Globals.ThisAddIn.Application;
		}

		private void btnImport_Click(object sender, RibbonControlEventArgs e)
		{
			Worksheet destination = _excel.ActiveWorkbook.ActiveSheet, source = null;

			try
			{
				var isEmpty = IsEmptySheet(destination);

				// Validate destination
				if (!isEmpty)
				{
					ValidateSheetFormat(destination, "Line #", "Despatched ex-works");
				}

				// Validate source
				var filePath = string.Empty;
				using (var openFileDialog = new OpenFileDialog())
				{
					openFileDialog.InitialDirectory = "c:\\";
					openFileDialog.Filter = "All Excel Files |*.xls;*.xlsx;*.xlsm";
					openFileDialog.FilterIndex = 2;
					openFileDialog.RestoreDirectory = true;
					if (openFileDialog.ShowDialog() == DialogResult.OK)
					{
						//Get the path of specified file
						filePath = openFileDialog.FileName;
					}
					else
					{
						return;
					}
				}

				var workbook = _excel.Workbooks.Open(filePath);
				source = workbook.Worksheets.get_Item("Items");
				ValidateSheetFormat(source, "Line #", "Date Planned", "Despatched ex-works");

				ListObject destinationTable, sourceTable;
				if (isEmpty) // Directly copy if empty
				{
					source.UsedRange.Copy(destination.get_Range("A1"));
					destinationTable = destination.ListObjects.Item[1];
				}
				else // Sort and copy
				{
					destinationTable = destination.ListObjects.Item[1];
					sourceTable = source.ListObjects.Item[1];

					// Overwrite despatched
					var colDespatched = destinationTable.ListColumns.get_Item("Despatched ex-works");
					colDespatched.DataBodyRange.Value = sourceTable.ListColumns.get_Item("Despatched ex-works").DataBodyRange.Value;

					ListColumn colPlanned;
					var colPlanedName = $"Date Planned[{DateTime.Today:dd/MM/yyyy}]";
					if (destinationTable.HeaderRowRange.Cells.OfType<Range>().SingleOrDefault(x => x.Value == colPlanedName) != null)
					{
						if (MessageBox.Show($"There's already a column named {colPlanedName}, overwrite it? Click YES to overwrite, click NO to skip.", "Attention", MessageBoxButtons.YesNo) == DialogResult.OK)
						{
							// Overwrite planned
							colPlanned = destinationTable.ListColumns.get_Item(colPlanedName);
							colPlanned.DataBodyRange.Value = sourceTable.ListColumns.get_Item("Date Planned").DataBodyRange.Value;
						} 
					}
					else
					{
						// Insert planned
						colPlanned = destinationTable.ListColumns.Add(colDespatched.Index);
						colPlanned.Name = colPlanedName;
						colPlanned.DataBodyRange.Value = sourceTable.ListColumns.get_Item("Date Planned").DataBodyRange.Value;
					}
				}

				// Format condition
				if (destinationTable.DataBodyRange.FormatConditions.Count > 0)
					destinationTable.DataBodyRange.FormatConditions.Delete();

				FormatCondition despatchedCondition = destinationTable.DataBodyRange.FormatConditions.Add(
					XlFormatConditionType.xlExpression,
					Formula1: $"=NOT(ISERR(DATEVALUE(SUBSTITUTE(INDIRECT(\"{destinationTable.Name}[@[Despatched ex-works]]\"),\".\",\" / \"))))"
					);
				despatchedCondition.Interior.ColorIndex = 4; // Green
				FormatCondition delayedCondition = destinationTable.DataBodyRange.FormatConditions.Add(
					XlFormatConditionType.xlExpression,
					Formula1: $"=DATEVALUE(SUBSTITUTE(OFFSET(INDIRECT(\"{destinationTable.Name}[@[Despatched ex-works]]\"), 0, -1),\".\",\" / \")) > DATEVALUE(SUBSTITUTE(OFFSET(INDIRECT(\"{destinationTable.Name}[@[Despatched ex-works]]\"), 0, -2),\".\",\" / \")) "
					);
				delayedCondition.Interior.ColorIndex = 6; // Yellow

				destination.Columns.AutoFit();
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
			finally
			{
				if (source != null)
					(source.Parent as Workbook).Close(false);
			}
		}

		private static bool IsEmptySheet(Worksheet worksheet) => worksheet.UsedRange.Address == "$A$1" && worksheet.get_Range("$A$1").Value2 == null;

		private static void ValidateSheetFormat(Worksheet worksheet, params string[] args)
		{
			var hasListObject = worksheet.ListObjects.Count > 0;
			ListObject table;

			// create table if not exist
			if (!hasListObject)
				table = worksheet.ListObjects.Add(
				   XlListObjectSourceType.xlSrcRange,
				   Source: worksheet.UsedRange,
				   XlListObjectHasHeaders: XlYesNoGuess.xlYes
				   );
			else
				table = worksheet.ListObjects.Item[1];

			try
			{
				foreach (var arg in args)
				{
					var col = table.ListColumns.get_Item(arg);
				}
			}
			catch
			{
				if (!hasListObject && table != null)
					table.Unlist();
				throw new InvalidExcelFileException($"Current worksheet is not a empty sheet, but at least one of {string.Join(", ", args)} column not exist. Please check the input file.");
			}
		}
	}
}


