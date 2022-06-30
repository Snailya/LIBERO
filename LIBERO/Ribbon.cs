using LIBERO.Exceptions;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace LIBERO
{
	public partial class Ribbon
	{
		private Microsoft.Office.Interop.Excel.Application _excel;
		private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
		{
			_excel = Globals.AddIn.Application;
		}

		private void btnImport_Click(object sender, RibbonControlEventArgs e)
		{
			Worksheet destination = _excel.ActiveWorkbook.ActiveSheet, source = null;
			DateTime? date = null;
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
					openFileDialog.Filter = "All Excel Files |*.xls;*.xlsx;*.xlsm";
					openFileDialog.FilterIndex = 2;
					openFileDialog.RestoreDirectory = true;

					if (openFileDialog.ShowDialog() != DialogResult.OK) return;

					//Get the path of specified file
					filePath = openFileDialog.FileName;

					// Check if file name contains date
					var regex = new Regex(@"(\d{4})(\d{2})(\d{2})");
					var match = regex.Match(filePath);
					if (match.Success)
					{
						try
						{
							date = DateTime.ParseExact(match.Value, "yyyyMMdd", CultureInfo.InvariantCulture);
						}
						catch
						{
							// ignore
						}
					}

					if (date == null)
					{
						if (MessageBox.Show("Do not detected any string that match a date, use today instead. Click Ok to process with today's date; Click Cancel to cancel.", "", MessageBoxButtons.OKCancel) == DialogResult.Cancel) return;
						date = DateTime.Now;
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

					// Rename planned column name if exist
					if (destinationTable.HeaderRowRange.Cells.OfType<Range>().SingleOrDefault(x => x.Value == "Date Planned") != null)
					{
						destinationTable.ListColumns.get_Item("Date Planned").Name = $"Date Planned[{date:dd/MM/yyyy}]";
					}
				}
				else // Sort and copy
				{
					destinationTable = destination.ListObjects.Item[1];
					sourceTable = source.ListObjects.Item[1];

					// Overwrite"Despatched ex-works"
					var colDespatched = destinationTable.ListColumns.get_Item("Despatched ex-works");
					colDespatched.DataBodyRange.Value = sourceTable.ListColumns.get_Item("Despatched ex-works").DataBodyRange.Value;

					// Overwrite "Disp.No."
					var colDispNo = destinationTable.ListColumns.get_Item("Disp.No.");
					colDispNo.DataBodyRange.Value = sourceTable.ListColumns.get_Item("Disp.No.").DataBodyRange.Value;

					ListColumn colPlanned;
					var colPlanedName = $"Date Planned[{date:dd/MM/yyyy}]";
					if (destinationTable.HeaderRowRange.Cells.OfType<Range>().SingleOrDefault(x => x.Value == colPlanedName) != null)
					{
						if (MessageBox.Show($"There's already a column named {colPlanedName}, overwrite it? Click YES to overwrite, click NO to skip.", "Attention", MessageBoxButtons.YesNo) == DialogResult.OK)
						{
							// Overwrite "Date Planned"
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

				var refDespatch = $"INDIRECT(\"{destinationTable.Name}[@[Despatched ex-works]]\")";
				Func<string, string> toDateFunction = value => $"DATEVALUE(TEXTJOIN(\"/\", TRUE,MID({value}, 4, 2), LEFT({value}, 2), RIGHT({value}, 4)))";

				var tmp = $"=NOT(ISERR({toDateFunction(refDespatch)}))";

				FormatCondition despatchedCondition = destinationTable.DataBodyRange.FormatConditions.Add(
					XlFormatConditionType.xlExpression,
					Formula1: $"=NOT(ISERR({toDateFunction(refDespatch)}))"
					);
				despatchedCondition.Interior.ColorIndex = 4; // Green
				FormatCondition delayedCondition = destinationTable.DataBodyRange.FormatConditions.Add(
					XlFormatConditionType.xlExpression,
					Formula1: $"={toDateFunction($"OFFSET({refDespatch}, 0, -1)")}> {toDateFunction($"OFFSET({refDespatch}, 0, -2)")}"
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


