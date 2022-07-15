using LIBERO.Exceptions;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System;
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

		private static string ToDateFunction(string indirectAddress) => $"DATEVALUE(CONCATENATE(RIGHT({indirectAddress},4),\"/\",MID({indirectAddress},4,2),\"/\",LEFT({indirectAddress},2)))";

		private static bool IsEmptySheet(Worksheet worksheet) => worksheet.UsedRange.Address == "$A$1" && worksheet.get_Range("$A$1").Value2 == null;

		private static void ValidateSheetFormat(Worksheet worksheet, params string[] args)
		{
			var hasListObject = worksheet.ListObjects.Count > 0;
			ListObject table;

			// create table if not exist
			if (!hasListObject)
			{
				table = worksheet.ListObjects.Add(XlListObjectSourceType.xlSrcRange, Source: worksheet.UsedRange, XlListObjectHasHeaders: XlYesNoGuess.xlYes);
				table.TableStyle = "";
			}
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

		private void btnFormat_Click(object sender, RibbonControlEventArgs e)
		{
			try
			{
				// validate the sheet contains the desired table
				ValidateSheetFormat(_excel.ActiveWorkbook.ActiveSheet, "Line #", "Despatched ex-works");
				ListObject table = _excel.ActiveWorkbook.ActiveSheet.ListObjects.Item[1];

				// clear the origin conditions
				if (table.DataBodyRange.FormatConditions.Count > 0)
					table.DataBodyRange.FormatConditions.Delete();

				// format string condition
				Func<string, string> ruleDespatched = (string reference) => $"NOT(ISERR({ToDateFunction(reference)}))";
				Func<string, string> ruleNoDate = (string reference) => $"ISERR({ToDateFunction($"OFFSET({reference},0,-1)")})";
				Func<string, string, string> ruleTooLate = (string reference, string dateText) => $"{ToDateFunction($"OFFSET({reference},0,-1)")}>DATEVALUE(\"{dateText}\")";
				Func<string, string> ruleDelayed = (string reference) => $"{ToDateFunction($"OFFSET({reference},0,-1)")}>{ToDateFunction($"OFFSET({reference},0,-2)")}";

				// Append a condition column at last if not exist
				ListColumn colStatusCode = null;
				if (table.HeaderRowRange.Cells.OfType<Range>().All(x => x.Value != "Status Code"))
				{
					colStatusCode = table.ListColumns.Add(table.ListColumns.Count + 1);
					colStatusCode.Name = "Status Code";
				}
				else
				{
					colStatusCode = table.ListColumns.get_Item("Status Code");
				}

				// set status function
				var refDespatch = $"INDIRECT(\"{table.Name}[@[Despatched ex-works]]\")";
				var formulaString = $"=IF({ruleDespatched(refDespatch)},\"G\",IF({ruleNoDate(refDespatch)},\"R\",IF({ruleTooLate(refDespatch, "2022/10/1")},\"O\",IF({ruleDelayed(refDespatch)},\"Y\",\"N\"))))";
				colStatusCode.DataBodyRange.Cells[1].Formula = formulaString;

				// Highlight despached into GREEN
				FormatCondition despatchedCondition = table.DataBodyRange.FormatConditions.Add(
					XlFormatConditionType.xlExpression,
					Formula1: $"={ruleDespatched(refDespatch)}"
					);
				despatchedCondition.Interior.Color = XlRgbColor.rgbLightGreen; // GREEN

				// Highlight no date  into RED
				FormatCondition noDateCondition = table.DataBodyRange.FormatConditions.Add(
					XlFormatConditionType.xlExpression,
					Formula1: $"={ruleNoDate(refDespatch)}"
					);
				noDateCondition.Interior.Color = XlRgbColor.rgbRed; // RED

				// Highlight too late date into RED
				FormatCondition tooLateCondition = table.DataBodyRange.FormatConditions.Add(
					XlFormatConditionType.xlExpression,
					Formula1: $"={ruleTooLate(refDespatch, "2022/10/1")}"
					);
				tooLateCondition.Interior.Color = XlRgbColor.rgbOrange; // ORANGE

				// Highlight delayed  into YELLOW
				FormatCondition delayedCondition = table.DataBodyRange.FormatConditions.Add(
					XlFormatConditionType.xlExpression,
					Formula1: $"={ruleDelayed(refDespatch)}"
					);
				delayedCondition.Interior.Color = XlRgbColor.rgbYellow; // YELLOW

				_excel.ActiveWorkbook.ActiveSheet.Columns.AutoFit();
			}
			catch (Exception)
			{
				// ignore
			}
		}

		private void btnPie_Click(object sender, RibbonControlEventArgs e)
		{
			Worksheet sheet = _excel.ActiveWorkbook.Sheets.Add();
			sheet.Name = $"Statics_{DateTime.Today:yyyyMMdd}";

			var value = new object[6, 2] {
				{ "Status", "Total" },
				{ "绿色-已发货", "=SUMIF(Table1[Status Code], \"G\",Table1[Requested])" },
				{ "黄色-货物有延期", "=SUMIF(Table1[Status Code], \"Y\",Table1[Requested])" } ,
				{ "红色-没有货物", "=SUMIF(Table1[Status Code], \"R\",Table1[Requested])"},
				{ "橙色-交货期晚于10月", "=SUMIF(Table1[Status Code], \"O\",Table1[Requested])"},
				{ "无色-没有延期但尚未发货", "=SUMIF(Table1[Status Code], \"N\",Table1[Requested])"}
			};
			sheet.Cells[1, 1].resize(6, 2).Value2 = value;

			var shape = sheet.Shapes.AddChart2(251, XlChartType.xlPie);
			shape.Chart.SetSourceData(sheet.UsedRange);
			shape.Chart.FullSeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = XlRgbColor.rgbLightGreen;
			shape.Chart.FullSeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB = XlRgbColor.rgbYellow;
			shape.Chart.FullSeriesCollection(1).Points(3).Format.Fill.ForeColor.RGB = XlRgbColor.rgbRed;
			shape.Chart.FullSeriesCollection(1).Points(4).Format.Fill.ForeColor.RGB = XlRgbColor.rgbOrange;
			shape.Chart.FullSeriesCollection(1).Points(5).Format.Fill.ForeColor.RGB = XlRgbColor.rgbGray;
		}
	}
}


