using System;
using System.Linq;
using System.Data;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;

namespace NPOI.DataSetExtensions
{
	internal static class XlsWriter
	{
		private static readonly int XlsMaxColumns = 256;
		private static readonly int XlsMaxRows = 65536;
		
		internal static void Write (DataSet dataSet, string fileName)
		{
			ValidateDataSet (dataSet);
			using (var fileStream = File.OpenWrite(fileName)) {
				var workbook = new HSSFWorkbook ();
				ImportFromDataSet (workbook, dataSet);
				workbook.Write (fileStream);
			}
		}

		internal static void Write (DataTable dataTable, string fileName)
		{
			ValidateDataTable (dataTable);
			using (var fileStream = File.OpenWrite(fileName)) {
				var workbook = new HSSFWorkbook ();
				ImportFromDataTable (workbook, dataTable);
				workbook.Write (fileStream);
			}
		}

		private static void ValidateDataSet (DataSet dataSet)
		{
			if (dataSet.Tables.Count == 0) {
				throw new InvalidOperationException ("DataSet.Tables.Count == 0");
			}

			foreach (var table in dataSet.Tables.Cast<DataTable>()) {
				ValidateDataTable (table);
			}
		}
		
		private static void ValidateDataTable (DataTable dataTable)
		{
			if (dataTable.Columns.Count > XlsMaxColumns) {
				throw new InvalidOperationException (
					string.Format ("table.Columns.Count > {0}", XlsMaxColumns));
			}
			
			if (dataTable.Rows.Count > XlsMaxRows) {
				throw new InvalidOperationException (
					string.Format ("table.Rows.Count > {0}", XlsMaxRows));
			}
		}

		private static void ImportFromDataSet (IWorkbook workbook, DataSet dataSet)
		{
			foreach (var dataTable in dataSet.Tables.Cast<DataTable>()) {
				ImportFromDataTable (workbook, dataTable);	
			}
		}
		
		private static void ImportFromDataTable (IWorkbook workbook, DataTable table)
		{
			var sheet = GetOrCreateSheet (workbook, table.TableName);
			var xlsRowIndex = 0;
			foreach (var dataRow in table.Rows.Cast<DataRow>()) {
				var xlsRow = GetOrCreateRow (sheet, xlsRowIndex);
				foreach (var dataColumn in table.Columns.Cast<DataColumn>()) {
					var xlsCell = GetOrCreateCell (xlsRow, dataColumn.Ordinal);
					SetCellValue (xlsCell, dataColumn.DataType, dataRow [dataColumn]);
				}
				xlsRowIndex++;
			}
		}
		
		private static ISheet GetOrCreateSheet (IWorkbook workbook, string sheetName)
		{
			var defaultSheetName = string.Format ("Sheet {0}", workbook.NumberOfSheets + 1);
			return workbook.CreateSheet (string.IsNullOrEmpty (sheetName) ? defaultSheetName : sheetName);
		}

		private static IRow GetOrCreateRow (ISheet sheet, int rowIndex)
		{
			var row = sheet.GetRow (rowIndex);
			return row ?? sheet.CreateRow (rowIndex);
		}
		
		private static ICell GetOrCreateCell (ISheet sheet, int cellIndex, int rowIndex)
		{
			var row = GetOrCreateRow (sheet, rowIndex);
			return GetOrCreateCell (row, cellIndex);
		}
		
		private static ICell GetOrCreateCell (IRow row, int cellIndex)
		{
			var cell = row.GetCell (cellIndex);
			return cell ?? row.CreateCell (cellIndex);
		}
		
		private static void SetCellValue (ICell cell, Type type, object value)
		{
			if (type == typeof(string)) {
				cell.SetCellValue ((string)value);
			} else if (type == typeof(bool)) {
				cell.SetCellValue ((bool)value);
			} else if (type == typeof(double)) {
				cell.SetCellValue ((double)value);
			} else if (type == typeof(DateTime)) {
				cell.SetCellValue ((DateTime)value);
			} else if (type == typeof(IRichTextString)) {
				cell.SetCellValue ((IRichTextString)value);
			} else {
				cell.SetCellValue (value == null ? string.Empty : value.ToString ());
			}
		}
	}
}

