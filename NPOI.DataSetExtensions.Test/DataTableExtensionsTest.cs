using NUnit.Framework;
using System;
using System.Data;
using System.Collections.Generic;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.DataSetExtensions;

namespace NPOI.DataSetExtensions
{
	[TestFixture()]
	public class DataTableExtensionsTest
	{
		private string _fileName;
		
		[SetUp()]
		public void SetUp ()
		{
			this._fileName = Path.GetTempFileName ();
		}
		
		[TearDown()]
		public void TearDown ()
		{
			File.Delete (this._fileName);
		}
		
		private long Performance (string path, int columnNumber, int rowNumber, int tryCount)
		{
			var table = DataTableUtils.Create ("Sheet 1", columnNumber, rowNumber);
			var sw = new System.Diagnostics.Stopwatch ();
			for (int i = 0; i < tryCount; i++) {
				GC.Collect ();
				sw.Start ();
				table.WriteXls (path);
				sw.Stop ();
			}
			
			return sw.ElapsedMilliseconds / tryCount;
		}
		
		private HSSFWorkbook OpenWorkbook (string fileName)
		{
			return new HSSFWorkbook (File.OpenRead (fileName));
		}
		
		[Test()]
		public void ファイルが存在すること ()
		{
			var table = DataTableUtils.Create ("Sheet 1", 0, 0);
			table.WriteXls (this._fileName);
			Assert.That (File.Exists (this._fileName), Is.True);
		}
		
		[Test()]
		public void ファイルが読み取り専用ならUnauthorizedAccessExceptionを投げること ()
		{
			var fileInfo = new FileInfo (this._fileName);
			fileInfo.IsReadOnly = true;
			var table = DataTableUtils.Create ("Sheet 1", 0, 0);
			Assert.Throws<UnauthorizedAccessException> (() => table.WriteXls (this._fileName));
			fileInfo.IsReadOnly = false;
		}
		
		[Test()]
		public void ファイルが0バイトより大きいこと ()
		{
			var table = DataTableUtils.Create ("Sheet 1", 0, 0);
			table.WriteXls (this._fileName);
			Assert.That (new FileInfo (this._fileName).Length, Is.GreaterThan (0));
		}
		
		[Test()]
		public void ファイルを開きなおして例外が投げられないこと ()
		{
			var table = DataTableUtils.Create ("Sheet 1", 0, 0);
			table.WriteXls (this._fileName);
			Assert.DoesNotThrow (() => OpenWorkbook (this._fileName));
		}
		
		[Test()]
		public void ファイルの先頭シート1列1行にR1C1が書き込まれること ()
		{
			var fileName = Path.GetTempFileName ();
			var table = DataTableUtils.Create ("Sheet 1", 1, 1);
			table.WriteXls (this._fileName);
			var workbook = OpenWorkbook (this._fileName);
			Assert.That (workbook.GetSheetAt (0).GetRow (0).GetCell (0).StringCellValue, Is.EqualTo ("R1C1"));
		}
		
		[Test()]
		public void ファイルの先頭シート2列1行にR1C2が書き込まれること ()
		{
			var fileName = Path.GetTempFileName ();
			var table = DataTableUtils.Create ("Sheet 1", 2, 1);
			table.WriteXls (this._fileName);
			var workbook = OpenWorkbook (this._fileName);
			Assert.That (workbook.GetSheetAt (0).GetRow (0).GetCell (1).StringCellValue, Is.EqualTo ("R1C2"));
		}
		
		[Test()]
		public void ファイルの先頭シート256列1行にR1C256が書き込まれること ()
		{
			var table = DataTableUtils.Create ("Sheet 1", 256, 1);
			table.WriteXls (this._fileName);
			var workbook = OpenWorkbook (this._fileName);
			Assert.That (workbook.GetSheetAt (0).GetRow (0).GetCell (255).StringCellValue, Is.EqualTo ("R1C256"));
		}
		
		[Test()]
		public void DataTableが257列のときInvalidOperationExceptionを投げること ()
		{
			var table = DataTableUtils.Create ("Sheet 1", 257, 1);
			Assert.Throws<InvalidOperationException> (() => table.WriteXls (this._fileName));
		}
		
		[Test()]
		public void ファイルの先頭シート1列2行にR2C1が書き込まれること ()
		{
			var table = DataTableUtils.Create ("Sheet 1", 1, 2);
			table.WriteXls (this._fileName);
			var workbook = OpenWorkbook (this._fileName);
			Assert.That (workbook.GetSheetAt (0).GetRow (1).GetCell (0).StringCellValue, Is.EqualTo ("R2C1"));
		}
		
		[Test()]
		public void ファイルの先頭シート1列65536行にA65536C1が出力されること ()
		{
			var table = DataTableUtils.Create ("Sheet 1", 1, 65536);
			table.WriteXls (this._fileName);
			var workbook = OpenWorkbook (this._fileName);
			Assert.That (workbook.GetSheetAt (0).GetRow (65535).GetCell (0).StringCellValue, Is.EqualTo ("R65536C1"));
		}
		
		[Test()]
		public void DataTableが65537行のときInvalidOperationExceptionを投げること ()
		{
			var fileName = Path.GetTempFileName ();
			var table = DataTableUtils.Create ("Sheet 1", 1, 65537);
			Assert.Throws<InvalidOperationException> (() => table.WriteXls (this._fileName));
		}
		
		[Ignore("OutOfMemoryException")]
		[TestCase(1000)]
		[TestCase(10000)]
		[TestCase(50000)]
		[TestCase(60000)]
		[TestCase(65536)]
		public void ファイルの先頭シート1列n行で例外を投げないこと (int rowNumber)
		{
			var fileName = Path.GetTempFileName ();
			var count = 100; // 試行回数
			var expected = 1000; // 1回あたりの平均時間(ms)
			Assert.That (Performance (fileName, 1, rowNumber, count), Is.LessThan (expected), rowNumber.ToString ());
		}
		
		[Ignore("OutOfMemoryException")]
		[Test()]
		public void ファイルの先頭シート256列65536行に文字が出力されること ()
		{
			var fileName = Path.GetTempFileName ();
			var table = DataTableUtils.Create ("Sheet 1", 256, 65536);
			table.WriteXls (this._fileName);
			var workbook = OpenWorkbook (this._fileName);
			Assert.That (workbook.GetSheetAt (0).GetRow (65536 - 1).GetCell (256 - 1).StringCellValue, Is.EqualTo ("R65536C256"));
		}
		
		[Test()]
		public void 書き込んだBooleanをBooleanCellValueで取得できること ()
		{
			var fileName = Path.GetTempFileName ();
			var table = new DataTable ("Sheet 1");
			table.Columns.Add ("C1", typeof(bool));
			table.Rows.Add (true);
			table.WriteXls (this._fileName);
			var workbook = OpenWorkbook (this._fileName);
			var cell = workbook.GetSheetAt (0).GetRow (0).GetCell (0);
			Assert.That (cell.CellType, Is.EqualTo (NPOI.SS.UserModel.CellType.BOOLEAN));
			Assert.That (cell.BooleanCellValue, Is.True);
		}
		
		[Test()]
		public void 書き込んだDateTimeをDateCellValueで取得できること ()
		{
			var today = DateTime.Today;
			var fileName = Path.GetTempFileName ();
			var table = new DataTable ("Sheet 1");
			table.Columns.Add ("C1", typeof(DateTime));
			table.Rows.Add (today);
			table.WriteXls (this._fileName);
			var workbook = OpenWorkbook (this._fileName);
			var cell = workbook.GetSheetAt (0).GetRow (0).GetCell (0);
			Assert.That (cell.CellType, Is.EqualTo (NPOI.SS.UserModel.CellType.NUMERIC));
			Assert.That (cell.DateCellValue, Is.EqualTo (today));
		}
		
		[Test()]
		public void 書き込んだDateTimeNowをDateCellValueで取得できること ()
		{
			var now = DateTime.UtcNow;
			var fileName = Path.GetTempFileName ();
			var table = new DataTable ("Sheet 1");
			table.Columns.Add ("C1", typeof(DateTime));
			table.Rows.Add (now);
			table.WriteXls (this._fileName);
			var workbook = OpenWorkbook (this._fileName);
			var cell = workbook.GetSheetAt (0).GetRow (0).GetCell (0);
			Assert.That (cell.CellType, Is.EqualTo (NPOI.SS.UserModel.CellType.NUMERIC));
			// EqualTo だと失敗する (内部的に double で保持しているためと考える
			Assert.That (cell.DateCellValue.ToString (), Is.EqualTo (now.ToString ()));
		}
		
		[Test()]
		public void 書き込んだDoubleをNumericCellValueで取得できること ()
		{
			var expected = 1.2D;
			var fileName = Path.GetTempFileName ();
			var table = new DataTable ("Sheet 1");
			table.Columns.Add ("C1", typeof(double));
			table.Rows.Add (expected);
			table.WriteXls (this._fileName);
			var workbook = OpenWorkbook (this._fileName);
			var cell = workbook.GetSheetAt (0).GetRow (0).GetCell (0);
			Assert.That (cell.CellType, Is.EqualTo (NPOI.SS.UserModel.CellType.NUMERIC));
			Assert.That (cell.NumericCellValue, Is.EqualTo (expected));
		}
	}
}

