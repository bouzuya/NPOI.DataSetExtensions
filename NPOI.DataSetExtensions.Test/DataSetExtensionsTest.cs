using NUnit.Framework;
using System;
using System.Data;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.DataSetExtensions;

namespace NPOI.DataSetExtensions
{
	public class DataSetExtensionsTest
	{
		[TestFixture()]
		public class DataTableが存在しないの場合
		{
			private string _fileName;
			private DataSet _dataSet;
			
			[SetUp()]
			public void SetUp ()
			{
				var fileName = Path.GetTempFileName ();
				var dataSet = new DataSet ();
				this._fileName = fileName;
				this._dataSet = dataSet;
			}
			
			[TearDown()]
			public void TearDown ()
			{
				File.Delete (this._fileName);
			}
			
			[Test()]
			public void InvalidOperationExceptionを投げる ()
			{
				Assert.Throws<InvalidOperationException>(
					() => this._dataSet.WriteXls (this._fileName));
			}
		}

		[TestFixture()]
		public class DataTableがひとつの場合
		{
			private string _fileName;
			private DataSet _dataSet;
			
			[SetUp()]
			public void SetUp ()
			{
				var fileName = Path.GetTempFileName ();
				var dataSet = new DataSet ();
				dataSet.Tables.Add (DataTableUtils.Create ("table1", 10, 10));
				this._fileName = fileName;
				this._dataSet = dataSet;
			}
			
			[TearDown()]
			public void TearDown ()
			{
				File.Delete (this._fileName);
			}
			
			[Test()]
			public void NumberOfSheetsが1を返す ()
			{
				this._dataSet.WriteXls (this._fileName);
				var workbook = new HSSFWorkbook (File.OpenRead (this._fileName));
				Assert.That (workbook.NumberOfSheets, Is.EqualTo (1));
			}
			
			[Test()]
			public void 各シートのシート名がそれぞれtable1である ()
			{
				this._dataSet.WriteXls (this._fileName);
				var workbook = new HSSFWorkbook (File.OpenRead (this._fileName));
				Assert.That (workbook.GetSheetAt (0).SheetName, Is.EqualTo ("table1"));
			}
			
			[Test()]
			public void シート1のセルR1C1がR1C1である ()
			{
				this._dataSet.WriteXls (this._fileName);
				var workbook = new HSSFWorkbook (File.OpenRead (this._fileName));
				Assert.That (workbook.GetSheetAt (0).GetRow (0).GetCell (0).StringCellValue, Is.EqualTo ("R1C1"));
			}
			
			[Test()]
			public void シート1のセルR10C10がR10C10である ()
			{
				this._dataSet.WriteXls (this._fileName);
				var workbook = new HSSFWorkbook (File.OpenRead (this._fileName));
				Assert.That (workbook.GetSheetAt (0).GetRow (9).GetCell (9).StringCellValue, Is.EqualTo ("R10C10"));
			}
		}

		[TestFixture()]
		public class DataTableがふたつの場合
		{
			private string _fileName;
			private DataSet _dataSet;
			
			[SetUp()]
			public void SetUp ()
			{
				var fileName = Path.GetTempFileName ();
				var dataSet = new DataSet ();
				dataSet.Tables.Add (DataTableUtils.Create ("table1", 10, 10));
				dataSet.Tables.Add (DataTableUtils.Create ("table2", 5, 5));
				this._fileName = fileName;
				this._dataSet = dataSet;
			}
			
			[TearDown()]
			public void TearDown ()
			{
				File.Delete (this._fileName);
			}

			[Test()]
			public void NumberOfSheetsが2を返す ()
			{
				this._dataSet.WriteXls (this._fileName);
				var workbook = new HSSFWorkbook (File.OpenRead (this._fileName));
				Assert.That (workbook.NumberOfSheets, Is.EqualTo (2));
			}

			[Test()]
			public void 各シートのシート名がそれぞれtable1とtable2である ()
			{
				this._dataSet.WriteXls (this._fileName);
				var workbook = new HSSFWorkbook (File.OpenRead (this._fileName));
				Assert.That (workbook.GetSheetAt (0).SheetName, Is.EqualTo ("table1"));
				Assert.That (workbook.GetSheetAt (1).SheetName, Is.EqualTo ("table2"));
			}
			
			[Test()]
			public void シート1のセルR1C1がR1C1である ()
			{
				this._dataSet.WriteXls (this._fileName);
				var workbook = new HSSFWorkbook (File.OpenRead (this._fileName));
				Assert.That (workbook.GetSheetAt (0).GetRow (0).GetCell (0).StringCellValue, Is.EqualTo ("R1C1"));
			}

			[Test()]
			public void シート2のセルR1C1がR1C1である ()
			{
				this._dataSet.WriteXls (this._fileName);
				var workbook = new HSSFWorkbook (File.OpenRead (this._fileName));
				Assert.That (workbook.GetSheetAt (1).GetRow (0).GetCell (0).StringCellValue, Is.EqualTo ("R1C1"));
			}

			[Test()]
			public void シート1のセルR10C10がR10C10である ()
			{
				this._dataSet.WriteXls (this._fileName);
				var workbook = new HSSFWorkbook (File.OpenRead (this._fileName));
				Assert.That (workbook.GetSheetAt (0).GetRow (9).GetCell (9).StringCellValue, Is.EqualTo ("R10C10"));
			}
			
			[Test()]
			public void シート2のセルR5C5がR5C5である ()
			{
				this._dataSet.WriteXls (this._fileName);
				var workbook = new HSSFWorkbook (File.OpenRead (this._fileName));
				Assert.That (workbook.GetSheetAt (1).GetRow (4).GetCell (4).StringCellValue, Is.EqualTo ("R5C5"));
			}
		}
	}
}

