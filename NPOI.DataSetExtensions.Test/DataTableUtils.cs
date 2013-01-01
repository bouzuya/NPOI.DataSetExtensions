using System;
using System.Collections.Generic;
using System.Data;

namespace NPOI.DataSetExtensions
{
	internal static class DataTableUtils
	{
		public static DataTable Create (string tableName, int columnNumber, int rowNumber)
		{
			var table = new DataTable (tableName);
			for (int i = 0; i < columnNumber; i++) {
				table.Columns.Add (string.Format ("C{0}", i + 1));
			}
			
			if (rowNumber > 0) {
				for (int rowIndex = 0; rowIndex < rowNumber; rowIndex++) {
					var values = new List<object> ();
					for (int columnIndex = 0; columnIndex < columnNumber; columnIndex++) {
						values.Add (string.Format ("R{0}C{1}", rowIndex + 1, columnIndex + 1));
					}
					table.Rows.Add (values.ToArray ());	
				}
			}
			
			return table;
		}
	}
}

