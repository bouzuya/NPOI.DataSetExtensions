using System;
using System.Data;

namespace NPOI.DataSetExtensions
{
	public static class DataTableExtensions
	{
		public static void WriteXls (this DataTable dataTable, string fileName)
		{
			if (dataTable == null) {
				throw new NullReferenceException ();
			}

			XlsWriter.Write (dataTable, fileName);
		}
	}
}

