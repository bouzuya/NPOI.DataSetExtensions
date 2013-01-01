using System;
using System.Data;

namespace NPOI.DataSetExtensions
{
	public static class DataSetExtensions
	{
		public static void WriteXls (this DataSet dataSet, string fileName)
		{
			if (dataSet == null) {
				throw new NullReferenceException ();
			}

			XlsWriter.Write (dataSet, fileName);
		}
	}
}

