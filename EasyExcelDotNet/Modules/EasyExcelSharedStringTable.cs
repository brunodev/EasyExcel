using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EasyExcelDotNet.Core;
using DocumentFormat.OpenXml.Spreadsheet;

namespace EasyExcelDotNet.Modules
{
	public class EasyExcelSharedStringTable : BaseModule
	{
		public SharedStringTable SharedStringTable { get; private set; }

		public EasyExcelSharedStringTable(EasyExcelDocument document, SharedStringTable sharedStringTable) : base(document)
		{
			SharedStringTable = sharedStringTable;
		}

		public string GetValue(int index)
		{
			return SharedStringTable.Descendants<SharedStringItem>().ElementAt(index).InnerText;
		}

		public int? GetIndex(string value)
		{
			int z = 0;

			foreach (var sharedStringItem in GetSharedStringItems())
			{
				if (sharedStringItem.InnerText.Contains(value))
					return z;

				z++;
			}

			return null;
		}

		#region SharedStringItem
		private IEnumerable<SharedStringItem> GetSharedStringItems()
		{
			return SharedStringTable.Descendants<SharedStringItem>();
		}
			#endregion
	}
}
