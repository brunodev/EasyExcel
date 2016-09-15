using EasyExcelDotNet.Core;

namespace EasyExcelDotNet.Modules
{
	public class BaseModule
	{
		protected EasyExcelDocument Document;

		protected BaseModule(EasyExcelDocument document)
		{
			Document = document;
		}
	}
}
