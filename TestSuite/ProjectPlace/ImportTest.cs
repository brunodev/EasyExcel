using EasyExcelDotNet.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TestSuite.ProjectPlace
{
	[TestClass]
	public class ImportTest
	{

		protected static string Path = System.IO.Path.GetFullPath("SPRINT-TMS-SPRINT-TMS-1-2016-08-13.xlsx");
		protected static EasyExcelDocument Document = new EasyExcelDocument(Path, true);

		[ClassInitialize]
		public static void AssemblyInitialize(TestContext test)
		{
			Document.Settings.AutoPrintArea = false;
		}

		[ClassCleanup]
		public static void AssemblyCleanup()
		{
			Document.Dispose();
		}

		[TestMethod]
		public void RunThroughRows()
		{
			string value;
			foreach(var row in Document.CurrentSheet.GetRows())
			{
				value = row.GetCell(2).GetValue();

				Assert.IsNotNull(value);
			}
		}

		//[TestMethod]
		//public void GetTitelRow()
		//{
		//	var titelCell = Document.CurrentSheet.GetCell("Titel");
		//	var titelRow = titelCell.Row;

		//	Assert.AreEqual(67, titelRow.GetCells().Count());
		//}
	}
}
