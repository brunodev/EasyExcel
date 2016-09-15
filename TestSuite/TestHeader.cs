using EasyExcelDotNet.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;

namespace TestSuite
{
	[TestClass]
	public class TestHeader
	{
		protected static string Path = System.IO.Path.GetFullPath("TestDocument.xlsx");
		protected static EasyExcelDocument Document = new EasyExcelDocument(Path, true);

		[ClassInitialize]
		public static void AssemblyInitialize(TestContext test)
		{
			Document.Settings.AutoPrintArea = true;
		}

		[ClassCleanup]
		public static void AssemblyCleanup()
		{
			File.WriteAllBytes("ParsedTestDocument.xlsx", Document.Build());

			Document.Dispose();
		}
	}
}
