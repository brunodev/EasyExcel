using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;

namespace TestSuite
{
	[TestClass]
	public class CellTest : TestHeader
	{
		[TestMethod]
		public void RemoveCalculationChain()
		{
			Document.CalculationChain.Remove();
		}

		[TestMethod]
		public void GetCellValue()
		{
			string value = Document.CurrentSheet.GetCell("Text3").GetValue();

			Assert.AreEqual("Text3", value);
		}

		[TestMethod]
		public void GetCells()
		{
			var cells = Document.CurrentSheet.GetCells("Text1");

			Assert.AreEqual(2, cells.Count());
		}

		[TestMethod]
		public void GetLastColumn()
		{
			var lastColumn = Document.CurrentSheet.GetLastColumnLetter();

			Assert.AreEqual("Q", lastColumn);
		}

	}
}
