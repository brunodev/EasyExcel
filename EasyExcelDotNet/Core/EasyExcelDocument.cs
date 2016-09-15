using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using EasyExcelDotNet.Modules;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace EasyExcelDotNet.Core
{
	public class EasyExcelDocument : IDisposable
    {
		#region Document
		#region Private
		private MemoryStream Stream = new MemoryStream();
		public SpreadsheetDocument Document { get; private set; }

		public int SheetIndex = 0;
		#endregion

		#region Public
		public EasyExcelSheet CurrentSheet { get; private set; }
		public EasyExcelCalculationChain CalculationChain { get; private set; }
		public EasyExcelSettings Settings { get; set; }
		#endregion
		#endregion

		#region Document parts
		public WorkbookPart WorkbookPart { get; private set; }
		public Workbook Workbook { get; private set; }
		public SharedStringTablePart SharedStringTablePart { get; private set; }
		public EasyExcelSharedStringTable sharedStringTable { get; private set; }
		
		public IEnumerable<EasyExcelSheet> Sheets { get; private set; }
		#endregion

		#region Constructors
		public EasyExcelDocument(string path, bool editAble)
		{
			Prepare(path, editAble, new EasyExcelSettings(this));
		}

		public EasyExcelDocument(string path, bool editAble, EasyExcelSettings settings)
		{
			Prepare(path, editAble, settings);
		}

		private void Prepare(string path, bool editAble, EasyExcelSettings settings)
		{
			var bytes = File.ReadAllBytes(path);
			Stream.Write(bytes, 0, bytes.Length);

			Document = SpreadsheetDocument.Open(Stream, editAble);

			WorkbookPart = Document.WorkbookPart;
			Workbook = WorkbookPart.Workbook;
			SharedStringTablePart = WorkbookPart.SharedStringTablePart;
			sharedStringTable = new EasyExcelSharedStringTable(this, SharedStringTablePart?.SharedStringTable);
			Sheets = Workbook.GetFirstChild<Sheets>().Descendants<Sheet>().Select(sheet =>
				EasyExcelSheet.Get(this, sheet));

			CurrentSheet = Sheets.ElementAt(SheetIndex);
			CalculationChain = new EasyExcelCalculationChain(this, WorkbookPart.CalculationChainPart);
		}
		#endregion

		#region Sheet
		public EasyExcelSheet NextSheet()
		{
			SheetIndex++;
			CurrentSheet = EasyExcelSheet.Get(this, SheetIndex);

			return CurrentSheet;
		}

		public EasyExcelSheet PreviousSheet()
		{
			SheetIndex--;

			CurrentSheet = EasyExcelSheet.Get(this, SheetIndex);

			return CurrentSheet;
		}

		public EasyExcelSheet SetCurrentSheet(int sheetIndex)
		{
			SheetIndex = sheetIndex;

			CurrentSheet = EasyExcelSheet.Get(this, SheetIndex);

			return CurrentSheet;
		}

		public EasyExcelSheet SetCurrentSheet(string sheetName)
		{
			CurrentSheet = Sheets.FirstOrDefault(t => t.Name.Equals("sheetName"));

			SheetIndex = Sheets.TakeWhile(sheet => sheet == CurrentSheet).Count() - 1;

			return CurrentSheet;
		}
		#endregion

		#region Build
		public byte[] Build()
		{
			PrepareBuild();
			Save();

			return Stream.ToArray();
		}

		private void PrepareBuild()
		{
			foreach(var sheet in Sheets)
			{
				DefinePrintArea(sheet);
			}
		}
		#endregion

		#region Custom settings methods
		private void DefinePrintArea(EasyExcelSheet sheet)
		{
			if(Settings.AutoPrintArea)
				sheet.DefinePrintArea();
		}
		#endregion

		public void Save()
		{
			Workbook.Save();
		}

		public void Dispose()
		{
			Document.Close();
			Document.Dispose();
			Stream.Dispose();
		}
	}
}