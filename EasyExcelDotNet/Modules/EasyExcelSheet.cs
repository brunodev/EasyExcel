using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using EasyExcelDotNet.Core;
using EasyExcelDotNet.Helpers;
using System.Collections.Generic;
using System.Linq;

namespace EasyExcelDotNet.Modules
{
	public class EasyExcelSheet : BaseModule
	{
		public Sheet Sheet { get; private set; }

		#region Sheet parts
		public WorksheetPart WorksheetPart { get; private set; }
		public Worksheet Worksheet { get; private set; }
		public SheetData SheetData { get; private set; }
		public DefinedName DefinedName { get; private set; }
		#endregion

		#region Sheet descendants
		private IEnumerable<Cell> Cells;
		private IEnumerable<Row> Rows;
		#endregion

		#region readonly
		private readonly string FirstCell = "$A$1";
		#endregion

		#region Properties
		public string Name
		{
			get { return Sheet.Name; }
			set { Sheet.Name = value; }
		}
		#endregion

		private EasyExcelSheet(EasyExcelDocument document, Sheet sheet) : base(document)
		{
			Sheet = sheet;

			WorksheetPart = Document.WorkbookPart.GetPartById(Sheet.Id) as WorksheetPart;
			Worksheet = WorksheetPart.Worksheet;
			SheetData = Worksheet.Descendants<SheetData>().SingleOrDefault();

			var definedNames = Document.Workbook.Descendants<DefinedNames>();

			DefinedName = definedNames?.FirstOrDefault()?
				.Descendants<DefinedName>().FirstOrDefault(definedName =>
					definedName.Text.StartsWith(Name)
					|| definedName.Text.StartsWith("'" + Name));

			Cells = SheetData.Descendants<Cell>();
			Rows = SheetData.Descendants<Row>();
		}

		public static EasyExcelSheet Get(EasyExcelDocument document, int sheetIndex)
		{
			return document.Sheets.ElementAt(sheetIndex);
		}

		public static EasyExcelSheet Get(EasyExcelDocument document, Sheet sheet)
		{
			return new EasyExcelSheet(document, sheet);
		}

		#region Cells
		public IEnumerable<EasyExcelCell> GetCells()
		{
			return Cells.Select(cell => EasyExcelCell.Get(Document, cell, GetRowByCell(cell)));
		}

		public IEnumerable<EasyExcelCell> GetCells(string value)
		{
			int? sharedStringItemIndex = Document.sharedStringTable.GetIndex(value);

			if (sharedStringItemIndex == null)
				return null;

			var cells = SheetData.Descendants<Cell>()
				.Where(t =>
					t.InnerText.Equals(sharedStringItemIndex.Value.ToString())
					&& (t.DataType != null && t.DataType.HasValue && t.DataType.Value == CellValues.SharedString));

			return cells.Select(cell => EasyExcelCell.Get(Document, cell, GetRowByCell(cell)));
		}

		public EasyExcelCell GetCell(string value)
		{
			return GetCells(value).FirstOrDefault();
		}

		#endregion

		#region Rows
		#region Get
		public IEnumerable<EasyExcelRow> GetRows()
		{
			return Rows.Select(row => EasyExcelRow.Get(Document, row));
		}

		public EasyExcelRow GetRow(int index)
		{
			return EasyExcelRow.Get(Document, Rows.ElementAt(index));
		}

		public EasyExcelRow GetLastRow()
		{
			return EasyExcelRow.Get(Document, Rows.LastOrDefault());
		}

		public EasyExcelRow GetRowByCell(EasyExcelCell cell)
		{
			return EasyExcelRow.Get(Document, Rows.ElementAt(cell.ReferenceDigits - 1));
		}

		public EasyExcelRow GetRowByCell(Cell cell)
		{
			return EasyExcelRow.Get(Document, Rows.ElementAt(cell.CellReference.Value.GetDigits() - 1));
		}
		#endregion

		#region Insert
		public void InsertRow(int index, IEnumerable<EasyExcelCell> cells)
		{

		}
		#endregion
		#endregion

		#region Columns
		public string GetLastColumnLetter()
		{
			var cellLetters = Cells.Select(t => t.CellReference.InnerText.GetLetters());
			var orderedCells = cellLetters.OrderByDescending(cell => cell.Length)
				.ThenByDescending(cell => cell.GetLetters());

			return orderedCells.FirstOrDefault();
		}
		#endregion
		#region Print area
		private DefinedName NewPrintArea(uint localSheetId, string text)
		{
			return new DefinedName()
			{
				Name = "_xlnm.Print_Area",
				LocalSheetId = localSheetId,
				Text = text
			};
		}
		public void DefinePrintArea()
		{
			//uint localSheetID = DefinedName.LocalSheetId;

			string lastCell = string.Format("${0}${1}", GetLastColumnLetter(), GetLastRow().Index);

			DefinedName.Text = string.Format("'{0}'!{1}:{2}", Name, FirstCell, lastCell);

			//var definedNames = Document.Workbook.Descendants<DefinedNames>().FirstOrDefault();
			//var newPrintArea = NewPrintArea(localSheetID, string.Format("'{0}'!{1}:{2}", Name, FirstCell, lastCell));
			//definedNames.Append(newPrintArea);
		}
		#endregion

	}
}
