using DocumentFormat.OpenXml.Spreadsheet;
using EasyExcelDotNet.Core;
using System.Collections.Generic;
using System.Linq;

namespace EasyExcelDotNet.Modules
{
	public class EasyExcelRow : BaseModule
	{
		public Row Row { get; private set; }

		private IEnumerable<Cell> Cells;

		#region Properties
		public int Index
		{
			get { return (int)Row.RowIndex.Value; }
		}
		#endregion

		private EasyExcelRow(EasyExcelDocument document, Row row) : base(document)
		{
			Row = row;
			Cells = Row.Descendants<Cell>();
		}

		public EasyExcelRow() : base (null)
		{

		}

		#region Get static methods
		public static EasyExcelRow Get(EasyExcelDocument document, Row row)
		{
			return new EasyExcelRow(document, row);
		}
		#endregion

		#region Cells
		public IEnumerable<EasyExcelCell> GetCells()
		{
			return Cells.Select(cell => EasyExcelCell.Get(Document, cell, this));
		}

		public EasyExcelCell GetCell(int index)
		{
			return EasyExcelCell.Get(Document, Row.Descendants<Cell>().ElementAt(index), this);
		}
		#endregion

		#region Insert
		public void AddCell(EasyExcelCell cell)
		{
			Cells.ToList().Add(cell.Cell);
		}
		#endregion
	}
}
