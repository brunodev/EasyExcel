using DocumentFormat.OpenXml.Spreadsheet;
using EasyExcelDotNet.Core;
using EasyExcelDotNet.Helpers;
using System;

namespace EasyExcelDotNet.Modules
{
	public class EasyExcelCell : BaseModule
	{
		public Cell Cell { get; private set; }
		public EasyExcelRow Row { get; private set; }

		#region Properties
		public string CellReference
		{
			get { return Cell.CellReference; }
			set { Cell.CellReference = value; }
		}

		public string ReferenceLetters { get { return CellReference.GetLetters(); } }
		public int ReferenceDigits { get { return CellReference.GetDigits(); } }
		#endregion

		private EasyExcelCell(EasyExcelDocument document, Cell cell, EasyExcelRow row) : base(document)
		{
			Cell = cell;
			Row = row;
		}

		public static EasyExcelCell Get(EasyExcelDocument document, Cell cell, EasyExcelRow row)
		{
			return new EasyExcelCell(document, cell, row);
		}

		/*
		 * TODO: Check if DataType.HasValue has any use,
		 * and what the consequence of false is
		 */
		public string GetValue()
		{
			if ((Cell.DataType == null) || (Cell.DataType != null && !Cell.DataType.HasValue))
				return Cell.InnerText;

			switch (Cell.DataType.Value)
			{
				case CellValues.SharedString:
					return Document.sharedStringTable.GetValue(Convert.ToInt32(Cell.CellValue.InnerText));
				default:
					return Cell.InnerText;
			}
		}

		

	}

}