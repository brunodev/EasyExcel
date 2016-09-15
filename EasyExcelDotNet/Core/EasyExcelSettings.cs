namespace EasyExcelDotNet.Core
{
	public class EasyExcelSettings
	{

		private EasyExcelDocument Document { get; set; }

		public EasyExcelSettings(EasyExcelDocument document)
		{
			Document = document;
		}

		#region OpenXML settings
		public bool ForceFullCalculation
		{
			get { return Document.Workbook.CalculationProperties.ForceFullCalculation; }
			set { Document.Workbook.CalculationProperties.ForceFullCalculation = value; }
		}

		public bool FullCalculationOnLoad
		{
			get { return Document.Workbook.CalculationProperties.FullCalculationOnLoad; }
			set { Document.Workbook.CalculationProperties.FullCalculationOnLoad = value; }
		}

		public bool CalculationOnSave
		{
			get { return Document.Workbook.CalculationProperties.CalculationOnSave; }
			set { Document.Workbook.CalculationProperties.CalculationOnSave = value; }
		}
		#endregion

		#region Custom settings
		public bool AutoPrintArea { get; set; }
		public bool RemoveCalculationChain { get; set; }
		#endregion
	}
}
