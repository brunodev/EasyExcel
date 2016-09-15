using DocumentFormat.OpenXml.Packaging;
using EasyExcelDotNet.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EasyExcelDotNet.Modules
{
	public class EasyExcelCalculationChain : BaseModule
	{
		public CalculationChainPart Part { get; private set; }

		public EasyExcelCalculationChain(EasyExcelDocument document, CalculationChainPart calculationChainPart) : base(document)
		{
			Part = calculationChainPart;
		}

		public void Remove()
		{
			Document.WorkbookPart.DeletePart(Part);	
		}
	}
}
