using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EasyExcelDotNet.Helpers
{
	public static class ReferenceHelper
	{
		public static string GetLetters(this string input)
		{
			return String.Join("", input.Where(Char.IsLetter));
		}

		public static int GetDigits(this string input)
		{
			return Convert.ToInt32(String.Join("", input.Where(Char.IsDigit)));
		}
	}
}
