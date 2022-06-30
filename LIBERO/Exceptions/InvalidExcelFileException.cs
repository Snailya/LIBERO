using System;

namespace LIBERO.Exceptions
{
	internal class InvalidExcelFileException : Exception
	{
		public InvalidExcelFileException(string message) : base(message)
		{
		}
	}
}
