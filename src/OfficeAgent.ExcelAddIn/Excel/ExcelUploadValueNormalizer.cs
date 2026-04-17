using System;
using System.Globalization;

namespace OfficeAgent.ExcelAddIn.Excel
{
    internal sealed class ExcelUploadValueNormalizer
    {
        public bool TryNormalize(object value, string numberFormat, out string normalized)
        {
            if (value == null)
            {
                normalized = string.Empty;
                return true;
            }

            if (value is string text)
            {
                normalized = text;
                return true;
            }

            if (value is bool booleanValue)
            {
                normalized = Convert.ToString(booleanValue, CultureInfo.InvariantCulture) ?? string.Empty;
                return true;
            }

            if (value is double numericValue)
            {
                if (RequiresDisplayTextFallback(numberFormat))
                {
                    normalized = null;
                    return false;
                }

                normalized = numericValue % 1d == 0d
                    ? numericValue.ToString("0", CultureInfo.InvariantCulture)
                    : numericValue.ToString("0.###############", CultureInfo.InvariantCulture);
                return true;
            }

            normalized = Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
            return true;
        }

        private static bool RequiresDisplayTextFallback(string numberFormat)
        {
            var format = numberFormat ?? string.Empty;
            if (string.IsNullOrWhiteSpace(format) || string.Equals(format, "General", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return format.Contains("%") ||
                   format.Contains("y") ||
                   format.Contains("m") ||
                   format.Contains("d") ||
                   format.Contains("h") ||
                   format.Contains("s") ||
                   format.Contains("0");
        }
    }
}
