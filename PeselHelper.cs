using System;

namespace HrAppka_Import_Pracowników
{
    public static class PeselHelper
    {
        public static bool ValidatePesel(string pesel, out DateTime? birthDate, out string gender)
        {
            birthDate = null;
            gender = "";

            if (string.IsNullOrWhiteSpace(pesel) || pesel.Length != 11 || !pesel.All(char.IsDigit))
            {
                return false;
            }

            int[] weights = { 1, 3, 7, 9, 1, 3, 7, 9, 1, 3 };
            int sum = 0;
            for (int i = 0; i < 10; i++)
            {
                sum += (pesel[i] - '0') * weights[i];
            }
            int checksum = (10 - (sum % 10)) % 10;
            if (checksum != pesel[10] - '0')
            {
                return false;
            }

            int year = int.Parse(pesel.Substring(0, 2));
            int month = int.Parse(pesel.Substring(2, 2));
            int day = int.Parse(pesel.Substring(4, 2));

            if (month >= 81 && month <= 92)
            {
                year += 1800;
                month -= 80;
            }
            else if (month >= 1 && month <= 12)
            {
                year += 1900;
            }
            else if (month >= 21 && month <= 32)
            {
                year += 2000;
                month -= 20;
            }
            else if (month >= 41 && month <= 52)
            {
                year += 2100;
                month -= 40;
            }
            else if (month >= 61 && month <= 72)
            {
                year += 2200;
                month -= 60;
            }

            try
            {
                birthDate = new DateTime(year, month, day);
            }
            catch
            {
                return false;
            }

            gender = (pesel[9] - '0') % 2 == 0 ? "Kobieta" : "Mężczyzna";
            return true;
        }

        public static string NormalizeWhitespace(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return string.Empty;
            return System.Text.RegularExpressions.Regex.Replace(input.Trim(), @"\s+", " ");
        }

        public static int CalculateAge(DateTime birthDate, DateTime atDate)
        {
            int age = atDate.Year - birthDate.Year;
            if (atDate < birthDate.AddYears(age))
                age--;
            return age;
        }

        /// <summary>
        /// Normalizes a PESEL string that may have been corrupted by Excel's
        /// scientific notation (e.g. "8.4030421022E10" instead of "84030421022").
        /// Also trims whitespace and removes any non-digit characters after conversion.
        /// </summary>
        public static string NormalizePesel(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
                return raw;

            string trimmed = raw.Trim();

            // Detect scientific notation (contains 'E' or 'e' with digits)
            if (trimmed.IndexOf('E') >= 0 || trimmed.IndexOf('e') >= 0)
            {
                if (double.TryParse(trimmed, System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out double numericValue))
                {
                    // Convert back to integer string (PESEL is always 11 digits, no decimals)
                    long peselLong = (long)numericValue;
                    string converted = peselLong.ToString();

                    // Pad with leading zero if needed (PESEL starting with 0)
                    if (converted.Length == 10)
                        converted = "0" + converted;

                    return converted;
                }
            }

            return trimmed;
        }
    }
}
