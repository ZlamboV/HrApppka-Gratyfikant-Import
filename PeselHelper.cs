using System;

namespace HrAppka_Import_Pracowników
{
    public static class PeselHelper
    {
        public static bool ValidatePesel(string pesel, out DateTime? birthDate, out string gender)
        {
            birthDate = null;
            gender = "";

            if (string.IsNullOrWhiteSpace(pesel) || pesel.Length != 11)
                return false;

            foreach (char c in pesel)
            {
                if (!char.IsDigit(c)) return false;
            }

            // Checksum
            int[] weights = { 1, 3, 7, 9, 1, 3, 7, 9, 1, 3 };
            int sum = 0;
            for (int i = 0; i < 10; i++)
            {
                sum += (pesel[i] - '0') * weights[i];
            }
            int control = (10 - (sum % 10)) % 10;
            if (control != (pesel[10] - '0'))
                return false;

            // Gender
            int genderDigit = pesel[9] - '0';
            gender = (genderDigit % 2 == 1) ? "Mężczyzna" : "Kobieta";

            // Date of birth
            int yy = int.Parse(pesel.Substring(0, 2));
            int mm = int.Parse(pesel.Substring(2, 2));
            int dd = int.Parse(pesel.Substring(4, 2));

            int year = 0;
            int month = 0;

            if (mm >= 1 && mm <= 12)
            {
                year = 1900 + yy;
                month = mm;
            }
            else if (mm >= 21 && mm <= 32)
            {
                year = 2000 + yy;
                month = mm - 20;
            }
            else if (mm >= 41 && mm <= 52)
            {
                year = 2100 + yy;
                month = mm - 40;
            }
            else if (mm >= 61 && mm <= 72)
            {
                year = 2200 + yy;
                month = mm - 60;
            }
            else if (mm >= 81 && mm <= 92)
            {
                year = 1800 + yy;
                month = mm - 80;
            }
            else
            {
                return false;
            }

            try
            {
                birthDate = new DateTime(year, month, dd);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static int CalculateAge(DateTime birthDate, DateTime atDate)
        {
            int age = atDate.Year - birthDate.Year;
            if (atDate < birthDate.AddYears(age))
                age--;
            return age;
        }
    }
}
