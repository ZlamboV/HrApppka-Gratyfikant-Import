using System;
using System.Globalization;
using System.Text.RegularExpressions;
using InsERT.Moria.Klienci;
using InsERT.Moria.ModelDanych;

namespace HrAppka_Import_Pracowników.Services
{
    /// <summary>
    /// Wynik identyfikacji dokumentu tożsamości pracownika gotowy do przypisania do encji Sfery.
    /// </summary>
    public sealed class DocumentIdentificationResult
    {
        public TypDokumentuTozsamosci DocumentType { get; set; } = TypDokumentuTozsamosci.Brak;
        public string DocumentNumber { get; set; } = string.Empty;
        public DateTime? IssueDate { get; set; }
        public DateTime? ExpirationDate { get; set; }
        public Panstwo? CountryOfIssue { get; set; }
        public string SourceColumn { get; set; } = string.Empty;
        public string? WarningMessage { get; set; }

        public bool HasDocument => DocumentType != TypDokumentuTozsamosci.Brak && !string.IsNullOrWhiteSpace(DocumentNumber);
    }

    /// <summary>
    /// Serwis inteligentnego doboru dokumentu tożsamości na podstawie obywatelstwa i priorytetu kolumn w pliku źródłowym:
    /// - Obywatele RP -> Dowód osobisty (kol. 21), TypDokumentuTozsamosci = Dowod (1).
    /// - Rezydenci UE -> Dokument tożsamości / Paszport biometryczny (kol. 27 lub 21), TypDokumentuTozsamosci = Inne (3) lub Paszport (2).
    /// - Cudzoziemcy spoza UE -> Paszport (kol. 24 lub 27), TypDokumentuTozsamosci = Paszport (2) z OBOWIĄZKOWYM krajem wydania.
    /// </summary>
    public static class DocumentIdentificationService
    {
        private static readonly Regex NonAlphanumericRegex = new Regex(@"[^A-Za-z0-9]", RegexOptions.Compiled);

        public static DocumentIdentificationResult DetermineDocument(PracownikImportRow row, Panstwo? panstwo)
        {
            var result = new DocumentIdentificationResult();
            if (row == null) return result;

            bool isPolish = panstwo != null && panstwo.KodPanstwaUE.Equals("PL", StringComparison.OrdinalIgnoreCase);
            bool isEu = panstwo != null && panstwo.CzlonekUE && !isPolish;
            bool isForeignNonEu = panstwo != null && !panstwo.CzlonekUE && !isPolish;

            // Scenariusz 1: Obywatele Polski -> Priorytet: DOWÓD osobisty (Kol 21)
            if (isPolish)
            {
                if (!string.IsNullOrWhiteSpace(row.Dowod))
                {
                    result.DocumentType = TypDokumentuTozsamosci.Dowod;
                    result.DocumentNumber = SanitizeDocumentNumber(row.Dowod, ZusSchemaLimits.ZusSeriaNrDokMaxLen);
                    result.IssueDate = ParseDate(row.DowodWydany);
                    result.ExpirationDate = ParseDate(row.DowodWazny);
                    result.CountryOfIssue = panstwo;
                    result.SourceColumn = "Dowód osobisty (kol. 21)";
                    return result;
                }

                // Fallback 1: Paszport (Kol 24)
                if (!string.IsNullOrWhiteSpace(row.Paszport))
                {
                    result.DocumentType = TypDokumentuTozsamosci.Paszport;
                    result.DocumentNumber = SanitizeDocumentNumber(row.Paszport, ZusSchemaLimits.ZusSeriaNrDokMaxLen);
                    result.IssueDate = ParseDate(row.PaszportWydany);
                    result.ExpirationDate = ParseDate(row.PaszportWazny);
                    result.CountryOfIssue = panstwo;
                    result.SourceColumn = "Paszport (kol. 24) [fallback]";
                    return result;
                }

                // Fallback 2: Paszport biometryczny (Kol 27)
                if (!string.IsNullOrWhiteSpace(row.PaszportBio))
                {
                    result.DocumentType = TypDokumentuTozsamosci.Paszport;
                    result.DocumentNumber = SanitizeDocumentNumber(row.PaszportBio, ZusSchemaLimits.ZusSeriaNrDokMaxLen);
                    result.IssueDate = ParseDate(row.PaszportBioWydany);
                    result.ExpirationDate = ParseDate(row.PaszportBioWazny);
                    result.CountryOfIssue = panstwo;
                    result.SourceColumn = "Paszport biometryczny (kol. 27) [fallback]";
                    return result;
                }

                // Fallback 3: Brak dokumentu (obywatel polski identyfikowany numerem PESEL)
                result.DocumentType = TypDokumentuTozsamosci.Brak;
                result.CountryOfIssue = panstwo;
                if (string.IsNullOrWhiteSpace(row.Pesel))
                {
                    result.WarningMessage = "Brak dowodu osobistego i brak numeru PESEL dla obywatela Polski.";
                }
                return result;
            }

            // Scenariusz 2: Rezydenci UE (poza PL) -> Priorytet: DOKUMENT tożsamości / Paszport biometryczny (Kol 27 / Kol 21)
            if (isEu)
            {
                if (!string.IsNullOrWhiteSpace(row.PaszportBio))
                {
                    result.DocumentType = TypDokumentuTozsamosci.Inne;
                    result.DocumentNumber = SanitizeDocumentNumber(row.PaszportBio, ZusSchemaLimits.ZusSeriaNrDokMaxLen);
                    result.IssueDate = ParseDate(row.PaszportBioWydany);
                    result.ExpirationDate = ParseDate(row.PaszportBioWazny);
                    result.CountryOfIssue = panstwo; // OBOWIĄZKOWE!
                    result.SourceColumn = "Paszport biometryczny (kol. 27)";
                    return result;
                }

                if (!string.IsNullOrWhiteSpace(row.Dowod))
                {
                    result.DocumentType = TypDokumentuTozsamosci.Inne;
                    result.DocumentNumber = SanitizeDocumentNumber(row.Dowod, ZusSchemaLimits.ZusSeriaNrDokMaxLen);
                    result.IssueDate = ParseDate(row.DowodWydany);
                    result.ExpirationDate = ParseDate(row.DowodWazny);
                    result.CountryOfIssue = panstwo; // OBOWIĄZKOWE!
                    result.SourceColumn = "Dokument tożsamości UE (kol. 21)";
                    return result;
                }

                if (!string.IsNullOrWhiteSpace(row.Paszport))
                {
                    result.DocumentType = TypDokumentuTozsamosci.Paszport;
                    result.DocumentNumber = SanitizeDocumentNumber(row.Paszport, ZusSchemaLimits.ZusSeriaNrDokMaxLen);
                    result.IssueDate = ParseDate(row.PaszportWydany);
                    result.ExpirationDate = ParseDate(row.PaszportWazny);
                    result.CountryOfIssue = panstwo; // OBOWIĄZKOWE!
                    result.SourceColumn = "Paszport (kol. 24) [fallback]";
                    return result;
                }

                result.DocumentType = TypDokumentuTozsamosci.Brak;
                result.CountryOfIssue = panstwo;
                result.WarningMessage = $"Brak dokumentu tożsamości dla rezydenta UE ({panstwo?.Nazwa}).";
                return result;
            }

            // Scenariusz 3: Cudzoziemcy spoza UE -> Priorytet: PASZPORT (Kol 24 lub Kol 27)
            if (isForeignNonEu)
            {
                if (!string.IsNullOrWhiteSpace(row.Paszport))
                {
                    result.DocumentType = TypDokumentuTozsamosci.Paszport;
                    result.DocumentNumber = SanitizeDocumentNumber(row.Paszport, ZusSchemaLimits.ZusSeriaNrDokMaxLen);
                    result.IssueDate = ParseDate(row.PaszportWydany);
                    result.ExpirationDate = ParseDate(row.PaszportWazny);
                    result.CountryOfIssue = panstwo; // BEZWZGLĘDNIE OBOWIĄZKOWE!
                    result.SourceColumn = "Paszport (kol. 24)";
                    return result;
                }

                if (!string.IsNullOrWhiteSpace(row.PaszportBio))
                {
                    result.DocumentType = TypDokumentuTozsamosci.Paszport;
                    result.DocumentNumber = SanitizeDocumentNumber(row.PaszportBio, ZusSchemaLimits.ZusSeriaNrDokMaxLen);
                    result.IssueDate = ParseDate(row.PaszportBioWydany);
                    result.ExpirationDate = ParseDate(row.PaszportBioWazny);
                    result.CountryOfIssue = panstwo; // BEZWZGLĘDNIE OBOWIĄZKOWE!
                    result.SourceColumn = "Paszport biometryczny (kol. 27)";
                    return result;
                }

                if (!string.IsNullOrWhiteSpace(row.Dowod))
                {
                    result.DocumentType = TypDokumentuTozsamosci.Inne;
                    result.DocumentNumber = SanitizeDocumentNumber(row.Dowod, ZusSchemaLimits.ZusSeriaNrDokMaxLen);
                    result.IssueDate = ParseDate(row.DowodWydany);
                    result.ExpirationDate = ParseDate(row.DowodWazny);
                    result.CountryOfIssue = panstwo; // BEZWZGLĘDNIE OBOWIĄZKOWE!
                    result.SourceColumn = "Dokument tożsamości / Karta pobytu (kol. 21) [fallback]";
                    return result;
                }

                result.DocumentType = TypDokumentuTozsamosci.Brak;
                result.CountryOfIssue = panstwo;
                result.WarningMessage = $"Wymagany paszport dla cudzoziemca spoza UE ({panstwo?.Nazwa}). Brak danych dokumentu w pliku.";
                return result;
            }

            // Scenariusz 4: Nieznany / nierozpoznany kraj -> Heurystyczne dopasowanie
            if (!string.IsNullOrWhiteSpace(row.Paszport))
            {
                result.DocumentType = TypDokumentuTozsamosci.Paszport;
                result.DocumentNumber = SanitizeDocumentNumber(row.Paszport, ZusSchemaLimits.ZusSeriaNrDokMaxLen);
                result.IssueDate = ParseDate(row.PaszportWydany);
                result.ExpirationDate = ParseDate(row.PaszportWazny);
                result.SourceColumn = "Paszport (kol. 24)";
                result.WarningMessage = "Nieznane obywatelstwo pracownika; nie można ustawić kraju wydania dokumentu.";
                return result;
            }

            if (!string.IsNullOrWhiteSpace(row.PaszportBio))
            {
                result.DocumentType = TypDokumentuTozsamosci.Paszport;
                result.DocumentNumber = SanitizeDocumentNumber(row.PaszportBio, ZusSchemaLimits.ZusSeriaNrDokMaxLen);
                result.IssueDate = ParseDate(row.PaszportBioWydany);
                result.ExpirationDate = ParseDate(row.PaszportBioWazny);
                result.SourceColumn = "Paszport biometryczny (kol. 27)";
                result.WarningMessage = "Nieznane obywatelstwo pracownika; nie można ustawić kraju wydania dokumentu.";
                return result;
            }

            if (!string.IsNullOrWhiteSpace(row.Dowod))
            {
                result.DocumentType = TypDokumentuTozsamosci.Dowod;
                result.DocumentNumber = SanitizeDocumentNumber(row.Dowod, ZusSchemaLimits.ZusSeriaNrDokMaxLen);
                result.IssueDate = ParseDate(row.DowodWydany);
                result.ExpirationDate = ParseDate(row.DowodWazny);
                result.SourceColumn = "Dowód (kol. 21)";
                result.WarningMessage = "Nieznane obywatelstwo pracownika; nie można ustawić kraju wydania dokumentu.";
                return result;
            }

            result.DocumentType = TypDokumentuTozsamosci.Brak;
            return result;
        }

        public static void ApplyToEmployee(Podmiot employeePodmiot, DocumentIdentificationResult docResult)
        {
            if (employeePodmiot?.Osoba == null) return;
            var osoba = employeePodmiot.Osoba;

            if (docResult.HasDocument)
            {
                osoba.TypDokumentuTozsamosci = (byte)docResult.DocumentType;
                osoba.SeriaINumerDokumentuTozsamosci = docResult.DocumentNumber;
                osoba.DataWystawieniaDokumentuTozsamosci = docResult.IssueDate;

                if (osoba.Pracownik != null)
                {
                    if (osoba.Pracownik.PracownikGr != null)
                    {
                        osoba.Pracownik.PracownikGr.WaznoscDokumentuTozsamosci = docResult.ExpirationDate;
                    }
                    osoba.Pracownik.KrajWydaniaDokumentuTozsamosci = docResult.CountryOfIssue;
                }
            }
            else
            {
                osoba.TypDokumentuTozsamosci = (byte)TypDokumentuTozsamosci.Brak;
                osoba.SeriaINumerDokumentuTozsamosci = null;
                osoba.DataWystawieniaDokumentuTozsamosci = null;

                if (osoba.Pracownik != null)
                {
                    if (osoba.Pracownik.PracownikGr != null)
                    {
                        osoba.Pracownik.PracownikGr.WaznoscDokumentuTozsamosci = null;
                    }
                    osoba.Pracownik.KrajWydaniaDokumentuTozsamosci = null;
                }
            }
        }

        /// <summary>
        /// Oczyszcza numer dokumentu tożsamości do czystego ciągu alfanumerycznego (A-Z, 0-9),
        /// usuwając wszelkie znaki specjalne (np. '№', '#', kropki, myślniki, ukośniki, spacje)
        /// w ścisłej zgodności ze schematem XSD ZUS Płatnik (typ t_SeriaNrDok o limicie do 9 znaków).
        /// </summary>
        public static string SanitizeDocumentNumber(string? rawNumber, int maxLength = 9)
        {
            if (string.IsNullOrWhiteSpace(rawNumber)) return string.Empty;

            string cleaned = NonAlphanumericRegex.Replace(rawNumber, "").ToUpperInvariant();

            if (cleaned.Length > maxLength)
            {
                return cleaned.Substring(0, maxLength);
            }

            return cleaned;
        }

        public static DateTime? ParseDate(string? rawDate)
        {
            if (string.IsNullOrWhiteSpace(rawDate)) return null;

            string s = rawDate.Trim();

            if (DateTime.TryParseExact(s, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
                return dt;

            if (DateTime.TryParseExact(s, "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
                return dt;

            if (DateTime.TryParse(s, CultureInfo.GetCultureInfo("pl-PL"), DateTimeStyles.None, out dt))
                return dt;

            if (DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
                return dt;

            return null;
        }
    }
}
