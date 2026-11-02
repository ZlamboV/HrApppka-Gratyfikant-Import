using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace HrAppka_Import_Pracowników.Services
{
    public interface IRfc4180CsvReader
    {
        List<List<string>> ReadRows(string filePath, char? delimiter = null, Encoding? encoding = null);
        IEnumerable<IReadOnlyList<string>> StreamRows(string filePath, char? delimiter = null, Encoding? encoding = null);
        IEnumerable<IReadOnlyList<string>> StreamRows(TextReader reader, char delimiter);
        char DetectDelimiter(string filePath, Encoding? encoding = null);
        Encoding DetectEncoding(string filePath);
    }

    /// <summary>
    /// Strumieniowy parser CSV w pełni zgodny ze standardem RFC 4180.
    /// Obsługuje wartości wieloliniowe w cudzysłowach, znaki ucieczki cudzysłowu (""),
    /// automatyczne wykrywanie separatora (średnik, przecinek, tabulacja) oraz kodowania (UTF-8 z BOM, UTF-8, Windows-1250).
    /// </summary>
    public class Rfc4180CsvReader : IRfc4180CsvReader
    {
        static Rfc4180CsvReader()
        {
            // Rejestracja dostawcy stron kodowych (m.in. dla Windows-1250 w .NET 8)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        public static List<List<string>> ReadAllRows(string filePath, char? delimiter = null, Encoding? encoding = null)
        {
            return Stream(filePath, delimiter, encoding).Select(r => r.ToList()).ToList();
        }

        public static List<List<string>> ReadRows(string filePath, char? delimiter = null)
        {
            return ReadAllRows(filePath, delimiter, null);
        }

        public static IEnumerable<IReadOnlyList<string>> Stream(string filePath, char? delimiter = null, Encoding? encoding = null)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("Ścieżka do pliku nie może być pusta.", nameof(filePath));
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"Plik CSV nie istnieje: {filePath}", filePath);

            var resolvedEncoding = encoding ?? AutoDetectEncoding(filePath);
            char resolvedDelimiter = delimiter ?? AutoDetectDelimiter(filePath, resolvedEncoding);

            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = new StreamReader(stream, resolvedEncoding, detectEncodingFromByteOrderMarks: true);

            foreach (var row in Stream(reader, resolvedDelimiter))
            {
                yield return row;
            }
        }

        public static IEnumerable<IReadOnlyList<string>> Stream(TextReader reader, char delimiter)
        {
            if (reader == null) throw new ArgumentNullException(nameof(reader));

            var currentRow = new List<string>();
            var currentField = new StringBuilder();
            bool inQuotes = false;
            bool hasDataInRow = false;
            int cInt;

            while ((cInt = reader.Read()) != -1)
            {
                char c = (char)cInt;

                if (inQuotes)
                {
                    if (c == '"')
                    {
                        int next = reader.Peek();
                        if (next == '"')
                        {
                            // Cudzysłów ucieczkowy: "" -> "
                            reader.Read();
                            currentField.Append('"');
                        }
                        else
                        {
                            // Zamykający cudzysłów
                            inQuotes = false;
                        }
                    }
                    else
                    {
                        // Wewnątrz cudzysłowu znaki nowej linii, separatory itp. są częścią pola
                        currentField.Append(c);
                    }
                }
                else
                {
                    if (c == '"')
                    {
                        inQuotes = true;
                        hasDataInRow = true;
                    }
                    else if (c == delimiter)
                    {
                        currentRow.Add(currentField.ToString());
                        currentField.Clear();
                        hasDataInRow = true;
                    }
                    else if (c == '\r')
                    {
                        if (reader.Peek() == '\n')
                        {
                            reader.Read(); // Konsumpcja '\n' dla CRLF
                        }
                        currentRow.Add(currentField.ToString());
                        currentField.Clear();
                        yield return currentRow;
                        currentRow = new List<string>();
                        hasDataInRow = false;
                    }
                    else if (c == '\n')
                    {
                        currentRow.Add(currentField.ToString());
                        currentField.Clear();
                        yield return currentRow;
                        currentRow = new List<string>();
                        hasDataInRow = false;
                    }
                    else
                    {
                        currentField.Append(c);
                        hasDataInRow = true;
                    }
                }
            }

            // Opróżnienie pozostałości wiersza, jeśli istnieją dane
            if (hasDataInRow || currentField.Length > 0 || currentRow.Count > 0)
            {
                currentRow.Add(currentField.ToString());
                // Pominięcie pustego pojedynczego wiersza na samym końcu pliku
                if (!(currentRow.Count == 1 && string.IsNullOrEmpty(currentRow[0])))
                {
                    yield return currentRow;
                }
            }
        }

        public static char AutoDetectDelimiter(string filePath, Encoding? encoding = null)
        {
            if (!File.Exists(filePath)) return ';';

            var enc = encoding ?? AutoDetectEncoding(filePath);
            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = new StreamReader(stream, enc, detectEncodingFromByteOrderMarks: true);

            var candidates = new Dictionary<char, int>
            {
                { ';', 0 },
                { ',', 0 },
                { '\t', 0 },
                { '|', 0 }
            };

            bool inQuotes = false;
            int cInt;
            while ((cInt = reader.Read()) != -1)
            {
                char c = (char)cInt;
                if (c == '"')
                {
                    int next = reader.Peek();
                    if (inQuotes && next == '"')
                    {
                        reader.Read();
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }
                }
                else if (!inQuotes)
                {
                    if (c == '\r' || c == '\n')
                    {
                        if (candidates.Values.Any(v => v > 0))
                            break;
                    }
                    else if (candidates.ContainsKey(c))
                    {
                        candidates[c]++;
                    }
                }
            }

            var best = candidates.OrderByDescending(kvp => kvp.Value).First();
            return best.Value > 0 ? best.Key : ';';
        }

        public static Encoding AutoDetectEncoding(string filePath)
        {
            if (!File.Exists(filePath)) return Encoding.UTF8;

            byte[] buffer = new byte[4];
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                int read = fs.Read(buffer, 0, buffer.Length);

                // 1. 4-bajtowe znaczniki kolejności bajtów (BOM) sprawdzane przed 2-bajtowym UTF-16 LE
                if (read >= 4 && buffer[0] == 0xFF && buffer[1] == 0xFE && buffer[2] == 0x00 && buffer[3] == 0x00)
                    return Encoding.UTF32; // UTF-32 LE (CodePage 12000)
                if (read >= 4 && buffer[0] == 0x00 && buffer[1] == 0x00 && buffer[2] == 0xFE && buffer[3] == 0xFF)
                    return new UTF32Encoding(true, true); // UTF-32 BE (CodePage 12001)

                // 2. 3-bajtowy UTF-8 z BOM
                if (read >= 3 && buffer[0] == 0xEF && buffer[1] == 0xBB && buffer[2] == 0xBF)
                    return new UTF8Encoding(true, false);

                // 3. 2-bajtowe UTF-16
                if (read >= 2 && buffer[0] == 0xFF && buffer[1] == 0xFE)
                    return Encoding.Unicode; // UTF-16 LE (CodePage 1200)
                if (read >= 2 && buffer[0] == 0xFE && buffer[1] == 0xFF)
                    return Encoding.BigEndianUnicode; // UTF-16 BE (CodePage 1201)
            }

            // Próba weryfikacji UTF-8 bez BOM z obsługą ucięcia sekwencji wielobajtowej na granicy bufora 4096 B
            try
            {
                byte[] sample = new byte[4096];
                int sampleLen;
                bool isEndOfFile;

                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    sampleLen = fs.Read(sample, 0, sample.Length);
                    isEndOfFile = fs.Position >= fs.Length;
                }

                if (sampleLen == 0) return Encoding.UTF8;

                int testLen = AdjustLengthForIncompleteTrailingUtf8(sample, sampleLen, isEndOfFile);

                var utf8Strict = new UTF8Encoding(false, true);
                utf8Strict.GetCharCount(sample, 0, testLen);
                return Encoding.UTF8;
            }
            catch (DecoderFallbackException)
            {
                try
                {
                    return Encoding.GetEncoding("windows-1250");
                }
                catch
                {
                    return Encoding.Default;
                }
            }
        }

        /// <summary>
        /// Odcina niekompletne bajty sekwencji UTF-8 na samym końcu bufora próbki (np. offset 4095/4096),
        /// zapobiegając błędnemu rzucaniu DecoderFallbackException i fałszywemu wykrywaniu Windows-1250.
        /// </summary>
        private static int AdjustLengthForIncompleteTrailingUtf8(byte[] buffer, int length, bool isEndOfFile)
        {
            if (isEndOfFile || length <= 0) return length;

            int maxLookback = Math.Min(4, length);
            for (int i = 1; i <= maxLookback; i++)
            {
                byte b = buffer[length - i];

                if ((b & 0x80) == 0)
                {
                    break; // Znak ASCII (1-bajtowy); brak uciętej sekwencji
                }

                if ((b & 0xC0) == 0xC0)
                {
                    // Bajt początkowy sekwencji wielobajtowej
                    int expectedSequenceLength;
                    if ((b & 0xE0) == 0xC0) expectedSequenceLength = 2;       // 110xxxxx
                    else if ((b & 0xF0) == 0xE0) expectedSequenceLength = 3;  // 1110xxxx
                    else if ((b & 0xF8) == 0xF0) expectedSequenceLength = 4;  // 11110xxx
                    else break; // Nieprawidłowy nagłówek; weryfikator UTF-8 rzuci wyjątek

                    int presentBytes = i;
                    if (presentBytes < expectedSequenceLength)
                    {
                        // Sekwencja została rozcięta na granicy bufora; odcinamy niepełną końcówkę
                        return length - presentBytes;
                    }
                    break;
                }
            }

            return length;
        }

        // IRfc4180CsvReader implementation
        List<List<string>> IRfc4180CsvReader.ReadRows(string filePath, char? delimiter, Encoding? encoding)
            => ReadAllRows(filePath, delimiter, encoding);

        IEnumerable<IReadOnlyList<string>> IRfc4180CsvReader.StreamRows(string filePath, char? delimiter, Encoding? encoding)
            => Stream(filePath, delimiter, encoding);

        IEnumerable<IReadOnlyList<string>> IRfc4180CsvReader.StreamRows(TextReader reader, char delimiter)
            => Stream(reader, delimiter);

        char IRfc4180CsvReader.DetectDelimiter(string filePath, Encoding? encoding)
            => AutoDetectDelimiter(filePath, encoding);

        Encoding IRfc4180CsvReader.DetectEncoding(string filePath)
            => AutoDetectEncoding(filePath);
    }
}
