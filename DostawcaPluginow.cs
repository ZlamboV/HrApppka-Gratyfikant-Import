using InsERT.Moria.Rozszerzanie;
using System.Collections.Generic;

namespace HrAppka_Import_Pracowników
{
    public class DostawcaPluginow : IDostawcaPluginow
    {
        public string Adres => "Kijów, Ukraina";

        public string AdresWWW => "www.example.com";

        public IEnumerable<string> Kontakty => new string[] { "support@example.com" };

        public string KRS => "";

        public string Nazwa => "Antigravity Dev Team";

        public string NIP => "";

        public string REGON => "";
    }
}
