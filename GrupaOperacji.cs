using InsERT.Moria.Rozszerzanie;
using InsERT.Moria.Rozszerzanie.Operacje;
using System;
using System.Collections.Generic;

namespace HrAppka_Import_Pracowników
{
    public class GrupaOperacji : IGrupaOperacji
    {
        public IDostawcaPluginow Dostawca => new DostawcaPluginow();

        public Guid Identyfikator => new Guid("8e2be35b-12d8-4f81-9b16-562a93bbef88");

        public string Nazwa => "Import z HrAppka";

        public string Opis => "Operacje integracji HrAppka z Gratyfikant Nexo Pro";

        public IEnumerable<Operacja> Operacje { get; } = new Operacja[]
        {
            new OperacjaImportuPracownikow()
        };
    }
}
