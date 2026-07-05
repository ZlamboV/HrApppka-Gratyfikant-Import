using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using InsERT.Moria.Sfera;
using InsERT.Moria.Rozszerzanie;
using InsERT.Mox.Validation;
using InsERT.Mox.ObiektyBiznesowe;
using InsERT.Mox.BusinessObjects;
using InsERT.Mox.DataAccess;

namespace HrAppka_Import_Pracowników
{
    //klasa pomocnicza, przechowująca informacje o błędzie
    internal class BladInfo
    {
        public string Tresc { get; private set; }
        public DataErrorSeverity Waznosc { get; private set; }
        public BladInfo(string tresc, DataErrorSeverity waznosc)
        {
            Tresc = tresc;
            Waznosc = waznosc;
        }
    }

    //klasa z metodą rozszerzającą
    internal static class Rozszerzenia
    {
        //metoda rozszerzająca Uchwyt
        internal static BladInfo[] PodajBledy(this IUchwyt sfera,
            IObiektBiznesowy obiektBiznesowy)
        {
            var lista = new List<BladInfo>();
            var store = sfera.PodajObiektTypu<IValidationMetadataStore>();
            PodajBledy((IBusinessObject)obiektBiznesowy, store, lista);
            var uow = ((IGetUnitOfWork)obiektBiznesowy).UnitOfWork;
            foreach (var innyObiektBiznesowy in uow.Participants.OfType<IBusinessObject>().Where(bo => bo != obiektBiznesowy))
            {
                PodajBledy(innyObiektBiznesowy, store, lista);
            }
            return lista.ToArray();
        }

        //metoda pomocnicza sprawdzająca występowanie błędów
        private static bool HasAnyErrors(IValidationMetadataStore store, ITypedDataErrorInfo errorInfo)
        {
            return
                errorInfo != null &&
                (
                    errorInfo.Errors.Any() || 
                    errorInfo.MemberErrors.Any()
                );
        }

        //metoda pomocnicza zbierająca informacje o błędach
        private static void PodajBledy(this IBusinessObject obiektBiznesowy,
            IValidationMetadataStore store,
            List<BladInfo> bledy)
        {
            HashSet<ITypedDataErrorInfo> invalidData = new HashSet<ITypedDataErrorInfo>();
            ((IGetDataDomain)obiektBiznesowy).DataDomain.TraverseData(
                obiektBiznesowy.Data, (o) =>
                {
                    ITypedDataErrorInfo? dataErrorInfoEx = o as ITypedDataErrorInfo;
                    if (dataErrorInfoEx != null && HasAnyErrors(store, dataErrorInfoEx))
                    {
                        invalidData.Add(dataErrorInfoEx);
                    }
                    return true;
                });

            foreach (var encjaZBledami in invalidData)
            {
                foreach (var bladNaCalejEncji in encjaZBledami.Errors)
                {
                    StringBuilder sb = new StringBuilder();
                    sb.AppendLine(bladNaCalejEncji.ToString());
                    sb.Append(" na encjach:" + encjaZBledami.GetType().Name);
                    DataErrorType errorType = store.GetEntryForClrType(bladNaCalejEncji.GetType());
                    bledy.Add(new BladInfo(sb.ToString(), errorType.Severity));
                }
                foreach (var bladNaKonkretnychPolach in encjaZBledami.MemberErrors)
                {
                    StringBuilder sb = new StringBuilder();
                    sb.AppendLine(bladNaKonkretnychPolach.Key.ToString());
                    sb.AppendLine(" na polach:");
                    sb.AppendLine(string.Join(", ", bladNaKonkretnychPolach.Select(b => encjaZBledami.GetType().Name + "." + b)));
                    DataErrorType errorType = store.GetEntryForClrType(bladNaKonkretnychPolach.Key.GetType());
                    bledy.Add(new BladInfo(sb.ToString(), errorType.Severity));
                }
            }
        }
    }
}
