using System;
using System.Linq;
using InsERT.Moria.Sfera;
using InsERT.Moria.ModelDanych;
using InsERT.Moria.Kadry.Duze;
using InsERT.Moria.Rozszerzanie;

namespace HrAppka_Import_Pracowników
{
    public static class DepartmentResolver
    {
        public static JednostkaOrganizacyjnaGr? ResolveDepartment(IUchwyt uchwyt, string departmentName)
        {
            if (string.IsNullOrWhiteSpace(departmentName))
                return null;

            var mgr = uchwyt.PodajObiektTypu<IJednostkiOrganizacyjne>();
            if (mgr == null) return null;

            // 1. Check if department exists globally by name (case-insensitive)
            var existing = mgr.Dane.Wszystkie()
                .Where(d => d.Nazwa.ToUpper() == departmentName.ToUpper())
                .AsEnumerable()
                .FirstOrDefault(d => !d.IsInRecycleBin);
            if (existing != null)
            {
                return existing;
            }

            // 2. Not found, find or create parent "FABRYKA"
            var parent = mgr.Dane.Wszystkie()
                .Where(d => d.Nazwa.ToUpper() == "FABRYKA" && d.JednostkaNadrzedna == null)
                .AsEnumerable()
                .FirstOrDefault(d => !d.IsInRecycleBin);
            if (parent == null)
            {
                using (var parentBO = mgr.Utworz())
                {
                    parentBO.Dane.Nazwa = "FABRYKA";
                    if (parentBO.Zapisz())
                    {
                        parent = parentBO.Dane;
                    }
                    else
                    {
                        throw new InvalidOperationException($"Nie udało się utworzyć działu nadrzędnego 'FABRYKA': {string.Join(", ", parentBO.PobierzKomunikatyBledow().Select(x => x.Tresc))}");
                    }
                }
            }

            // 3. Create subdepartment under parent "FABRYKA"
            using (var subBO = mgr.Utworz())
            {
                subBO.Dane.Nazwa = departmentName;
                subBO.Dane.JednostkaNadrzedna = parent;
                if (subBO.Zapisz())
                {
                    return subBO.Dane;
                }
                else
                {
                    throw new InvalidOperationException($"Nie udało się utworzyć poddziału '{departmentName}' pod 'FABRYKA': {string.Join(", ", subBO.PobierzKomunikatyBledow().Select(x => x.Tresc))}");
                }
            }
        }
    }
}
