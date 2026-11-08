using System;
using InsERT.Moria.ModelDanych;
using InsERT.Moria.Rozszerzanie;
using HrAppka_Import_Pracowników.Services;

namespace HrAppka_Import_Pracowników
{
    public static class DepartmentResolver
    {
        public static JednostkaOrganizacyjnaGr? ResolveDepartment(IUchwyt uchwyt, string departmentName)
        {
            if (string.IsNullOrWhiteSpace(departmentName))
                return null;

            return SferaDictionaryCache.Instance.ResolveDepartment(uchwyt, departmentName);
        }
    }
}
