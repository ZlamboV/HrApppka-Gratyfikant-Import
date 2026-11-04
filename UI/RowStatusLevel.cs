using System;

namespace HrAppka_Import_Pracowników
{
    /// <summary>
    /// Poziomy statusu walidacji wiersza w DataGrid dla Wymagania R6.
    /// Zgodny z testami R6_InteractiveWpfUiTests.
    /// </summary>
    public enum RowStatusLevel
    {
        Ready,
        Warning,
        CriticalError
    }
}
