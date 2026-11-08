using System;

namespace HrAppka_Import_Pracowników
{
    /// <summary>
    /// Raport postępu asynchronicznego importu pracowników.
    /// Zgodny z testem R6_InteractiveWpfUiTests §R6_05.
    /// </summary>
    public record ImportProgressReport(
        int Current, 
        int Total, 
        string CurrentEmployee, 
        string StatusMessage);
}
