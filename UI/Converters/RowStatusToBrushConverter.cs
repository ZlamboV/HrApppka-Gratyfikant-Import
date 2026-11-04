using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace HrAppka_Import_Pracowników.UI.Converters
{
    /// <summary>
    /// Konwerter mapujący poziom walidacji wiersza na kolor tła zgodny z Wymaganiem R6:
    /// - CriticalError -> #FFDDDD (czerwony)
    /// - Warning       -> #FFF9C4 (żółty)
    /// - Ready         -> #E8F5E9 (zielony)
    /// - Domyślny      -> #FFFFFF (biały)
    /// </summary>
    public class RowStatusToBrushConverter : IValueConverter
    {
        private static readonly SolidColorBrush BrushError = new(Color.FromRgb(0xFF, 0xDD, 0xDD));
        private static readonly SolidColorBrush BrushWarning = new(Color.FromRgb(0xFF, 0xF9, 0xC4));
        private static readonly SolidColorBrush BrushReady = new(Color.FromRgb(0xE8, 0xF5, 0xE9));
        private static readonly SolidColorBrush BrushDefault = new(Colors.White);

        static RowStatusToBrushConverter()
        {
            BrushError.Freeze();
            BrushWarning.Freeze();
            BrushReady.Freeze();
            BrushDefault.Freeze();
        }

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is RowStatusLevel level)
            {
                return level switch
                {
                    RowStatusLevel.CriticalError => BrushError,
                    RowStatusLevel.Warning => BrushWarning,
                    RowStatusLevel.Ready => BrushReady,
                    _ => BrushDefault
                };
            }
            return BrushDefault;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }
    }
}
