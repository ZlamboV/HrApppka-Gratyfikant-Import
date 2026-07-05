using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using InsERT.Moria.Raporty;

namespace HrAppka_Import_Pracowników
{
    public class PracownikImportRow : INotifyPropertyChanged, IDataErrorInfo
    {
        private string imie = "";
        private string drugieImie = "";
        private string nazwisko = "";
        private string plecExcel = "";
        private string birthDateExcel = "";
        private string pesel = "";
        private string obywatelstwo = "";
        private string paszport = "";
        private string paszportWydany = "";
        private string paszportWazny = "";
        private string ulica = "";
        private string nrDomu = "";
        private string miejscowosc = "";
        private string kodPocztowy = "";
        private string gmina = "";
        private string powiat = "";
        private string wojewodztwo = "";
        private string kraj = "";
        private string telefon = "";
        private string email = "";
        private string bankStr = "";
        private string nfzStr = "";
        private string typUmowy = "";
        private string dataRozpoczecia = "";
        private string dataZakonczenia = "";
        private string dataZawarcia = "";
        private string dzialNazwa = "";
        private string kodZawodu = "";
        private string jestStudentem = "";
        private string status = "Gotowy";
        private string ostrzezenia = "";
        private string urzadSkarbowy = "";

        // Hiding auxiliary fields
        [Browsable(false)]
        [BezKolumny]
        public string UrzadSkarbowy
        {
            get => urzadSkarbowy;
            set => SetProperty(ref urzadSkarbowy, value);
        }
        [Browsable(false)]
        [BezKolumny]
        public string Imie
        {
            get => imie;
            set
            {
                if (SetProperty(ref imie, value))
                {
                    OnPropertyChanged(nameof(PracownikNazwa));
                }
            }
        }

        [Browsable(false)]
        [BezKolumny]
        public string DrugieImie
        {
            get => drugieImie;
            set => SetProperty(ref drugieImie, value);
        }

        [Browsable(false)]
        [BezKolumny]
        public string Nazwisko
        {
            get => nazwisko;
            set
            {
                if (SetProperty(ref nazwisko, value))
                {
                    OnPropertyChanged(nameof(PracownikNazwa));
                }
            }
        }

        [Browsable(false)]
        [BezKolumny]
        public string PlecExcel
        {
            get => plecExcel;
            set => SetProperty(ref plecExcel, value);
        }

        [Browsable(false)]
        [BezKolumny]
        public string BirthDateExcel
        {
            get => birthDateExcel;
            set => SetProperty(ref birthDateExcel, value);
        }

        [Browsable(false)]
        [BezKolumny]
        public string Obywatelstwo
        {
            get => obywatelstwo;
            set => SetProperty(ref obywatelstwo, value);
        }

        [Browsable(false)]
        [BezKolumny]
        public string Paszport
        {
            get => paszport;
            set => SetProperty(ref paszport, value);
        }

        [Browsable(false)]
        [BezKolumny]
        public string PaszportWydany
        {
            get => paszportWydany;
            set => SetProperty(ref paszportWydany, value);
        }

        [Browsable(false)]
        [BezKolumny]
        public string PaszportWazny
        {
            get => paszportWazny;
            set => SetProperty(ref paszportWazny, value);
        }

        [Browsable(false)]
        [BezKolumny]
        public string Ulica
        {
            get => ulica;
            set => SetProperty(ref ulica, value);
        }

        [Browsable(false)]
        [BezKolumny]
        public string NrDomu
        {
            get => nrDomu;
            set => SetProperty(ref nrDomu, value);
        }

        [Browsable(false)]
        [BezKolumny]
        public string Miejscowosc
        {
            get => miejscowosc;
            set => SetProperty(ref miejscowosc, value);
        }

        [Browsable(false)]
        [BezKolumny]
        public string KodPocztowy
        {
            get => kodPocztowy;
            set => SetProperty(ref kodPocztowy, value);
        }

        [Browsable(false)]
        [BezKolumny]
        public string Gmina
        {
            get => gmina;
            set => SetProperty(ref gmina, value);
        }

        [Browsable(false)]
        [BezKolumny]
        public string Powiat
        {
            get => powiat;
            set => SetProperty(ref powiat, value);
        }

        [Browsable(false)]
        [BezKolumny]
        public string Wojewodztwo
        {
            get => wojewodztwo;
            set => SetProperty(ref wojewodztwo, value);
        }

        [Browsable(false)]
        [BezKolumny]
        public string Kraj
        {
            get => kraj;
            set => SetProperty(ref kraj, value);
        }

        [Browsable(false)]
        [BezKolumny]
        public string Telefon
        {
            get => telefon;
            set => SetProperty(ref telefon, value);
        }

        [Browsable(false)]
        [BezKolumny]
        public string Email
        {
            get => email;
            set => SetProperty(ref email, value);
        }

        [Browsable(false)]
        [BezKolumny]
        public string BankStr
        {
            get => bankStr;
            set => SetProperty(ref bankStr, value);
        }

        [Browsable(false)]
        [BezKolumny]
        public string NfzStr
        {
            get => nfzStr;
            set => SetProperty(ref nfzStr, value);
        }

        [Browsable(false)]
        [BezKolumny]
        public string TypUmowy
        {
            get => typUmowy;
            set
            {
                if (SetProperty(ref typUmowy, value))
                {
                    OnPropertyChanged(nameof(OpisUmowy));
                }
            }
        }

        [Browsable(false)]
        [BezKolumny]
        public string DataRozpoczecia
        {
            get => dataRozpoczecia;
            set
            {
                if (SetProperty(ref dataRozpoczecia, value))
                {
                    OnPropertyChanged(nameof(OpisUmowy));
                }
            }
        }

        [Browsable(false)]
        [BezKolumny]
        public string DataZakonczenia
        {
            get => dataZakonczenia;
            set
            {
                if (SetProperty(ref dataZakonczenia, value))
                {
                    OnPropertyChanged(nameof(OpisUmowy));
                }
            }
        }

        [Browsable(false)]
        [BezKolumny]
        public string DataZawarcia
        {
            get => dataZawarcia;
            set => SetProperty(ref dataZawarcia, value);
        }

        [Browsable(false)]
        [BezKolumny]
        public string KodZawodu
        {
            get => kodZawodu;
            set => SetProperty(ref kodZawodu, value);
        }

        [Browsable(false)]
        [BezKolumny]
        public string JestStudentem
        {
            get => jestStudentem;
            set => SetProperty(ref jestStudentem, value);
        }

        // Visible properties
        [DisplayName("Pracownik")]
        public string PracownikNazwa => $"{Nazwisko} {Imie}".Trim();

        [DisplayName("PESEL")]
        public string Pesel
        {
            get => pesel;
            set => SetProperty(ref pesel, value);
        }

        [DisplayName("Umowa")]
        public string OpisUmowy => $"{TypUmowy} ({DataRozpoczecia} - {DataZakonczenia})";

        [DisplayName("Dział")]
        public string DzialNazwa
        {
            get => dzialNazwa;
            set => SetProperty(ref dzialNazwa, value);
        }

        [DisplayName("Status")]
        public string Status
        {
            get => status;
            set => SetProperty(ref status, value);
        }

        [DisplayName("Ostrzeżenia")]
        public string Ostrzezenia
        {
            get => ostrzezenia;
            set => SetProperty(ref ostrzezenia, value);
        }

        // WPF Data-Binding (INotifyPropertyChanged)
        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        protected bool SetProperty<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
        {
            if (Equals(field, value)) return false;
            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }

        // WPF Validation (IDataErrorInfo)
        [Browsable(false)]
        [BezKolumny]
        public string? Error => null;

        [Browsable(false)]
        [BezKolumny]
        public string? this[string columnName]
        {
            get
            {
                if (columnName == nameof(Imie))
                {
                    if (string.IsNullOrWhiteSpace(Imie))
                        return "Imię jest wymagane.";
                }
                else if (columnName == nameof(Nazwisko))
                {
                    if (string.IsNullOrWhiteSpace(Nazwisko))
                        return "Nazwisko jest wymagane.";
                }
                else if (columnName == nameof(Pesel))
                {
                    if (!string.IsNullOrWhiteSpace(Pesel))
                    {
                        if (!PeselHelper.ValidatePesel(Pesel, out _, out _))
                            return "Nieprawidłowy PESEL lub suma kontrolna.";
                    }
                }
                return null;
            }
        }

        public override string ToString()
        {
            return $"{Imie} {Nazwisko}";
        }
    }
}
