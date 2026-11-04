using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Linq;
using InsERT.Moria.Raporty;
using MiniExcelLibs.Attributes;

namespace HrAppka_Import_Pracowników
{
    public enum DecyzjaIstniejacyPracownik { AktualizujKartotekeIDodajUmowe, TylkoDodajUmowe, NadpiszUmowe, PominWiersz }
    public class PracownikImportRow : INotifyPropertyChanged, IDataErrorInfo
    {
        [Browsable(false)]
        [BezKolumny]
        [ExcelColumnName("Pracownik")]
        public string PracownikCol { get; set; } = "";

        [Browsable(false)]
        [BezKolumny]
        public DateTime? ResolvedBirthDate { get; set; }

        [Browsable(false)]
        [BezKolumny]
        public string? ResolvedGender { get; set; }

        private bool isSelected = true;
        private object? selectedSzablonUmowy;
        private InsERT.Moria.ModelDanych.OrganPodatkowy? selectedOrganPodatkowy;
        private DecyzjaIstniejacyPracownik decyzjaDlaIstniejacego = DecyzjaIstniejacyPracownik.AktualizujKartotekeIDodajUmowe;

        [DisplayName("Wybór")]
        public bool IsSelected
        {
            get => isSelected;
            set => SetProperty(ref isSelected, value);
        }

        [Browsable(false)]
        [BezKolumny]
        public object? SelectedSzablonUmowy
        {
            get => selectedSzablonUmowy;
            set => SetProperty(ref selectedSzablonUmowy, value);
        }

        [Browsable(false)]
        [BezKolumny]
        public InsERT.Moria.ModelDanych.OrganPodatkowy? SelectedOrganPodatkowy
        {
            get => selectedOrganPodatkowy;
            set
            {
                if (SetProperty(ref selectedOrganPodatkowy, value))
                {
                    OnPropertyChanged(nameof(StatusLevel));
                }
            }
        }

        [Browsable(false)]
        [BezKolumny]
        public DecyzjaIstniejacyPracownik DecyzjaDlaIstniejacego
        {
            get => decyzjaDlaIstniejacego;
            set
            {
                if (SetProperty(ref decyzjaDlaIstniejacego, value))
                {
                    if (value == DecyzjaIstniejacyPracownik.PominWiersz && IsDuplicate)
                    {
                        Status = "Pominięty (Duplikat)";
                    }
                    else if (IsDuplicate && Status == "Pominięty (Duplikat)")
                    {
                        Status = "Istnieje w bazie";
                    }
                    OnPropertyChanged(nameof(StatusLevel));
                    OnPropertyChanged(nameof(IsDuplicate));
                }
            }
        }

        [Browsable(false)]
        [BezKolumny]
        public bool IsDuplicate => (Status != null && Status.Contains("Istnieje", StringComparison.OrdinalIgnoreCase)) ||
                                   (Status != null && Status.Contains("Duplikat", StringComparison.OrdinalIgnoreCase)) ||
                                   (Ostrzezenia != null && Ostrzezenia.Contains("już istnieje", StringComparison.OrdinalIgnoreCase));

        [Browsable(false)]
        [BezKolumny]
        public RowStatusLevel StatusLevel
        {
            get
            {
                // 1. Critical Errors
                if (string.IsNullOrWhiteSpace(Imie) || string.IsNullOrWhiteSpace(Nazwisko))
                    return RowStatusLevel.CriticalError;

                if (string.IsNullOrWhiteSpace(Pesel) && string.IsNullOrWhiteSpace(Paszport) && string.IsNullOrWhiteSpace(Dowod))
                    return RowStatusLevel.CriticalError;

                if (!string.IsNullOrWhiteSpace(Pesel) && !PeselHelper.ValidatePesel(Pesel, out _, out _))
                    return RowStatusLevel.CriticalError;

                if (Status != null && (
                    Status.StartsWith("Błąd", StringComparison.OrdinalIgnoreCase) ||
                    Status.StartsWith("Błędny", StringComparison.OrdinalIgnoreCase) ||
                    Status.Contains("Brak PESEL", StringComparison.OrdinalIgnoreCase)))
                {
                    return RowStatusLevel.CriticalError;
                }

                // 2. Warnings / Conflicts
                if (IsDuplicate || 
                    !string.IsNullOrWhiteSpace(Ostrzezenia) ||
                    (SelectedOrganPodatkowy == null && !string.IsNullOrWhiteSpace(UrzadSkarbowy)))
                {
                    return RowStatusLevel.Warning;
                }

                // 3. Ready
                return RowStatusLevel.Ready;
            }
        }
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
        private string dowod = "";
        private string dowodWydany = "";
        private string dowodWazny = "";
        private string paszportBio = "";
        private string paszportBioWydany = "";
        private string paszportBioWazny = "";
        private string ulica = "";
        private string nrDomu = "";
        private string nrLokalu = "";
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
                    OnPropertyChanged(nameof(StatusLevel));
                    OnPropertyChanged(nameof(IsDuplicate));
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
                    OnPropertyChanged(nameof(StatusLevel));
                    OnPropertyChanged(nameof(IsDuplicate));
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
            set
            {
                if (SetProperty(ref paszport, value))
                {
                    if (!string.IsNullOrWhiteSpace(paszport) && (status == "Brak PESEL i Paszportu" || string.IsNullOrWhiteSpace(status)))
                    {
                        Status = "Gotowy (Cudzoziemiec)";
                    }
                    OnPropertyChanged(nameof(StatusLevel));
                    OnPropertyChanged(nameof(IsDuplicate));
                }
            }
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
        public string Dowod
        {
            get => dowod;
            set
            {
                if (SetProperty(ref dowod, value))
                {
                    if (!string.IsNullOrWhiteSpace(dowod) && (status == "Brak PESEL i Paszportu" || string.IsNullOrWhiteSpace(status)))
                    {
                        Status = "Gotowy";
                    }
                    OnPropertyChanged(nameof(StatusLevel));
                    OnPropertyChanged(nameof(IsDuplicate));
                }
            }
        }

        [Browsable(false)]
        [BezKolumny]
        public string DowodWydany
        {
            get => dowodWydany;
            set => SetProperty(ref dowodWydany, value);
        }

        [Browsable(false)]
        [BezKolumny]
        public string DowodWazny
        {
            get => dowodWazny;
            set => SetProperty(ref dowodWazny, value);
        }

        [Browsable(false)]
        [BezKolumny]
        public string PaszportBio
        {
            get => paszportBio;
            set => SetProperty(ref paszportBio, value);
        }

        [Browsable(false)]
        [BezKolumny]
        public string PaszportBioWydany
        {
            get => paszportBioWydany;
            set => SetProperty(ref paszportBioWydany, value);
        }

        [Browsable(false)]
        [BezKolumny]
        public string PaszportBioWazny
        {
            get => paszportBioWazny;
            set => SetProperty(ref paszportBioWazny, value);
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
        public string NrLokalu
        {
            get => nrLokalu;
            set => SetProperty(ref nrLokalu, value);
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
            set
            {
                if (SetProperty(ref pesel, value))
                {
                    if (!string.IsNullOrWhiteSpace(pesel) && PeselHelper.ValidatePesel(pesel, out var bd, out var g))
                    {
                        ResolvedBirthDate = bd;
                        ResolvedGender = g;
                        if (status == "Błędny PESEL" || status == "Brak PESEL i Paszportu")
                        {
                            Status = "Gotowy";
                        }
                    }
                    OnPropertyChanged(nameof(StatusLevel));
                    OnPropertyChanged(nameof(IsDuplicate));
                }
            }
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
            set
            {
                if (SetProperty(ref status, value))
                {
                    OnPropertyChanged(nameof(StatusLevel));
                    OnPropertyChanged(nameof(IsDuplicate));
                }
            }
        }

        [DisplayName("Ostrzeżenia")]
        public string Ostrzezenia
        {
            get => ostrzezenia;
            set
            {
                if (SetProperty(ref ostrzezenia, value))
                {
                    OnPropertyChanged(nameof(StatusLevel));
                    OnPropertyChanged(nameof(IsDuplicate));
                }
            }
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

        // WPF Validation (IDataErrorInfo) Explicit Interface Implementation
        string IDataErrorInfo.Error => string.Empty;

        string IDataErrorInfo.this[string columnName]
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
                return string.Empty;
            }
        }

        public override string ToString()
        {
            return $"{Imie} {Nazwisko}";
        }
    }
}
