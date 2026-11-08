using System.Linq;

namespace InsERT.Mox.QuickInstallerCore
{
	public static class DeployParametersValidator
	{
		public static bool Validate(this DeployParameters deployParameters, out string errorMessage)
		{
			errorMessage = null;

			if (deployParameters == null)
			{
				errorMessage = "Brak parametrów instalacji.";
				return false;
			}
			if (string.IsNullOrWhiteSpace(deployParameters.SqlServer))
			{
				errorMessage = "Parametry nie zawierają nazwy serwera SQL";
				return false;
			}
			if (!deployParameters.WindowsAuthentication &&
				string.IsNullOrWhiteSpace(deployParameters.User))
			{
				errorMessage = "Parametry nie zawierają nazwy użytkownika, wymaganej do połączenia z serwerem SQL";
				return false;
			}
			if (deployParameters.Databases == null || deployParameters.Databases.Count() == 0)
			{
				errorMessage = "Parametry nie zawierają nazw baz danych, w których ma zostać zainstalowany pakiet";
				return false;
			}

			return true;
		}
	}
}