using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace InsERT.Mox.QuickInstallerCore
{
	internal static class DeployParametersParser
	{
		private const string ParamValueRegex = @"^(?<param>\w*)\s*=(?<value>.*)";

		internal static DeployParameters Parse(params string[] paramLines)
		{
			var deployParameters = new DeployParameters();

			var parameters = MatchParams(paramLines);

			deployParameters.SqlServer = GetParamValueOrDefault(parameters, "Serwer");
			deployParameters.WindowsAuthentication =
				string.Equals("Tak", GetParamValueOrDefault(parameters, "UwierzytelnianieWindows"), System.StringComparison.InvariantCultureIgnoreCase);
			deployParameters.User = GetParamValueOrDefault(parameters, "Uzytkownik");
			deployParameters.Password = GetParamValueOrDefault(parameters, "Haslo");

			var keys = parameters.Keys.Where(k => k.ParamName == "BazaDanych");
			if (keys.Any())
			{
				deployParameters.Databases = [];
				foreach (var key in keys)
				{
					string baza = parameters[key];
					if (!string.IsNullOrWhiteSpace(baza))
					{
						deployParameters.Databases.Add(baza);
					}
				}
			}

			return deployParameters;
		}

		private static Dictionary<(string ParamName, int Index), string> MatchParams(string[] paramLines)
		{
			Dictionary<(string ParamName, int Index), string> parameters = [];
			for (int i = 0; i < paramLines.Length; i++)
			{
				string line = paramLines[i];
				if (string.IsNullOrWhiteSpace(line))
				{
					continue;
				}

				var match = Regex.Match(line, ParamValueRegex);
				if (!match.Success)
				{
					continue;
				}

				parameters.Add((match.Groups["param"].Value, i), match.Groups["value"].Value);
			}

			return parameters;
		}

		private static string GetParamValueOrDefault(Dictionary<(string ParamName, int Index), string> parameters, string paramName)
		{
			if (parameters.Keys.Any(k => k.ParamName == paramName))
			{
				var key = parameters.Keys.First(k => k.ParamName == paramName);
				return parameters[key];
			}

			return null;
		}
	}
}