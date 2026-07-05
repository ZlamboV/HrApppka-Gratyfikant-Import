@echo off
echo.
echo === AUTOMATIC INSTALLATION PROCESS ===
echo Building project and packaging plugin...
dotnet build -c Debug
if %ERRORLEVEL% neq 0 (
    echo Build failed!
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo Running automated C# installer utility...
dotnet run --project "HrAppka Import Pracowników.csproj" -- --install
if %ERRORLEVEL% neq 0 (
    echo Installation failed!
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo Plugin version 1.0.0.2 has been successfully installed and registered!
pause
