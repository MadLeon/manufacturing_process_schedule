@echo off
chcp 65001 >nul
echo ======================================
echo Tolerance Converter - ?????
echo ======================================
echo.

echo [1/4] ?????...
dotnet clean ToleranceConverter.csproj -c Release
echo.

echo [2/4] ?????...
dotnet publish ToleranceConverter.csproj -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true /p:DebugType=None /p:DebugSymbols=false
echo.

echo [3/4] ?????...
powershell -Command "Compress-Archive -Path 'bin\Release\net6.0-windows\win-x64\publish\ToleranceConverter.exe' -DestinationPath 'ToleranceConverter_v1.0_Win-x64.zip' -Force"
echo.

echo [4/4] ??????...
powershell -Command "Get-ChildItem 'bin\Release\net6.0-windows\win-x64\publish\ToleranceConverter.exe' | Select-Object Name, @{Name='Size(MB)';Expression={[math]::Round($_.Length/1MB, 2)}}"
powershell -Command "Get-ChildItem 'ToleranceConverter_v1.0_Win-x64.zip' | Select-Object Name, @{Name='Size(MB)';Expression={[math]::Round($_.Length/1MB, 2)}}"
echo.

echo ======================================
echo ????!
echo ======================================
echo ??? EXE: bin\Release\net6.0-windows\win-x64\publish\ToleranceConverter.exe
echo ???:     ToleranceConverter_v1.0_Win-x64.zip
echo.

pause
