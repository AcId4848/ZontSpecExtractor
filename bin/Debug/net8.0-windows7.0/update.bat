@echo off
timeout /t 2 /nobreak >nul
taskkill /F /IM "ZontSpecExtractor.exe" >nul 2>&1
timeout /t 1 /nobreak >nul
if exist "C:\Users\JUSTI\AppData\Local\Temp\ZontSpecExtractor_Update\app_settings.json" (
    del /F /Q "C:\Users\JUSTI\AppData\Local\Temp\ZontSpecExtractor_Update\app_settings.json" >nul 2>&1
)
xcopy /Y /E /I "C:\Users\JUSTI\AppData\Local\Temp\ZontSpecExtractor_Update" "C:\Users\JUSTI\Desktop\v3ZontSpec\bin\Debug\net8.0-windows7.0" >nul 2>&1
if exist "C:\Users\JUSTI\AppData\Local\Temp\ZontSpecExtractor_Update\app_settings.json.backup" (
    copy /Y "C:\Users\JUSTI\AppData\Local\Temp\ZontSpecExtractor_Update\app_settings.json.backup" "C:\Users\JUSTI\Desktop\v3ZontSpec\bin\Debug\net8.0-windows7.0\app_settings.json" >nul
)
start "" "C:\Users\JUSTI\Desktop\v3ZontSpec\bin\Debug\net8.0-windows7.0\ZontSpecExtractor.exe"
rmdir /S /Q "C:\Users\JUSTI\AppData\Local\Temp\ZontSpecExtractor_Update" >nul 2>&1
