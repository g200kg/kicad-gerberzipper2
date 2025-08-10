echo %USERPROFILE%
md %USERPROFILE%\Documents\KiCad\9.0\plugins\gerberzipper2
copy .\plugins\gerberzipper2.py %USERPROFILE%\Documents\KiCad\9.0\plugins\gerberzipper2\gerberzipper2.py
copy .\plugins\plugin.json %USERPROFILE%\Documents\KiCad\9.0\plugins\gerberzipper2\plugin.json
copy .\plugins\requirements.txt %USERPROFILE%\Documents\KiCad\9.0\plugins\gerberzipper2\requirements.txt
md %USERPROFILE%\Documents\KiCad\9.0\plugins\gerberzipper2\Assets
copy .\plugins\Assets\*.png %USERPROFILE%\Documents\KiCad\9.0\plugins\gerberzipper2\Assets
md %USERPROFILE%\Documents\KiCad\9.0\plugins\gerberzipper2\Locale
copy .\plugins\Locale\*.json %USERPROFILE%\Documents\KiCad\9.0\plugins\gerberzipper2\Locale
md %USERPROFILE%\Documents\KiCad\9.0\plugins\gerberzipper2\Manufacturers
copy .\plugins\Manufacturers\*.json %USERPROFILE%\Documents\KiCad\9.0\plugins\gerberzipper2\Manufacturers
