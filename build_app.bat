@echo off
echo =======================================================
echo HAI-TECH FORM EXTRACTOR - EXECUTABLE BUILDER
echo =======================================================
echo.
echo Installing PyInstaller...
pip install pyinstaller

echo.
echo Building the executable...
echo This may take 5-10 minutes depending on your system. Please wait...
pyinstaller --noconfirm --windowed --icon "icon.ico" --name "FormExtractor" --collect-all customtkinter --collect-all pillow_heif --add-data ".env;." --add-data "icon.ico;." app.py

echo.
echo =======================================================
echo BUILD COMPLETE!
echo =======================================================
echo You can find your packaged application inside the 'dist/FormExtractor' folder.
echo You can Zip that folder and share it with your friend!
echo.
pause
