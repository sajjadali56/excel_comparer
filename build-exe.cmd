@REM echo "Clean previous builds (alternative)"
if exist dist rmdir /s /q dist
if exist build rmdir /s /q build

@REM echo "Build with clean spec"
pyinstaller .\GENERATE_EXE.spec --clean --noconfirm

@REM Test
@REM cd dist\EOS_Gratuity_Tool_V3
@REM EOS_Gratuity_Tool_V3.exe