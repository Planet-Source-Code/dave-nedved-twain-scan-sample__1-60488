@echo off
:start
@cls
@echo Initiating setup for TWAIN Sample...
@echo Complete.
@echo -----------------------------------------------------------------------
@pause
@ren "Setup\~187394.bin" "EZTW32.TXT"
@ren "Setup\~843823.bin" "EZTW32.DLL"
@ren "Setup\~623463.bin" "Twain Scanner.exe"
@ren "Setup\~743238.bin" "Twain Scanner.exe.manifest"
@copy "Setup\EZTW32.TXT" "Source\EZTW32.TXT"
@copy "Setup\EZTW32.DLL" "Source\EZTW32.DLL"
@copy "Setup\Twain Scanner.exe" "Source\Twain Scanner.exe"
@copy "Setup\Twain Scanner.exe.manifest" "Source\Twain Scanner.exe.manifest"
@ren "Setup\EZTW32.TXT" "~187394.bin"
@ren "Setup\EZTW32.DLL" "~843823.bin"
@ren "Setup\Twain Scanner.exe" "~623463.bin"
@ren "Setup\Twain Scanner.exe.manifest" "~743238.bin"
@echo -----------------------------------------------------------------------
@cls
:end
@echo -----------------------------------------------------------------------
@echo Thanks for using This Sample In case of problems, suggestions, bug -
@echo reports, please report them to the author via email.
@echo David Nedved
@echo dnedved@datosoftware.com
@echo Making your Programming Life Easier...
@echo -----------------------------------------------------------------------
@pause
@echo on
54481999