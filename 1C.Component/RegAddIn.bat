@ECHO OFF
set DOTNETFX4=%SystemRoot%\Microsoft.NET\Framework\v4.0.30319
echo ---------------------------------------------------
%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\regasm.exe 1C.Component.dll /codebase
echo ---------------------------------------------------
pause