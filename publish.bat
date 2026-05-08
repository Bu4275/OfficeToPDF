@echo off
echo Publishing PDF Converter (self-contained, single file)...
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:EnableCompressionInSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true
echo.
echo Done! Output: bin\Release\net9.0-windows\win-x64\publish\
pause
