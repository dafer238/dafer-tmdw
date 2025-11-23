@echo off
setlocal enabledelayedexpansion

echo ========================================
echo Building CoolPropWrapper
echo ========================================
echo.

REM Check if dotnet is installed
dotnet --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: .NET SDK is not installed or not in PATH
    echo Please install .NET SDK from https://dotnet.microsoft.com/download
    exit /b 1
)

REM Clean previous builds
echo Cleaning previous builds...
if exist bin rmdir /s /q bin
if exist obj rmdir /s /q obj

REM Restore packages
echo.
echo Restoring NuGet packages...
dotnet restore
if errorlevel 1 (
    echo ERROR: Failed to restore packages
    exit /b 1
)

REM Build for .NET Framework 4.8 (x64)
echo.
echo ========================================
echo Building for .NET Framework 4.8 (x64)
echo ========================================
echo.
echo NOTE: ExcelDNA may show errors about missing DLL during packing.
echo This is normal and can be ignored if the build completes.
echo.
dotnet build -c Release -p:Platform=x64

REM Check if the main DLL was created (this is what matters)
if not exist "bin\x64\Release\net48\CoolPropWrapper.dll" (
    echo.
    echo ========================================
    echo ERROR: Build failed!
    echo ========================================
    echo CoolPropWrapper.dll was not created.
    exit /b 1
)

echo.
echo ========================================
echo Build Successful!
echo ========================================

REM Create compiled directories if they don't exist
echo.
echo Creating output directories...
if not exist "compiled\net48" mkdir "compiled\net48"

REM Copy files to compiled\net48
echo.
echo Copying files to compiled\net48...

copy /Y "bin\x64\Release\net48\CoolPropWrapper.dll" "compiled\net48\" >nul
copy /Y "bin\x64\Release\net48\CoolPropWrapper.dna" "compiled\net48\" >nul
copy /Y "CoolProp.dll" "compiled\net48\" >nul

REM Find and copy the best .xll file available
set XLL_COPIED=0

if exist "bin\x64\Release\net48\CoolPropWrapper64.xll" (
    echo Copying CoolPropWrapper64.xll as CoolPropWrapper.xll...
    copy /Y "bin\x64\Release\net48\CoolPropWrapper64.xll" "compiled\net48\CoolPropWrapper.xll" >nul
    set XLL_COPIED=1
)

if exist "bin\x64\Release\net48\CoolPropWrapper.xll" (
    if !XLL_COPIED! == 0 (
        echo Copying CoolPropWrapper.xll...
        copy /Y "bin\x64\Release\net48\CoolPropWrapper.xll" "compiled\net48\CoolPropWrapper.xll" >nul
        set XLL_COPIED=1
    )
)

if exist "bin\x64\Release\net48\publish\CoolPropWrapper-packed.xll" (
    echo Found packed version, using that instead...
    copy /Y "bin\x64\Release\net48\publish\CoolPropWrapper-packed.xll" "compiled\net48\CoolPropWrapper.xll" >nul
    set XLL_COPIED=1
)

if !XLL_COPIED! == 0 (
    echo.
    echo ========================================
    echo WARNING: No .xll file found!
    echo ========================================
    echo The Excel add-in file was not created.
    echo Check the build output above for errors.
    pause
    exit /b 1
)

echo.
echo ========================================
echo Build Complete!
echo ========================================
echo.
echo Files copied to: compiled\net48\
echo.
echo Files included:
dir /b compiled\net48
echo.
echo You can now commit these files to git and create a release tag.
echo See RELEASE.md for instructions.
echo.

pause
