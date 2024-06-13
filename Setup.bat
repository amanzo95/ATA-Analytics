@echo off
setlocal

REM Ruta donde se descargará y descomprimirá el repositorio temporalmente
set REPO_DIR=C:\ATA-Analytics-TEMP
set REPO_DIR_FINAL=C:\ATA-Analytics
set EXE_PATH_FINAL=%REPO_DIR_FINAL%\net8.0-windows\GK_Analytics.exe

REM URL del archivo ZIP del repositorio de GitHub
set REPO_URL=https://github.com/amanzo95/ATA-Analytics/archive/refs/heads/main.zip

REM Ruta del archivo ZIP descargado
set ZIP_PATH=%TEMP%\repo.zip

REM Comprueba si la carpeta temporal ya existe y la elimina si es necesario
if exist %REPO_DIR% (
    echo La carpeta temporal %REPO_DIR% ya existe. Eliminando carpeta...
    rmdir /S /Q %REPO_DIR%
)

REM Descargar el archivo ZIP del repositorio
echo Descargando el repositorio...
curl -L %REPO_URL% --output %ZIP_PATH%

REM Verificar si la descarga fue exitosa
if %errorlevel% neq 0 (
    echo Error: No se pudo descargar el repositorio.
    pause
    exit /b 1
)

REM Crear las carpetas necesarias
mkdir %REPO_DIR%

REM Descomprimir el archivo ZIP en la carpeta temporal del repositorio
echo Descomprimiendo el archivo ZIP...
powershell -Command "Expand-Archive -Path '%ZIP_PATH%' -DestinationPath '%REPO_DIR%'"

REM Verificar si la descomprimión fue exitosa
if %errorlevel% neq 0 (
    echo Error: No se pudo descomprimir el archivo ZIP.
    pause
    exit /b 1
)

REM Definir la ruta del archivo ejecutable en el repositorio descargado
set SOURCE_EXE_PATH=%REPO_DIR%\ATA-Analytics-main\ATA_Analisis\bin\Debug\net8.0-windows\GK_Analytics.exe

REM Verificar si el archivo ejecutable existe en el repositorio descargado
if not exist %SOURCE_EXE_PATH% (
    echo Error: El archivo ejecutable %SOURCE_EXE_PATH% no existe.
    pause
    exit /b 1
)

REM Verificar si el archivo ejecutable existe en el destino final
if exist %EXE_PATH_FINAL% (
    echo Verificando si los archivos son iguales...

    REM Calcular el hash del archivo fuente y destino
    certutil -hashfile %SOURCE_EXE_PATH% SHA256 > %TEMP%\hash_source.txt
    certutil -hashfile %EXE_PATH_FINAL% SHA256 > %TEMP%\hash_dest.txt

    REM Leer los hashes calculados
    set /p HASH_SOURCE=<%TEMP%\hash_source.txt
    set /p HASH_DEST=<%TEMP%\hash_dest.txt

    REM Comparar los hashes
    if "%HASH_SOURCE%" == "%HASH_DEST%" (
        echo Los archivos son iguales. No se realiza ninguna acción.
        
        REM Limpiar archivos temporales y carpetas
        del %ZIP_PATH%
        rmdir /S /Q %REPO_DIR%
        del %TEMP%\hash_source.txt 2>nul
        del %TEMP%\hash_dest.txt 2>nul
        
        endlocal
        pause
        exit /b 0
    ) else (
        echo Los archivos son diferentes. Procediendo con la copia...
    )
)

REM Crear la carpeta de destino final si no existe
if not exist %REPO_DIR_FINAL% (
    mkdir %REPO_DIR_FINAL%
)

REM Copiar la carpeta específica al directorio de destino final
echo Copiando la carpeta especifica...
set SOURCE_DIR=%REPO_DIR%\ATA-Analytics-main\ATA_Analisis\bin\Debug\net8.0-windows
set DEST_DIR=%REPO_DIR_FINAL%

xcopy /E /I /Y %SOURCE_DIR% %DEST_DIR%
echo Carpeta copiada exitosamente a %DEST_DIR%.

REM Eliminar el directorio descomprimido temporal
rmdir /S /Q %REPO_DIR%

REM Eliminar el archivo ZIP descargado y archivos temporales de hash
del %ZIP_PATH%
del %TEMP%\hash_source.txt 2>nul
del %TEMP%\hash_dest.txt 2>nul

echo Proceso completado exitosamente.
endlocal
pause
