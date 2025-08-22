@echo off
ECHO Iniciando Comparador SUA con PM2...

:: Verifica si PM2 está instalado
where pm2 >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    ECHO Error: PM2 no está instalado. Instalándolo globalmente...
    npm install -g pm2
    IF %ERRORLEVEL% NEQ 0 (
        ECHO Error: No se pudo instalar PM2. Asegúrate de tener npm instalado.
        pause
        exit /b 1
    )
)

:: Inicia la aplicación con PM2
pm2 start server.js --name "ComparadorSUA"

:: Guarda la lista de procesos para reinicio automático
pm2 save

:: Muestra el estado de los procesos
pm2 list

ECHO ¡Aplicación iniciada! Presiona cualquier tecla para salir.
pause