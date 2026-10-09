@echo off
setlocal
cd /d "%~dp0"

rem ============================================================================
rem  One fell swoop:  StatementOrganizer (console)  ->  OneNoteSync (GUI)
rem
rem  1. StatementOrganizer  - LLM-classifies every PDF in input\, copies them
rem                           into organized\, and writes organized\statements.json
rem  2. OneNoteSync         - reads organized\statements.json and opens the GUI;
rem                           you then click "Run sync" when ready.
rem ============================================================================

echo.
echo ============================================================================
echo  One fell swoop:  StatementOrganizer  -^>  OneNoteSync
echo ============================================================================

echo.
echo [1/3] Building StatementOrganizer...
dotnet build StatementOrganizer\StatementOrganizer.csproj -v q
if errorlevel 1 (
    echo.
    echo StatementOrganizer build failed - aborting.
    pause
    exit /b 1
)

echo.
echo [2/3] Running StatementOrganizer  (seeding organized\statements.json)...
StatementOrganizer\bin\Debug\net9.0\StatementOrganizer.exe
if errorlevel 1 (
    echo.
    echo StatementOrganizer FAILED - check the output above and the LLM_API_KEY in .env
    echo Not launching OneNoteSync.
    pause
    exit /b 1
)

echo.
echo [3/3] Building OneNoteSync (best-effort), then launching...
dotnet build OneNoteSync\OneNoteSync.csproj -v q >nul 2>&1
if errorlevel 1 (
    echo   OneNoteSync build skipped - it is probably running; using the existing exe.
)
if not exist "OneNoteSync\bin\Debug\OneNoteSync.exe" (
    echo OneNoteSync\bin\Debug\OneNoteSync.exe not found - build it first.
    pause
    exit /b 1
)
start "" "OneNoteSync\bin\Debug\OneNoteSync.exe"
echo OneNoteSync launched - review the mapping and click "Run sync" when ready.
exit /b 0
