@echo off
setlocal
cd /d "%~dp0"

echo Running QA Salesforce login ^(single test^) from:
echo %CD%
echo Mode: 1 worker, Chromium only, headed ^(visible browser^)
echo.

call npx playwright test tests/salesforce-login.spec.ts --workers=1 --project=chromium --headed
set EXITCODE=%ERRORLEVEL%

echo.
if %EXITCODE% neq 0 (
  echo Tests finished with errors ^(exit code %EXITCODE%^).
) else (
  echo Tests finished successfully.
)
pause
exit /b %EXITCODE%
