@echo off
REM Sync automatico do dashboard de midia (API Google+Meta -> cache -> HTML -> push).
REM Rodado pela tarefa agendada Dashboard-Midia-Sync. Caminhos relativos ao .bat (%~dp0).
cd /d "%~dp0"
set "DASH_XL_PATH=%~dp0data\dash_data.xlsx"
set "DASH_HTML_OUT=%~dp0deploy\index.html"
set "LOG=%~dp0sync.log"

echo ===================== %date% %time% ===================== >> "%LOG%"
C:\Python313\python.exe scripts\dash_sync.py --incremental 7 >> "%LOG%" 2>&1
C:\Python313\python.exe scripts\fetch_ads_detail.py --incremental 7 >> "%LOG%" 2>&1
C:\Python313\python.exe gerar_dashboard_midia.py >> "%LOG%" 2>&1
git -C "%~dp0deploy" add -A >> "%LOG%" 2>&1
git -C "%~dp0deploy" commit -m "Atualizacao automatica do dashboard" >> "%LOG%" 2>&1
git -C "%~dp0deploy" push origin main >> "%LOG%" 2>&1
echo ----- fim %time% ----- >> "%LOG%"
