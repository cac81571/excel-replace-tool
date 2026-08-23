@echo off
setlocal
cd /d "%~dp0"
if not exist "target\excel-replace-tool-1.0.0-all.jar" (
  echo 先に mvn package を実行してください。
  exit /b 1
)
java -jar "target\excel-replace-tool-1.0.0-all.jar"
