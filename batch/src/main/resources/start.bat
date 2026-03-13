@echo off
setlocal

REM 当前脚本目录
set BASE_DIR=%~dp0

REM JAR 文件
set JAR_NAME=batch-0.0.1-SNAPSHOT.jar

REM 配置文件
set CONFIG_FILE=%BASE_DIR%application.yml

echo ================================
echo Starting batch application...
echo JAR: %BASE_DIR%%JAR_NAME%
echo CONFIG: %CONFIG_FILE%
echo ================================

java -jar "%BASE_DIR%%JAR_NAME%" ^
--spring.config.location="%CONFIG_FILE%"

pause