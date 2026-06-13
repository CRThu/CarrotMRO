@echo off
echo Building frontend...
cd frontend
call bun install
call bun run build
cd ..

echo Cleaning backend static files...
if exist "backend\static" rmdir /s /q "backend\static"
mkdir "backend\static"

echo Copying frontend build to backend...
xcopy "frontend\dist\*" "backend\static\" /e /i /h /y

echo Build complete.