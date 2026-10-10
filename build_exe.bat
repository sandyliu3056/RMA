@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 安裝打包工具與套件...
py -3 -m pip install -r requirements.txt pyinstaller || python -m pip install -r requirements.txt pyinstaller
echo.
echo 開始打包（約 2-5 分鐘）...
py -3 -m PyInstaller --noconfirm --clean --onefile --windowed --name RMA_PartsPlanner --hidden-import 零件規劃平台 --hidden-import rma_engine --hidden-import i18n --hidden-import PIL.ImageTk --hidden-import PIL._tkinter_finder --collect-data pptx --collect-data matplotlib app_launcher.py || python -m PyInstaller --noconfirm --clean --onefile --windowed --name RMA_PartsPlanner --hidden-import 零件規劃平台 --hidden-import rma_engine --hidden-import i18n --hidden-import PIL.ImageTk --hidden-import PIL._tkinter_finder --collect-data pptx --collect-data matplotlib app_launcher.py
echo.
if exist dist\RMA_PartsPlanner.exe (
  echo 完成：dist\RMA_PartsPlanner.exe
  echo 把這個檔案複製到任何 Windows 電腦都能直接雙擊執行。
) else (
  echo 打包失敗，請把上面的訊息截圖。
)
pause
