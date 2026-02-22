@echo off
title Configuration rclone pour Google Drive
echo ============================================
echo   Configuration de rclone pour Google Drive
echo ============================================
echo.
echo Suivez ces instructions:
echo.
echo 1. Tapez 'n' pour nouveau remote
echo 2. Nom: gdrive
echo 3. Storage: tapez 'drive' ou le numero Google Drive
echo 4-8. Laissez vide (appuyez sur Entree)
echo 9. Edit advanced config: n
echo 10. Use auto config: y (navigateur s'ouvrira)
echo 11. Configure as team drive: n
echo 12. Confirmez avec 'y', puis 'q' pour quitter
echo.
echo ============================================
echo.

"C:\Users\Philippe\AppData\Local\Microsoft\WinGet\Packages\Rclone.Rclone_Microsoft.Winget.Source_8wekyb3d8bbwe\rclone-v1.73.0-windows-amd64\rclone.exe" config

echo.
echo Configuration terminee!
pause
