; Установщик книги рецептов для Windows. Собирается скриптом build_windows.sh.
; Права администратора не нужны: рецепты лежат рядом с программой, и ей нужна запись в свою папку.
Unicode true
SetCompressor /SOLID lzma
RequestExecutionLevel user

!include "MUI2.nsh"

!define APP_NAME "Книга рецептов"
!define UNINSTALL_KEY "Software\Microsoft\Windows\CurrentVersion\Uninstall\KnigaRetseptov"

Name "${APP_NAME}"
OutFile "${OUTFILE}"
InstallDir "$LOCALAPPDATA\${APP_NAME}"
InstallDirRegKey HKCU "${UNINSTALL_KEY}" "InstallLocation"
BrandingText "${APP_NAME}"

!define MUI_ICON "${SRC}\static\icon.ico"
!define MUI_UNICON "${SRC}\static\icon.ico"
!define MUI_ABORTWARNING

!define MUI_DIRECTORYPAGE_TEXT_TOP "Выберите папку для программы. Рецепты и фото будут храниться в ней же, поэтому нужна папка с правом записи: не выбирайте Program Files."
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!define MUI_FINISHPAGE_RUN
!define MUI_FINISHPAGE_RUN_FUNCTION OpenBook
!define MUI_FINISHPAGE_RUN_TEXT "Открыть книгу рецептов"
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

!insertmacro MUI_LANGUAGE "Russian"

; Работающая книга держит файлы программы, поэтому закрываем её. Рецепты при этом не страдают.
!macro CloseBook
  nsExec::Exec 'powershell -NoProfile -NonInteractive -Command "try { Invoke-WebRequest -UseBasicParsing -Method Post -TimeoutSec 2 http://127.0.0.1:8765/quit | Out-Null } catch {}"'
  Pop $0
  Sleep 1500
!macroend

Function OpenBook
  ; Запускаем из папки программы: она может содержать пробелы и кириллицу, поэтому пути в кавычках.
  SetOutPath "$INSTDIR"
  Exec '"$INSTDIR\python\pythonw.exe" "$INSTDIR\launcher.pyw"'
FunctionEnd

Section "Книга рецептов"
  !insertmacro CloseBook
  ; При обновлении заменяем только программу. Папка data с рецептами остаётся нетронутой.
  RMDir /r "$INSTDIR\python"
  RMDir /r "$INSTDIR\templates"
  RMDir /r "$INSTDIR\static"
  SetOutPath "$INSTDIR"
  File /r "${SRC}\*"

  ; Ярлык создаёт установщик, поэтому запускалке создавать его второй раз не нужно.
  CreateDirectory "$INSTDIR\data"
  FileOpen $0 "$INSTDIR\data\.shortcut-created" w
  FileClose $0
  CreateShortcut "$DESKTOP\${APP_NAME}.lnk" "$INSTDIR\python\pythonw.exe" '"$INSTDIR\launcher.pyw"' "$INSTDIR\static\icon.ico"
  CreateShortcut "$SMPROGRAMS\${APP_NAME}.lnk" "$INSTDIR\python\pythonw.exe" '"$INSTDIR\launcher.pyw"' "$INSTDIR\static\icon.ico"

  WriteUninstaller "$INSTDIR\Удалить программу.exe"
  WriteRegStr HKCU "${UNINSTALL_KEY}" "DisplayName" "${APP_NAME}"
  WriteRegStr HKCU "${UNINSTALL_KEY}" "DisplayIcon" "$INSTDIR\static\icon.ico"
  WriteRegStr HKCU "${UNINSTALL_KEY}" "InstallLocation" "$INSTDIR"
  WriteRegStr HKCU "${UNINSTALL_KEY}" "UninstallString" '"$INSTDIR\Удалить программу.exe"'
  WriteRegDWORD HKCU "${UNINSTALL_KEY}" "NoModify" 1
  WriteRegDWORD HKCU "${UNINSTALL_KEY}" "NoRepair" 1
SectionEnd

Section "Uninstall"
  !insertmacro CloseBook
  Delete "$DESKTOP\${APP_NAME}.lnk"
  Delete "$SMPROGRAMS\${APP_NAME}.lnk"
  RMDir /r "$INSTDIR\python"
  RMDir /r "$INSTDIR\templates"
  RMDir /r "$INSTDIR\static"
  RMDir /r "$INSTDIR\__pycache__"
  Delete "$INSTDIR\app.py"
  Delete "$INSTDIR\site_import.py"
  Delete "$INSTDIR\launcher.pyw"
  Delete "$INSTDIR\schema.sql"
  Delete "$INSTDIR\Книга рецептов.bat"
  Delete "$INSTDIR\Прочитай меня.txt"
  Delete "$INSTDIR\Примеры рецептов.json"
  Delete "$INSTDIR\Удалить программу.exe"
  DeleteRegKey HKCU "${UNINSTALL_KEY}"
  ; По умолчанию рецепты остаются: случайно нажатая кнопка не должна их стирать.
  MessageBox MB_YESNO|MB_ICONQUESTION|MB_DEFBUTTON2 "Удалить также все рецепты и фото?$\r$\n$\r$\nЭто нельзя отменить. Если нажать «Нет», они останутся в папке:$\r$\n$INSTDIR\data" /SD IDNO IDNO keep_data
    RMDir /r "$INSTDIR\data"
  keep_data:
  RMDir "$INSTDIR"
SectionEnd
