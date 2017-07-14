; Script generated by the HM NIS Edit Script Wizard.

; HM NIS Edit Wizard helper defines
!define PRODUCT_NAME "Perfect Lecture"
!define PRODUCT_VERSION "0.5"
!define PRODUCT_PUBLISHER "Cheng-Kuan Wei"
!define PRODUCT_WEB_SITE "https://github.com/kennywei815/PerfectLecture"
!define PRODUCT_DIR_REGKEY "Software\Microsoft\Windows\CurrentVersion\App Paths\convert.exe"
!define PRODUCT_UNINST_KEY "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}"
!define PRODUCT_UNINST_ROOT_KEY "HKLM"

; MUI 1.67 compatible ------
!include "MUI.nsh"

; MUI Settings
!define MUI_ABORTWARNING
!define MUI_ICON "${NSISDIR}\Contrib\Graphics\Icons\modern-install.ico"
!define MUI_UNICON "${NSISDIR}\Contrib\Graphics\Icons\modern-uninstall.ico"

; Welcome page
!insertmacro MUI_PAGE_WELCOME
; Directory page
!insertmacro MUI_PAGE_DIRECTORY
; Instfiles page
!insertmacro MUI_PAGE_INSTFILES
; Finish page
!define MUI_FINISHPAGE_SHOWREADME "$INSTDIR\安裝手冊 - Perfect Lecture.docx"
!insertmacro MUI_PAGE_FINISH

; Uninstaller pages
!insertmacro MUI_UNPAGE_INSTFILES

; Language files
!insertmacro MUI_LANGUAGE "English"

; MUI end ------

Name "${PRODUCT_NAME} ${PRODUCT_VERSION}"
OutFile "Setup_Perfect_Lecture.exe"
InstallDir "$PROGRAMFILES\Perfect Lecture"
InstallDirRegKey HKLM "${PRODUCT_DIR_REGKEY}" ""
ShowInstDetails show
ShowUnInstDetails show

Section "MainSection" SEC01
  SetOutPath "$APPDATA\Microsoft\AddIns\Perfect_Lecture"
  ;SetOverwrite try  ;Default
  SetOverwrite on
  File "Install\AppData\Roaming\Microsoft\AddIns\Perfect_Lecture\genPostProcessVBA.py"
  SetOutPath "$APPDATA\Microsoft\AddIns\Perfect_Lecture\ImageMagick-portable"
  File "Install\AppData\Roaming\Microsoft\AddIns\Perfect_Lecture\ImageMagick-portable\colors.xml"
  File "Install\AppData\Roaming\Microsoft\AddIns\Perfect_Lecture\ImageMagick-portable\configure.xml"
  File "Install\AppData\Roaming\Microsoft\AddIns\Perfect_Lecture\ImageMagick-portable\convert.exe"
  File "Install\AppData\Roaming\Microsoft\AddIns\Perfect_Lecture\ImageMagick-portable\delegates.xml"
  File "Install\AppData\Roaming\Microsoft\AddIns\Perfect_Lecture\ImageMagick-portable\english.xml"
  File "Install\AppData\Roaming\Microsoft\AddIns\Perfect_Lecture\ImageMagick-portable\ffmpeg.exe"
  File "Install\AppData\Roaming\Microsoft\AddIns\Perfect_Lecture\ImageMagick-portable\ImageMagick.rdf"
  File "Install\AppData\Roaming\Microsoft\AddIns\Perfect_Lecture\ImageMagick-portable\locale.xml"
  File "Install\AppData\Roaming\Microsoft\AddIns\Perfect_Lecture\ImageMagick-portable\log.xml"
  File "Install\AppData\Roaming\Microsoft\AddIns\Perfect_Lecture\ImageMagick-portable\magic.xml"
  File "Install\AppData\Roaming\Microsoft\AddIns\Perfect_Lecture\ImageMagick-portable\mime.xml"
  File "Install\AppData\Roaming\Microsoft\AddIns\Perfect_Lecture\ImageMagick-portable\policy.xml"
  File "Install\AppData\Roaming\Microsoft\AddIns\Perfect_Lecture\ImageMagick-portable\quantization-table.xml"
  File "Install\AppData\Roaming\Microsoft\AddIns\Perfect_Lecture\ImageMagick-portable\sRGB.icc"
  File "Install\AppData\Roaming\Microsoft\AddIns\Perfect_Lecture\ImageMagick-portable\thresholds.xml"
  File "Install\AppData\Roaming\Microsoft\AddIns\Perfect_Lecture\ImageMagick-portable\type-ghostscript.xml"
  File "Install\AppData\Roaming\Microsoft\AddIns\Perfect_Lecture\ImageMagick-portable\type.xml"
  SetOutPath "$APPDATA\Microsoft\AddIns\Perfect_Lecture"
  File "Install\AppData\Roaming\Microsoft\AddIns\Perfect_Lecture\pdf2mp4.py"
  File "Install\AppData\Roaming\Microsoft\AddIns\Perfect_Lecture\pdf2mp4_size_spec.py"
  SetOutPath "$APPDATA\Microsoft\AddIns\Perfect_Lecture\TTS_engine"
  File "Install\AppData\Roaming\Microsoft\AddIns\Perfect_Lecture\TTS_engine\TTS_engine.exe"
  File "Install\AppData\Roaming\Microsoft\AddIns\Perfect_Lecture\TTS_engine\TTS_engine.exe.config"
  File "Install\AppData\Roaming\Microsoft\AddIns\Perfect_Lecture\TTS_engine\TTS_engine.pdb"
  SetOutPath "$APPDATA\Microsoft\AddIns"
  File "Install\AppData\Roaming\Microsoft\AddIns\Perfect_Lecture.dotm"
  File "Install\AppData\Roaming\Microsoft\AddIns\Perfect_Lecture.ppam"
  File "Install\AppData\Roaming\Microsoft\AddIns\Perfect_Lecture.xlam"
  SetOutPath "$APPDATA\Microsoft\AddIns\TeX4Office_Editor\app.publish"
  File "Install\AppData\Roaming\Microsoft\AddIns\TeX4Office_Editor\app.publish\TeX4Office_Editor.exe"
  SetOutPath "$APPDATA\Microsoft\AddIns\TeX4Office_Editor"
  File "Install\AppData\Roaming\Microsoft\AddIns\TeX4Office_Editor\default.png"
  SetOutPath "$APPDATA\Microsoft\AddIns\TeX4Office_Editor\ImageMagick-portable"
  File "Install\AppData\Roaming\Microsoft\AddIns\TeX4Office_Editor\ImageMagick-portable\convert.exe"
  SetOutPath "$APPDATA\Microsoft\AddIns\TeX4Office_Editor"
  File "Install\AppData\Roaming\Microsoft\AddIns\TeX4Office_Editor\System.Windows.Forms.Ribbon35.dll"
  SetOutPath "$APPDATA\Microsoft\AddIns\TeX4Office_Editor\templates"
  File "Install\AppData\Roaming\Microsoft\AddIns\TeX4Office_Editor\templates\LuaLaTeX.Math.Chinese_Equations.tex"
  File "Install\AppData\Roaming\Microsoft\AddIns\TeX4Office_Editor\templates\PDFLaTeX.Math.Chinese_Equations.tex"
  File "Install\AppData\Roaming\Microsoft\AddIns\TeX4Office_Editor\templates\PDFLaTeX.Music.tex"
  File "Install\AppData\Roaming\Microsoft\AddIns\TeX4Office_Editor\templates\XeLaTeX.Chinese_fake_bold_slant.tex"
  File "Install\AppData\Roaming\Microsoft\AddIns\TeX4Office_Editor\templates\XeLaTeX.Math.Chinese_Equations.tex"
  SetOutPath "$APPDATA\Microsoft\AddIns\TeX4Office_Editor"
  File "Install\AppData\Roaming\Microsoft\AddIns\TeX4Office_Editor\TeX4Office_Editor.application"
  File "Install\AppData\Roaming\Microsoft\AddIns\TeX4Office_Editor\TeX4Office_Editor.exe"
  File "Install\AppData\Roaming\Microsoft\AddIns\TeX4Office_Editor\TeX4Office_Editor.exe.config"
  File "Install\AppData\Roaming\Microsoft\AddIns\TeX4Office_Editor\TeX4Office_Editor.exe.manifest"
  File "Install\AppData\Roaming\Microsoft\AddIns\TeX4Office_Editor\TeX4Office_Editor.pdb"
  File "Install\AppData\Roaming\Microsoft\AddIns\TeX4Office_Editor\使用手冊.pdf"
  SetOutPath "$APPDATA\Microsoft\Excel\xlstart"
  File "Install\AppData\Roaming\Microsoft\Excel\xlstart\Perfect_Lecture.xlam"
  File "Install\AppData\Roaming\Microsoft\Excel\xlstart\TeX4Office_Beta1.xlam"
  SetOutPath "$APPDATA\Microsoft\Word\startup"
  File "Install\AppData\Roaming\Microsoft\Word\startup\Perfect_Lecture.dotm"
  File "Install\AppData\Roaming\Microsoft\Word\startup\TeX4Office_Beta1.dotm"
  SetOutPath "$INSTDIR"
  File "Document\FAQ.txt"
  File "Document\使用手冊 - Perfect Lecture.docx"
  File "Document\安裝手冊 - Perfect Lecture.docx"
  File "Document\開發者手冊 - Perfect Lecture.docx"
SectionEnd

Section -AdditionalIcons
  WriteIniStr "$INSTDIR\${PRODUCT_NAME}.url" "InternetShortcut" "URL" "${PRODUCT_WEB_SITE}"
  CreateDirectory "$SMPROGRAMS\Perfect Lecture"
  CreateShortCut "$SMPROGRAMS\Perfect Lecture\Website.lnk" "$INSTDIR\${PRODUCT_NAME}.url"
  CreateShortCut "$SMPROGRAMS\Perfect Lecture\Uninstall.lnk" "$INSTDIR\uninst.exe"
  CreateShortCut "$SMPROGRAMS\Perfect Lecture\使用手冊.lnk" "$INSTDIR\使用手冊 - Perfect Lecture.docx"
  CreateShortCut "$SMPROGRAMS\Perfect Lecture\安裝手冊.lnk" "$INSTDIR\安裝手冊 - Perfect Lecture.docx"
  CreateShortCut "$SMPROGRAMS\Perfect Lecture\開發者手冊.lnk" "$INSTDIR\開發者手冊 - Perfect Lecture.docx"
SectionEnd

Section -Post
  WriteUninstaller "$INSTDIR\uninst.exe"
  WriteRegStr HKLM "${PRODUCT_DIR_REGKEY}" "" "$INSTDIR"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayName" "$(^Name)"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "UninstallString" "$INSTDIR\uninst.exe"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayIcon" "$APPDATA\Microsoft\AddIns\Perfect_Lecture\ImageMagick-portable\convert.exe"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayVersion" "${PRODUCT_VERSION}"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "URLInfoAbout" "${PRODUCT_WEB_SITE}"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "Publisher" "${PRODUCT_PUBLISHER}"
SectionEnd


Function un.onUninstSuccess
  HideWindow
  MessageBox MB_ICONINFORMATION|MB_OK "$(^Name) 已成功地從你的電腦移除。"
FunctionEnd

Function un.onInit
  MessageBox MB_ICONQUESTION|MB_YESNO|MB_DEFBUTTON2 "你確定要完全移除 $(^Name) ，其及所有的元件？" IDYES +2
  Abort
FunctionEnd

Section Uninstall
  Delete "$INSTDIR\${PRODUCT_NAME}.url"
  Delete "$INSTDIR\uninst.exe"
  Delete "$INSTDIR\開發者手冊 - Perfect Lecture.docx"
  Delete "$INSTDIR\安裝手冊 - Perfect Lecture.docx"
  Delete "$INSTDIR\使用手冊 - Perfect Lecture.docx"
  Delete "$INSTDIR\FAQ.txt"
  Delete "$APPDATA\Microsoft\Word\startup\TeX4Office_Beta1.dotm"
  Delete "$APPDATA\Microsoft\Word\startup\Perfect_Lecture.dotm"
  Delete "$APPDATA\Microsoft\Excel\xlstart\TeX4Office_Beta1.xlam"
  Delete "$APPDATA\Microsoft\Excel\xlstart\Perfect_Lecture.xlam"
  Delete "$APPDATA\Microsoft\AddIns\TeX4Office_Editor\使用手冊.pdf"
  Delete "$APPDATA\Microsoft\AddIns\TeX4Office_Editor\TeX4Office_Editor.pdb"
  Delete "$APPDATA\Microsoft\AddIns\TeX4Office_Editor\TeX4Office_Editor.exe.manifest"
  Delete "$APPDATA\Microsoft\AddIns\TeX4Office_Editor\TeX4Office_Editor.exe.config"
  Delete "$APPDATA\Microsoft\AddIns\TeX4Office_Editor\TeX4Office_Editor.exe"
  Delete "$APPDATA\Microsoft\AddIns\TeX4Office_Editor\TeX4Office_Editor.application"
  Delete "$APPDATA\Microsoft\AddIns\TeX4Office_Editor\templates\XeLaTeX.Math.Chinese_Equations.tex"
  Delete "$APPDATA\Microsoft\AddIns\TeX4Office_Editor\templates\XeLaTeX.Chinese_fake_bold_slant.tex"
  Delete "$APPDATA\Microsoft\AddIns\TeX4Office_Editor\templates\PDFLaTeX.Music.tex"
  Delete "$APPDATA\Microsoft\AddIns\TeX4Office_Editor\templates\PDFLaTeX.Math.Chinese_Equations.tex"
  Delete "$APPDATA\Microsoft\AddIns\TeX4Office_Editor\templates\LuaLaTeX.Math.Chinese_Equations.tex"
  Delete "$APPDATA\Microsoft\AddIns\TeX4Office_Editor\System.Windows.Forms.Ribbon35.dll"
  Delete "$APPDATA\Microsoft\AddIns\TeX4Office_Editor\ImageMagick-portable\convert.exe"
  Delete "$APPDATA\Microsoft\AddIns\TeX4Office_Editor\default.png"
  Delete "$APPDATA\Microsoft\AddIns\TeX4Office_Editor\app.publish\TeX4Office_Editor.exe"
  Delete "$APPDATA\Microsoft\AddIns\Perfect_Lecture.xlam"
  Delete "$APPDATA\Microsoft\AddIns\Perfect_Lecture.ppam"
  Delete "$APPDATA\Microsoft\AddIns\Perfect_Lecture.dotm"
  Delete "$APPDATA\Microsoft\AddIns\Perfect_Lecture\TTS_engine\TTS_engine.pdb"
  Delete "$APPDATA\Microsoft\AddIns\Perfect_Lecture\TTS_engine\TTS_engine.exe.config"
  Delete "$APPDATA\Microsoft\AddIns\Perfect_Lecture\TTS_engine\TTS_engine.exe"
  Delete "$APPDATA\Microsoft\AddIns\Perfect_Lecture\pdf2mp4_size_spec.py"
  Delete "$APPDATA\Microsoft\AddIns\Perfect_Lecture\pdf2mp4.py"
  Delete "$APPDATA\Microsoft\AddIns\Perfect_Lecture\ImageMagick-portable\type.xml"
  Delete "$APPDATA\Microsoft\AddIns\Perfect_Lecture\ImageMagick-portable\type-ghostscript.xml"
  Delete "$APPDATA\Microsoft\AddIns\Perfect_Lecture\ImageMagick-portable\thresholds.xml"
  Delete "$APPDATA\Microsoft\AddIns\Perfect_Lecture\ImageMagick-portable\sRGB.icc"
  Delete "$APPDATA\Microsoft\AddIns\Perfect_Lecture\ImageMagick-portable\quantization-table.xml"
  Delete "$APPDATA\Microsoft\AddIns\Perfect_Lecture\ImageMagick-portable\policy.xml"
  Delete "$APPDATA\Microsoft\AddIns\Perfect_Lecture\ImageMagick-portable\mime.xml"
  Delete "$APPDATA\Microsoft\AddIns\Perfect_Lecture\ImageMagick-portable\magic.xml"
  Delete "$APPDATA\Microsoft\AddIns\Perfect_Lecture\ImageMagick-portable\log.xml"
  Delete "$APPDATA\Microsoft\AddIns\Perfect_Lecture\ImageMagick-portable\locale.xml"
  Delete "$APPDATA\Microsoft\AddIns\Perfect_Lecture\ImageMagick-portable\ImageMagick.rdf"
  Delete "$APPDATA\Microsoft\AddIns\Perfect_Lecture\ImageMagick-portable\ffmpeg.exe"
  Delete "$APPDATA\Microsoft\AddIns\Perfect_Lecture\ImageMagick-portable\english.xml"
  Delete "$APPDATA\Microsoft\AddIns\Perfect_Lecture\ImageMagick-portable\delegates.xml"
  Delete "$APPDATA\Microsoft\AddIns\Perfect_Lecture\ImageMagick-portable\convert.exe"
  Delete "$APPDATA\Microsoft\AddIns\Perfect_Lecture\ImageMagick-portable\configure.xml"
  Delete "$APPDATA\Microsoft\AddIns\Perfect_Lecture\ImageMagick-portable\colors.xml"
  Delete "$APPDATA\Microsoft\AddIns\Perfect_Lecture\genPostProcessVBA.py"

  Delete "$SMPROGRAMS\Perfect Lecture\Uninstall.lnk"
  Delete "$SMPROGRAMS\Perfect Lecture\Website.lnk"
  Delete "$SMPROGRAMS\Perfect Lecture\使用手冊.lnk"
  Delete "$SMPROGRAMS\Perfect Lecture\安裝手冊.lnk"
  Delete "$SMPROGRAMS\Perfect Lecture\開發者手冊.lnk"

  RMDir "$SMPROGRAMS\Perfect Lecture"
  RMDir "$INSTDIR"
  RMDir "$APPDATA\Microsoft\AddIns\TeX4Office_Editor\templates"
  RMDir "$APPDATA\Microsoft\AddIns\TeX4Office_Editor\ImageMagick-portable"
  RMDir "$APPDATA\Microsoft\AddIns\TeX4Office_Editor\app.publish"
  RMDir "$APPDATA\Microsoft\AddIns\TeX4Office_Editor"
  RMDir "$APPDATA\Microsoft\AddIns\Perfect_Lecture\TTS_engine"
  RMDir "$APPDATA\Microsoft\AddIns\Perfect_Lecture\ImageMagick-portable"
  RMDir "$APPDATA\Microsoft\AddIns\Perfect_Lecture"

  DeleteRegKey ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}"
  DeleteRegKey HKLM "${PRODUCT_DIR_REGKEY}"
  SetAutoClose true
SectionEnd