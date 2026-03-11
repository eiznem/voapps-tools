; installer.nsh — Custom NSIS hooks for VoApps Tools
;
; Detects whether VoApps Tools is running before installation begins.
; nsProcess.nsh is already included by allowOnlyOneInstallerInstance.nsh
; (which electron-builder always includes), so no extra !include is needed.
;
; preInit runs inside .onInit, before UAC elevation.
; If VoApps Tools.exe is running the user sees an OK/Cancel dialog:
;   OK     → gracefully closes the app (WM_CLOSE), waits 2 s,
;             force-kills if it still hasn't exited, then continues
;   Cancel → aborts the installer immediately

!macro preInit
  ${nsProcess::FindProcess} "VoApps Tools.exe" $R0
  ${If} $R0 == 0
    ; App is running — prompt the user
    MessageBox MB_OKCANCEL|MB_ICONEXCLAMATION \
      "VoApps Tools is currently running.$\n$\n\
Click OK to close it automatically and continue the installation,$\n\
or click Cancel to abort." \
      IDOK voapps_close_and_continue

    ; ── Cancel: unload plugin and quit the installer ─────────────────────────
    ${nsProcess::Unload}
    Quit

    ; ── OK: close the app gracefully, force-kill if needed ───────────────────
    voapps_close_and_continue:
      ${nsProcess::CloseProcess} "VoApps Tools.exe" $R0
      Sleep 2000
      ; If still alive after 2 s, terminate it
      ${nsProcess::FindProcess} "VoApps Tools.exe" $R0
      ${If} $R0 == 0
        ${nsProcess::KillProcess} "VoApps Tools.exe" $R0
        Sleep 1000
      ${EndIf}
  ${EndIf}
  ${nsProcess::Unload}
!macroend
