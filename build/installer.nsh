; installer.nsh — Custom NSIS hooks for VoApps Tools
;
; Detects whether VoApps Tools is running before installation begins.
; If it is, prompts the user with an OK/Cancel dialog:
;   OK     → gracefully closes the app (force-kills if it doesn't exit in 2 s)
;   Cancel → aborts the installer immediately

; ── include nsProcess plugin (bundled with electron-builder) ─────────────────
!macro customHeader
  !include "nsProcess.nsh"
!macroend

; ── preInit runs inside .onInit, before UAC elevation ───────────────────────
!macro preInit
  ${nsProcess::FindProcess} "VoApps Tools.exe" $R0
  ${If} $R0 == 0
    ; App is running — ask the user what to do
    MessageBox MB_OKCANCEL|MB_ICONEXCLAMATION \
      "VoApps Tools is currently running.$\n$\n\
Click OK to close it automatically and continue the installation,$\n\
or click Cancel to abort." \
      IDOK voapps_close_and_continue

    ; ── User clicked Cancel ──────────────────────────────────────────────────
    ${nsProcess::Unload}
    Quit

    ; ── User clicked OK — close the app ─────────────────────────────────────
    voapps_close_and_continue:
      ; Send WM_CLOSE for a graceful shutdown
      ${nsProcess::CloseProcess} "VoApps Tools.exe" $R0
      Sleep 2000

      ; If it's still alive after 2 s, force-terminate it
      ${nsProcess::FindProcess} "VoApps Tools.exe" $R0
      ${If} $R0 == 0
        ${nsProcess::KillProcess} "VoApps Tools.exe" $R0
        Sleep 1000
      ${EndIf}
  ${EndIf}
  ${nsProcess::Unload}
!macroend
