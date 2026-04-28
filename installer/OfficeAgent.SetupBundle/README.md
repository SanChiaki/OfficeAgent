# X-ISDP Setup Bundle

This directory contains the WiX Burn bundle source for the offline installer build.

When integrated into the installer build flow, it produces:

- `artifacts/installer/X-ISDP.Setup.exe`

Required staged prerequisite installers:

- `prereqs/vstor_redist.exe`
- `prereqs/MicrosoftEdgeWebView2RuntimeInstallerX86.exe`
- `prereqs/MicrosoftEdgeWebView2RuntimeInstallerX64.exe`

The MSI inputs (`X-ISDP.Setup-x86.msi` and `X-ISDP.Setup-x64.msi`) are produced and wired by the installer build pipeline.
