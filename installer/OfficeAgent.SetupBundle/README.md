# xISDP Setup Bundle

This directory contains the WiX Burn bundle source for the offline installer build.

When integrated into the installer build flow, it produces:

- `artifacts/installer/xISDP.Setup.exe`

Required staged prerequisite installers:

- `prereqs/vstor_redist.exe`

The MSI inputs (`xISDP.Setup-x86.msi` and `xISDP.Setup-x64.msi`) are produced and wired by the installer build pipeline.

To enable release update reminders, pass the private manifest endpoint with
`installer/OfficeAgent.Setup/build.ps1 -UpdateManifestUrl "<url>"`. The MSI writes
that value to `HKCU\Software\OfficeAgent\UpdateManifestUrl`; leaving the argument
empty keeps update checks disabled.
