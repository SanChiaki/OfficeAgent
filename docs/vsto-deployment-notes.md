# OfficeAgent VSTO Deployment Notes

## Packaging outputs

- `installer/OfficeAgent.Setup/build.ps1` builds two MSI packages:
  - `artifacts/installer/xISDP.Setup-x86.msi`
  - `artifacts/installer/xISDP.Setup-x64.msi`
- Deploy the MSI that matches the target Excel bitness.

## Runtime prerequisites

- The installer blocks when the `VSTO Runtime` prerequisite is missing.
- The installer blocks when the `Microsoft Edge WebView2 Runtime` prerequisite is missing.
- The current MVP does not bundle the WebView2 runtime installer into the MSI.
- For online environments, install the Evergreen WebView2 Runtime before the MSI.
- For intranet or offline environments, distribute your approved offline WebView2 runtime package before the MSI.

## Trust and signing

- The current MVP build signs VSTO manifests with `OfficeAgent Dev Certificate`.
- That certificate is suitable for local development and controlled internal testing only.
- Before broad distribution, replace it with a trusted enterprise or public code-signing certificate and rebuild the add-in.
- If you keep the development certificate temporarily, your deployment process must import the publisher certificate into trusted stores on target machines before first launch.

## End-user experience target

- No Office Add-in manifest sideload is required.
- No shared catalog or localhost certificate trust is required.
- After prerequisites are present and the package is installed, Excel should load the add-in from the local VSTO registration on startup.
