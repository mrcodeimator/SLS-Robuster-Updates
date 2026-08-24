# SLS Robuster — Downloads & Updates

This is the **public** home for SLS Robuster Mac installs and in-app updates.

You do **not** need a GitHub login to download the app from Releases.

## Download the Mac app

1. Open **[Releases](https://github.com/mrcodeimator/SLS-Robuster-Updates/releases)**
2. Download the newest `.zip` (for example `SLS_ROBUSTER-macos-arm64.zip`)
3. Unzip → drag **`SLS_ROBUSTER.app`** to **Applications**
4. First open: **Right-click** the app → **Open** → **Open** again

## Handoff / how-to guide (PDF)

Full non-technical documentation (install, setup, how it works, where settings are saved):

- In this repo: [SLS_Robuster_Handoff_Guide.pdf](./SLS_Robuster_Handoff_Guide.pdf)
- Also attached to the [v2.2.0 Release](https://github.com/mrcodeimator/SLS-Robuster-Updates/releases/tag/v2.2.0)

## Where settings are saved on a Mac

```text
~/Library/Application Support/SLSRobuster/
```

Replacing the app does **not** wipe YouTube logins or settings.

## Check for Updates (after 2.2.0)

Inside the app: **Settings → General → Check for Updates**.  
The app reads `version.json` in this repo and downloads the zip from **Releases** here.

## Source code (private)

Programming / build source lives in the private repo:  
[mrcodeimator/SLS-ROBUSTER-V2](https://github.com/mrcodeimator/SLS-ROBUSTER-V2)

A merge there is **not** a public release by itself — the zip and this repo’s `version.json` must be updated for Macs to see a new build.

## Quick links

| Need | Link |
|------|------|
| Download app | [Releases](https://github.com/mrcodeimator/SLS-Robuster-Updates/releases) |
| Handoff PDF | [SLS_Robuster_Handoff_Guide.pdf](./SLS_Robuster_Handoff_Guide.pdf) |
| Current version metadata | [version.json](./version.json) |
| Private source | [SLS-ROBUSTER-V2](https://github.com/mrcodeimator/SLS-ROBUSTER-V2) |
