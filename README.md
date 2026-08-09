# OfficeFileGuard OWA Add-in

**OWA_Addin** is the Outlook add-in component of [OfficeFileGuard](https://github.com/OfficeFileGuard).

It extends OfficeFileGuard protection to:

* **New Outlook for Windows**
* **Outlook on the web**

The add-in is based on the Microsoft **Office.js** platform.

> **Experimental:** OWA_Addin is currently experimental and under active development. Its behavior, installation process, supported Outlook environments, and implementation may change in future releases.

---

## How It Fits Into OfficeFileGuard

OfficeFileGuard is composed of two complementary components:

| Repository                                                                            | Purpose                                                                                       |
| ------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------- |
| **[OfficeFileOutlook_WIN](https://github.com/OfficeFileGuard/OfficeFileOutlook_WIN)** | Main Windows component, including Windows Explorer integration and Classic Outlook protection |
| **[OWA_Addin](https://github.com/OfficeFileGuard/OWA_Addin)**                         | Outlook add-in for New Outlook and Outlook on the web                                         |

The two components address different Outlook environments:

* **Classic Outlook for Windows** → protected by the VSTO add-in included in `OfficeFileOutlook_WIN`
* **New Outlook for Windows** → protected by `OWA_Addin`
* **Outlook on the web (OWA)** → protected by `OWA_Addin`

---

## Do I Need OWA_Addin?

**Only if you use New Outlook or Outlook on the web.**

If you use **Classic Outlook for Windows**, you do **not** need to install OWA_Addin. Classic Outlook protection is already provided by the Outlook VSTO add-in included in **OfficeFileOutlook_WIN**.

| Your Outlook                | What you need                         |
| --------------------------- | ------------------------------------- |
| Classic Outlook for Windows | **OfficeFileOutlook_WIN**             |
| New Outlook for Windows     | **OfficeFileOutlook_WIN + OWA_Addin** |
| Outlook on the web (OWA)    | **OfficeFileOutlook_WIN + OWA_Addin** |

---

## Dependency on OfficeFileOutlook_WIN

**OWA_Addin is not a standalone component.**

**OfficeFileOutlook_WIN must be installed** on the user's Windows computer.

For the main OfficeFileGuard Windows component, see:

**[OfficeFileOutlook_WIN](https://github.com/OfficeFileGuard/OfficeFileOutlook_WIN)**

---

## Installation

Before installing OWA_Addin, make sure **OfficeFileOutlook_WIN** is installed.

1. In Outlook, open **Add-Ins for Outlook → My add-ins**.
   (If you cannot find this window, open the following link with a web browser: **https://aka.ms/olksideload**)

2. Under **Custom Addins**, select **Add a custom add-in → Add from File**.

3. Download the OfficeFileGuard manifest **https://officefileguard.com/download/manifest-outlook-com.xlm**

4. Select the downloaded manifest file and confirm.

---

## Supported Outlook Environments

OWA_Addin is intended for:

* New Outlook for Windows
* Outlook on the web (OWA)

It does **not** replace the VSTO add-in used by Classic Outlook.

The overall architecture is:

```text
                         OfficeFileGuard
                              │
              ┌───────────────┴───────────────┐
              │                               │
   OfficeFileOutlook_WIN                  OWA_Addin
              │                               │
       ┌──────┴──────┐                Office.js Add-in
       │             │                       │
 Windows Explorer  Classic Outlook      New Outlook
                                           +
                                      Outlook on the web
```

---

## How OWA_Addin Works

OWA_Addin uses the Microsoft **Office.js** add-in platform to integrate OfficeFileGuard protection into New Outlook and Outlook on the web.

The add-in participates in the email sending workflow and checks the message context required by OfficeFileGuard to determine whether protection should be applied.

The OWA implementation is separate from the .NET/VSTO implementation used by `OfficeFileOutlook_WIN` because New Outlook and Outlook on the web use a different add-in architecture.

---

## Current Status

**OWA_Addin is experimental.**

The add-in is being developed to extend the OfficeFileGuard protection model to Microsoft's newer Outlook platforms.

Features, compatibility, installation procedures, and implementation details may change as development progresses.

Users who require the most established OfficeFileGuard experience should use **OfficeFileOutlook_WIN** with Classic Outlook.

---

## Development

OWA_Addin is based on the Microsoft **Office.js** add-in framework.

The project has its own development environment and build process and is maintained independently from the .NET solution contained in `OfficeFileOutlook_WIN`.

Development and installation instructions will be documented here as the project matures.

---

## Privacy

OfficeFileGuard follows a privacy-first approach.

Document inspection is performed locally and OfficeFileGuard does not upload documents or their contents to external servers.

For the complete privacy model, see the documentation in the main **[OfficeFileOutlook_WIN](https://github.com/OfficeFileGuard/OfficeFileOutlook_WIN)** repository.

---

## Security

OfficeFileGuard is intended to reduce the risk of accidental information disclosure.

It is **not** a replacement for enterprise Data Loss Prevention (DLP), Microsoft Information Protection, or document classification systems.

For security reporting, see [SECURITY.md](https://github.com/OfficeFileGuard/OWA_Addin/blob/main/SECURITY.md).

For contribution guidelines, see [CONTRIBUTING.md](https://github.com/OfficeFileGuard/OWA_Addin/blob/main/CONTRIBUTING.md).

---

## License

OfficeFileGuard is open source and licensed under the **Apache License 2.0**.

See the [LICENSE](https://github.com/OfficeFileGuard/OWA_Addin/blob/main/LICENSE) file for details.
