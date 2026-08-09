# Security Policy

## Supported Versions

OfficeFileGuard is an actively maintained project.

Security fixes are provided for the latest released version available on GitHub.

- Latest release: supported and receives security updates.
- Older releases: may continue to work, but security updates and bug fixes are not guaranteed. Users are encouraged to upgrade to the latest available release.

---

## Reporting a Vulnerability

If you believe you have discovered a security vulnerability in OfficeFileGuard, please do **not** open a public GitHub issue.

Instead, report the issue privately by contacting the project maintainer.

Please include, whenever possible:

* A description of the vulnerability.
* Steps required to reproduce the issue.
* The affected version.
* Any proof-of-concept or sample files (if applicable).
* Suggestions for a possible fix, if available.

The maintainer will acknowledge the report, investigate the issue, and work toward a fix as quickly as possible.

Once a fix is available, the vulnerability may be publicly disclosed together with the corresponding release notes.

---

## Security Scope

OfficeFileGuard is designed to reduce the risk of accidental disclosure of Microsoft Office documents.

It is **not** intended to replace enterprise security solutions such as:

* Data Loss Prevention (DLP)
* Information Protection
* Email Gateway Security
* Document Classification Systems
* Endpoint Protection Platforms

Instead, OfficeFileGuard provides an additional user confirmation layer before confidential Office documents are sent to recipients outside the sender's organization.

---

## Privacy

Privacy is a fundamental design principle of OfficeFileGuard.

The application is designed to operate entirely on the local computer.

Specifically:

* No documents are uploaded.
* No document contents are transmitted.
* No attachment metadata is sent to external services.
* No cloud services are required.
* No telemetry is collected.
* No analytics are performed.
* No user information is tracked.

All processing is performed locally using standard Microsoft Office and Windows APIs.

---

## Responsible Disclosure

Please allow reasonable time for security issues to be investigated and corrected before publicly disclosing vulnerabilities.
Responsible disclosure helps protect all users while fixes are being prepared and released.
Thank you for helping improve the security of OfficeFileGuard.
