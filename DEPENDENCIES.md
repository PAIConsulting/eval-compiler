# Dependencies

This document lists all third-party dependencies of the Evaluation Compiler.
It serves as a human-readable Software Bill of Materials (SBOM) summary.

For machine-readable SBOM generation, see the bottom of this file.

## Direct Dependencies

These are the libraries directly imported by the Evaluation Compiler code:

| Name | Version | License | Purpose |
|------|---------|---------|---------|
| python-docx | 1.2.0 | MIT | Generate Word (.docx) documents |
| PyYAML | 6.0.3 | MIT | Parse the keyword dictionary YAML file |
| matplotlib | 3.10.8 | Python Software Foundation License (matplotlib license) | Generate charts embedded in reports |

All three are widely used, well-maintained open-source Python libraries.

### Verification

Verify installed versions match the pinned versions in `requirements.txt`:

```bash
pip show python-docx PyYAML matplotlib
```

## Standard Library Modules Used

These are part of Python itself and have no separate license or distribution:

- `csv` — Read/write CSV files
- `os` — File system operations (creating output directory)
- `random` — Generate synthetic test data (test data generator only)
- `collections` — Counter and defaultdict for tallying selections
- `io.BytesIO` — In-memory buffer for chart image data

## What's NOT in the dependency tree

These categories of libraries are **deliberately absent** from the codebase:

- **AI / ML libraries.** No openai, anthropic, transformers, torch,
  tensorflow, langchain, scikit-learn, etc.
- **Networking libraries.** No requests, urllib (other than as a
  transitive dependency of matplotlib for cache lookups, which is not
  invoked in our code paths), httpx, socket, http.client.
- **Cloud SDKs.** No boto3, google-cloud-*, azure-*, etc.
- **Database connectors.** No SQL or NoSQL database drivers.
- **Web frameworks.** No flask, django, fastapi, etc.

## Transitive Dependencies

The three direct dependencies bring in their own dependencies (transitive
dependencies). The most notable is matplotlib, which depends on numpy, pillow,
and a few others. None of these are AI-related or networking-related at runtime.

To see the full transitive dependency tree:

```bash
pip install pipdeptree
pipdeptree -p python-docx,PyYAML,matplotlib
```

## Generating a Machine-Readable SBOM

If your audit process requires a CycloneDX-format SBOM file:

```bash
pip install cyclonedx-bom
cyclonedx-py requirements -i requirements.txt -o sbom.json
```

Or for a license-focused report:

```bash
pip install pip-licenses
pip-licenses --format=markdown --with-urls > LICENSES_REPORT.md
```

## Updating Dependencies

If you update any version in `requirements.txt`:

1. Test the application with the new version
2. Re-run any security audits (CodeQL, SBOM checks)
3. Note the change in your project log

Dependency updates should generally happen during planned maintenance, not
in the middle of a production reporting cycle.
