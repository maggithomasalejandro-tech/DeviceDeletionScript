# Device Deletion Script

Java utility for safely bulk-deleting devices from **Resource Panorama** through its **GraphQL API**.

The script reads MongoDB device IDs from an Excel file, validates each device before deletion, processes requests in controlled batches, retries failed operations, and generates a CSV report with the result of every deletion attempt.

> ⚠️ **Warning**
>
> This tool can perform irreversible deletions against production.
> Always validate the input file and test against the development environment first.

---

## Features

- 📄 Reads MongoDB device IDs from an Excel (`.xlsx`) file
- 🔎 Verifies that each device exists before attempting deletion
- 🗑️ Deletes devices through the Resource Panorama GraphQL API
- ⚡ Processes devices concurrently in controlled batches
- ⏱️ Adds a configurable delay between batches to avoid overloading the API
- 🔄 Automatically retries failed requests
- 🛡️ Continues processing if an individual device fails
- 📊 Generates a CSV report containing the result of every processed ID
- 🔐 Keeps authentication credentials outside the repository

---

## How it works

For every device ID, the script performs the following flow:

```text
Excel (.xlsx)
     │
     ▼
Read device ID
     │
     ▼
Verify device exists
Device(where: {_id})
     │
     ├── Not found ──────► Record result
     │
     ▼
Delete device
deleteDevice(_id, hardDelete)
     │
     ├── Failure ────────► Retry
     │
     ▼
Record result
     │
     ▼
device_deletion_results.csv
```

Devices are processed in batches of **3 concurrent requests**, with a **10-second delay** between batches by default.

A failure affecting one device does not stop the remaining devices from being processed.

---

## Requirements

- **Java 21**
- **Maven**
- Access to the Resource Panorama GraphQL API
- Valid Keycloak credentials
- IntelliJ IDEA (recommended)

### Apache POI

Excel files are read using Apache POI:

```xml
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi-ooxml</artifactId>
    <version>5.2.5</version>
</dependency>
```

The dependency is already defined in `pom.xml`.

---

## Configuration

### Environment variables

Authentication credentials must be provided through environment variables.

| Variable | Description |
|---|---|
| `RP_KEYCLOAK_TOKEN_URL` | Full Keycloak token endpoint URL |
| `RP_KEYCLOAK_CLIENT_ID` | Keycloak client ID |
| `RP_KEYCLOAK_USERNAME` | Service account username |
| `RP_KEYCLOAK_PASSWORD` | Service account password |

Credentials must **never be hardcoded or committed to the repository**.

They can be configured through:

- IntelliJ **Run Configuration → Environment variables**
- A local `.env` file using the IntelliJ EnvFile plugin
- System environment variables

### Special characters in `.env`

If a value contains characters such as `#`, `!`, `@` or `$`, wrap it in quotes:

```env
RP_KEYCLOAK_PASSWORD="MyP@ssw0rd#123!"
```

Without quotes, some `.env` parsers may interpret `#` as the beginning of a comment and silently truncate the value.

This can result in authentication errors such as:

```text
invalid_grant
Invalid user credentials
```

---

## Application configuration

Local application configuration is stored in:

```text
src/main/resources/application.yml
```

This file is intentionally excluded from Git because it may contain environment-specific or sensitive configuration.

Make sure your local `application.yml` is correctly configured before running the script.

---

## Script settings

The main processing settings can be adjusted in `DeviceDeletionScript.java`.

| Constant | Default | Description |
|---|---:|---|
| `GRAPHQL_ENDPOINT` | Production endpoint | Resource Panorama GraphQL endpoint |
| `EXCEL_PATH` | `device_ids.xlsx` | Input Excel file |
| `EXCEL_COLUMN_INDEX` | `0` | Column containing device IDs |
| `EXCEL_HEADER_ROWS` | `1` | Number of header rows to skip |
| `BATCH_SIZE` | `3` | Devices processed concurrently |
| `WAIT_BETWEEN_BATCHES_MS` | `10000` | Delay between batches (ms) |
| `MAX_RETRIES` | `1` | Retry attempts after a failure |
| `HARD_DELETE` | `true` | Enables complete associated-data cleanup |

When `HARD_DELETE` is enabled, the deletion may include associated resources such as Keycloak roles, alarms, loggers and dictionaries.

---

## Input file

By default, the script expects:

```text
device_ids.xlsx
```

The Excel file should contain one MongoDB device ID per row.

Example:

| id |
|---|
| `6a758b3f31bc64b96b4154a0` |
| `696e0b52568e2f26a99c7b78` |
| `...` |

The ID column and number of header rows can be changed through:

```java
EXCEL_COLUMN_INDEX
EXCEL_HEADER_ROWS
```

---

## Running the script

1. Configure the required Keycloak environment variables.
2. Configure your local `application.yml`.
3. Place the Excel file containing the device IDs at the location defined by `EXCEL_PATH`.
4. Verify that the GraphQL endpoint points to the intended environment.
5. Run:

```java
DeviceDeletionScript.main()
```

6. Wait for all batches to complete.
7. Review the generated results file:

```text
device_deletion_results.csv
```

---

## Output

The script generates:

```text
device_deletion_results.csv
```

The report contains the result of each processed device, including:

- Device ID
- Verification result
- Deletion result
- Status/error message
- Timestamp

This makes it possible to identify successful deletions, skipped devices and failed requests without relying only on console logs.

---

## Production safety

Before running against production:

1. **Test against the development environment first** using only 2–3 test devices.
2. Verify every ID in the input Excel file.
3. Confirm that the authentication credentials belong to the same environment as the GraphQL endpoint.
4. Confirm whether `hardDelete: true` is appropriate for the operation.
5. Review the configured batch size and delay.
6. Never commit credentials, tokens or environment-specific secrets.

> **Production deletions are irreversible.**

---

## Troubleshooting

| Problem | Possible cause | Solution |
|---|---|---|
| POI `Cannot resolve symbol` | Maven dependency is missing or not loaded | Verify `pom.xml` and run **Maven Reload** |
| `invalid_grant` / `Invalid user credentials` | Incorrect credentials or `.env` value truncated | Verify credentials and quote values containing special characters |
| `Illegal character in authority` | Unresolved Spring-style variable syntax in URL | Use the resolved URL instead of `${VAR:value}` |
| HTTP 500 from GraphQL | Credentials and API endpoint belong to different environments | Verify that Keycloak and GraphQL are using the same environment |
| Device not deleted | Device does not exist or GraphQL mutation failed | Check the generated CSV and application logs |
| GitHub blocks the push | A credential or secret was detected | Remove the secret from Git history and keep sensitive configuration ignored |

---

## Security

The repository must never contain:

```text
Passwords
Client secrets
Access tokens
Refresh tokens
Production credentials
Private environment configuration
```

Sensitive configuration should be provided through environment variables or ignored local configuration files.

The following files should remain excluded from version control:

```gitignore
src/main/resources/application.yml
.env
device_ids.xlsx
device_deletion_results.csv
```

---

## Project structure

```text
DeviceDeletionScript/
│
├── src/
│   └── main/
│       ├── java/
│       │   └── org/example/
│       │       └── DeviceDeletionScript.java
│       │
│       └── resources/
│           └── application.yml      # Local / ignored
│
├── device_ids.xlsx                  # Input / ignored
├── device_deletion_results.csv      # Generated / ignored
├── pom.xml
├── .gitignore
└── README.md
```

---

## Disclaimer

This utility performs destructive operations against Resource Panorama.

Use it only with authorized credentials and after validating the target environment and device list.
