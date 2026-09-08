# DeviceDeletionScript

Script Java para borrar en bloque devices de Resource Panorama vía la API GraphQL.
Java script to bulk-delete devices from Resource Panorama via the GraphQL API.

---

## 🇪🇸 Español

### Qué hace

- Lee una lista de ids de MongoDB (uno por fila) desde un Excel.
- Por cada id: primero **verifica** que el device existe (`Device(where: {_id})`), y si existe, lo **borra** (`deleteDevice(_id, hardDelete)`).
- Procesa en **lotes de 3** en paralelo, esperando **10 segundos** entre lotes, para no saturar la API.
- Si una llamada falla, **reintenta una vez**, deja constancia en el log y **continúa** con el resto (no aborta todo el proceso).
- Al terminar, genera un **CSV de resultados** (`device_deletion_results.csv`) con el detalle de cada id: si se verificó, si se borró, mensaje y timestamp.

### Requisitos

- Java 21, Maven.
- Dependencia Apache POI (lectura de Excel), añadida en `pom.xml`:
  ```xml
  <dependency>
      <groupId>org.apache.poi</groupId>
      <artifactId>poi-ooxml</artifactId>
      <version>5.2.5</version>
  </dependency>
  ```

### Configuración

El script lee las credenciales de Keycloak desde **variables de entorno** (nunca hardcodeadas en el código):

| Variable | Descripción |
|---|---|
| `RP_KEYCLOAK_TOKEN_URL` | URL completa del endpoint de token de Keycloak |
| `RP_KEYCLOAK_CLIENT_ID` | Client id |
| `RP_KEYCLOAK_USERNAME` | Usuario de servicio |
| `RP_KEYCLOAK_PASSWORD` | Contraseña del usuario de servicio |

Se pueden definir en un archivo `.env` en la raíz del proyecto (usando el plugin EnvFile de IntelliJ en la Run Configuration) o directamente en el campo **Environment variables** de la Run Configuration.

⚠️ **Importante:** si la contraseña (o cualquier valor) contiene caracteres especiales como `#`, `!`, `@`, `$`, ponla **entre comillas** en el `.env`:
```
RP_KEYCLOAK_PASSWORD="MiP@ssw0rd#123!"
```
Sin comillas, un `#` puede hacer que el parser trate el resto de la línea como un comentario y trunque el valor silenciosamente — este fue precisamente el bug que causó errores `invalid_grant` durante el desarrollo del script.

Otras constantes configurables directamente en el código (`DeviceDeletionScript.java`):

| Constante | Por defecto | Descripción |
|---|---|---|
| `GRAPHQL_ENDPOINT` | prod: `https://resourcepanorama-api.greenpowermonitor.com/graphql` | Endpoint de la API. Cambiar a `-dev` para probar en desarrollo |
| `EXCEL_PATH` | `device_ids.xlsx` | Ruta del Excel con los ids |
| `EXCEL_COLUMN_INDEX` | `0` (columna A) | Columna donde están los ids |
| `EXCEL_HEADER_ROWS` | `1` | Filas de cabecera a saltar |
| `BATCH_SIZE` | `3` | Ids procesados en paralelo por lote |
| `WAIT_BETWEEN_BATCHES_MS` | `10000` (10s) | Espera entre lotes |
| `MAX_RETRIES` | `1` | Reintentos por id si falla |
| `HARD_DELETE` | `true` | Si se pasa `hardDelete: true` a la mutation (limpieza completa: rol Keycloak, alarmas, loggers, diccionarios) |

### Formato del Excel de entrada

Una columna con un id de MongoDB por fila (con o sin fila de cabecera, según `EXCEL_HEADER_ROWS`):

| id |
|---|
| 6a758b3f31bc64b96b4154a0 |
| 696e0b52568e2f26a99c7b78 |
| ... |

### Ejecución

1. Configurar las variables de entorno de Keycloak.
2. Colocar el Excel con los ids en la ruta indicada por `EXCEL_PATH`.
3. Confirmar el `GRAPHQL_ENDPOINT` correcto (dev o prod).
4. Ejecutar `DeviceDeletionScript.main()`.
5. Revisar `device_deletion_results.csv` al finalizar.

### ⚠️ Antes de correr contra producción

- **Probar siempre primero en el entorno de dev** con 2-3 ids de prueba.
- Confirmar con el equipo si `hardDelete: true` es necesario para el caso de uso (limpieza completa de datos asociados).
- El borrado en producción es **irreversible**. Verificar la lista de ids antes de lanzar el proceso completo.
- No compartir ni comitear nunca las credenciales de Keycloak (usar `.env` + `.gitignore`, o variables de entorno del sistema).

### Solución de problemas conocidos

| Síntoma | Causa | Solución |
|---|---|---|
| `Cannot resolve symbol` de POI | Falta la dependencia o no se recargó Maven | Añadir dependencia y hacer *Maven Reload* |
| `invalid_grant / Invalid user credentials` | Caracteres especiales en la contraseña truncados por el `.env`, o valores incorrectos | Poner el valor entre comillas en el `.env` |
| `Illegal character in authority` en la URL | La URL de Keycloak se copió con la sintaxis de plantilla de Spring (`${VAR:valor}`) sin resolver | Usar la URL ya resuelta, sin `${}` |
| HTTP 500 al llamar al GraphQL | Token de un entorno (ej. prod) usado contra el endpoint de otro entorno (ej. dev) | Verificar que `GRAPHQL_ENDPOINT` y las credenciales de Keycloak sean del mismo entorno |

---

## 🇬🇧 English

### What it does

- Reads a list of MongoDB ids (one per row) from an Excel file.
- For each id: first **verifies** the device exists (`Device(where: {_id})`), then **deletes** it (`deleteDevice(_id, hardDelete)`) if it does.
- Processes ids in **batches of 3** in parallel, waiting **10 seconds** between batches to avoid overloading the API.
- On failure, **retries once**, logs the outcome, and **continues** with the rest (does not abort the whole run).
- Writes a **results CSV** (`device_deletion_results.csv`) with per-id detail: verified, deleted, message, timestamp.

### Requirements

- Java 21, Maven.
- Apache POI dependency (Excel reading), added to `pom.xml`:
  ```xml
  <dependency>
      <groupId>org.apache.poi</groupId>
      <artifactId>poi-ooxml</artifactId>
      <version>5.2.5</version>
  </dependency>
  ```

### Configuration

Keycloak credentials are read from **environment variables** (never hardcoded):

| Variable | Description |
|---|---|
| `RP_KEYCLOAK_TOKEN_URL` | Full Keycloak token endpoint URL |
| `RP_KEYCLOAK_CLIENT_ID` | Client id |
| `RP_KEYCLOAK_USERNAME` | Service account username |
| `RP_KEYCLOAK_PASSWORD` | Service account password |

These can be set via a `.env` file at the project root (using IntelliJ's EnvFile plugin in the Run Configuration) or directly in the Run Configuration's **Environment variables** field.

⚠️ **Important:** if the password (or any value) contains special characters like `#`, `!`, `@`, `$`, wrap it **in quotes** in the `.env` file:
```
RP_KEYCLOAK_PASSWORD="MyP@ssw0rd#123!"
```
Without quotes, a `#` can cause the parser to treat the rest of the line as a comment and silently truncate the value — this was exactly the bug behind the `invalid_grant` errors hit during development.

Other configurable constants (in `DeviceDeletionScript.java`):

| Constant | Default | Description |
|---|---|---|
| `GRAPHQL_ENDPOINT` | prod: `https://resourcepanorama-api.greenpowermonitor.com/graphql` | API endpoint. Switch to `-dev` for testing |
| `EXCEL_PATH` | `device_ids.xlsx` | Path to the Excel file with ids |
| `EXCEL_COLUMN_INDEX` | `0` (column A) | Column containing the ids |
| `EXCEL_HEADER_ROWS` | `1` | Header rows to skip |
| `BATCH_SIZE` | `3` | Ids processed in parallel per batch |
| `WAIT_BETWEEN_BATCHES_MS` | `10000` (10s) | Wait time between batches |
| `MAX_RETRIES` | `1` | Retries per id on failure |
| `HARD_DELETE` | `true` | Whether `hardDelete: true` is passed to the mutation (full cleanup: Keycloak role, alarms, loggers, dictionaries) |

### Input Excel format

One column with one MongoDB id per row (with or without a header row, per `EXCEL_HEADER_ROWS`):

| id |
|---|
| 6a758b3f31bc64b96b4154a0 |
| 696e0b52568e2f26a99c7b78 |
| ... |

### Running it

1. Set the Keycloak environment variables.
2. Place the Excel file with the ids at the path in `EXCEL_PATH`.
3. Confirm the correct `GRAPHQL_ENDPOINT` (dev or prod).
4. Run `DeviceDeletionScript.main()`.
5. Check `device_deletion_results.csv` when it finishes.

### ⚠️ Before running against production

- **Always test in dev first** with 2-3 sample ids.
- Confirm with the team whether `hardDelete: true` is needed for full associated-data cleanup.
- Deletion in production is **irreversible**. Double-check the id list before running the full batch.
- Never share or commit Keycloak credentials (use `.env` + `.gitignore`, or system environment variables).

### Known troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| POI `Cannot resolve symbol` | Missing dependency or Maven not reloaded | Add dependency and run *Maven Reload* |
| `invalid_grant / Invalid user credentials` | Special characters in password truncated by `.env`, or wrong values | Wrap the value in quotes in `.env` |
| `Illegal character in authority` in URL | Keycloak URL copied with unresolved Spring template syntax (`${VAR:value}`) | Use the fully resolved URL, no `${}` |
| HTTP 500 calling GraphQL | Token from one environment (e.g. prod) used against another environment's endpoint (e.g. dev) | Make sure `GRAPHQL_ENDPOINT` and Keycloak credentials belong to the same environment |