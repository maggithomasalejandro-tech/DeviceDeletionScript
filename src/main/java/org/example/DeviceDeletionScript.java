package org.example;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.net.URI;
import java.net.URLEncoder;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.charset.StandardCharsets;
import java.time.Duration;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.CompletableFuture;
import java.util.Scanner;

/**
 * Script para borrar en bloque 112 devices en Resource Panorama vía GraphQL.
 *
 * Estrategia:
 *  - Lee los ids de device desde un Excel (una columna, una fila por id).
 *  - Antes de borrar, verifica que el device existe (query getDevice).
 *  - Borra en lotes de BATCH_SIZE (3), en paralelo dentro del lote.
 *  - Espera WAIT_BETWEEN_BATCHES_MS (10s) entre lotes para no saturar la API.
 *  - Si una llamada falla, reintenta una vez, loggea el resultado y continúa
 *    (no aborta el proceso completo).
 *  - Escribe un CSV de resultados: id, verificado, borrado, mensaje.
 *
 * AJUSTAR ANTES DE EJECUTAR:
 *  - GRAPHQL_ENDPOINT ya apunta al entorno DEV. Cambiar si se prueba en otro entorno.
 *  - Los datos de Keycloak (KEYCLOAK_TOKEN_URL, CLIENT_ID, CLIENT_SECRET) hay que
 *    rellenarlos con los mismos que usa metadata-migrator/GraphQLClient.java.
 *  - EXCEL_PATH y el nombre/índice de la columna con los ids.
 *  - El nombre exacto del campo de verificación (aquí "device(id: ...)") y de la
 *    mutation de borrado si difieren de lo puesto aquí.
 */
public class DeviceDeletionScript {

    // ---- Configuración ----
    private static final String GRAPHQL_ENDPOINT =
            "https://resourcepanorama-api.greenpowermonitor.com/graphql";

    // Credenciales de Keycloak leídas desde variables de entorno (no hardcodear nunca aquí).
    // Define estas variables en tu sistema o en la Run Configuration de IntelliJ:
    //   RP_KEYCLOAK_TOKEN_URL, RP_KEYCLOAK_CLIENT_ID, RP_KEYCLOAK_USERNAME, RP_KEYCLOAK_PASSWORD
    private static final String KEYCLOAK_TOKEN_URL = requireEnv("RP_KEYCLOAK_TOKEN_URL");
    private static final String CLIENT_ID = requireEnv("RP_KEYCLOAK_CLIENT_ID");
    private static final String USERNAME = requireEnv("RP_KEYCLOAK_USERNAME");
    private static final String PASSWORD = requireEnv("RP_KEYCLOAK_PASSWORD");

    private static final String EXCEL_PATH = "device_ids.xlsx";
    private static final int EXCEL_COLUMN_INDEX = 0; // columna A
    private static final int EXCEL_HEADER_ROWS = 1;  // saltar cabecera si la hay

    private static final int BATCH_SIZE = 3;
    private static final long WAIT_BETWEEN_BATCHES_MS = 10_000; // 10 segundos
    private static final int MAX_RETRIES = 1; // reintenta 1 vez si falla

    // Si hardDelete:true es necesario para limpiar rol de Keycloak, alarmas, loggers
    // y diccionarios asociados. Confirmar con el equipo antes de correr contra los 112 reales.
    private static final boolean HARD_DELETE = true;

    private static final String RESULTS_CSV = "device_deletion_results.csv";

    private static final HttpClient httpClient = HttpClient.newBuilder()
            .connectTimeout(Duration.ofSeconds(15))
            .build();

    private static String urlEncode(String value) {
        return URLEncoder.encode(value, StandardCharsets.UTF_8);
    }

    private static String requireEnv(String name) {
        String value = System.getenv(name);
        if (value == null || value.isBlank()) {
            throw new IllegalStateException(
                    "Falta la variable de entorno '" + name + "'. Defínela antes de ejecutar el script.");
        }
        return value;
    }

    public static void main(String[] args) throws Exception {
        System.out.println("TOKEN_URL: [" + KEYCLOAK_TOKEN_URL + "]");
        System.out.println("CLIENT_ID longitud=" + CLIENT_ID.length());
        System.out.println("USERNAME longitud=" + USERNAME.length() + " ultimos3=" + USERNAME.substring(Math.max(0, USERNAME.length()-3)));
        System.out.println("PASSWORD longitud=" + PASSWORD.length());
        List<String> ids = readIdsFromExcel(EXCEL_PATH, EXCEL_COLUMN_INDEX, EXCEL_HEADER_ROWS);
        System.out.println("Ids a procesar: " + ids.size());

        String token = fetchKeycloakToken();

        List<DeletionResult> results = new ArrayList<>();

        List<List<String>> batches = partition(ids, BATCH_SIZE);
        for (int batchIndex = 0; batchIndex < batches.size(); batchIndex++) {
            List<String> batch = batches.get(batchIndex);
            System.out.printf("Procesando lote %d/%d: %s%n", batchIndex + 1, batches.size(), batch);

            // Ejecutar el lote en paralelo (hasta BATCH_SIZE peticiones a la vez)
            List<CompletableFuture<DeletionResult>> futures = new ArrayList<>();
            for (String id : batch) {
                futures.add(CompletableFuture.supplyAsync(() -> processDevice(id, token)));
            }
            for (CompletableFuture<DeletionResult> f : futures) {
                results.add(f.join());
            }

            // Esperar entre lotes, salvo tras el último
            if (batchIndex < batches.size() - 1) {
                Thread.sleep(WAIT_BETWEEN_BATCHES_MS);
            }
        }

        writeResultsCsv(RESULTS_CSV, results);

        long ok = results.stream().filter(r -> r.deleted).count();
        System.out.printf("Terminado. Borrados OK: %d/%d. Detalle en %s%n", ok, results.size(), RESULTS_CSV);
    }

    /** Verifica y borra un device individual, con un reintento si falla. */
    private static DeletionResult processDevice(String id, String token) {
        int attempt = 0;
        Exception lastError = null;

        while (attempt <= MAX_RETRIES) {
            attempt++;
            try {
                boolean exists = verifyDeviceExists(id, token);
                if (!exists) {
                    return new DeletionResult(id, false, false, "Device no encontrado, se omite el borrado");
                }

                boolean deleted = deleteDevice(id, token);
                String msg = deleted ? "Borrado correctamente" : "La API respondió deleted=false";
                return new DeletionResult(id, true, deleted, msg);

            } catch (Exception e) {
                lastError = e;
                System.out.printf("Error con id %s (intento %d): %s%n", id, attempt, e.getMessage());
            }
        }

        return new DeletionResult(id, false, false,
                "Fallo tras " + (MAX_RETRIES + 1) + " intentos: " + (lastError != null ? lastError.getMessage() : "desconocido"));
    }

    /** Query de verificación antes de borrar: Device(where: {_id: ...}) { items { _id } }. */
    private static boolean verifyDeviceExists(String id, String token) throws IOException, InterruptedException {
        String query = """
                { "query": "query { Device(where: { _id: \\"%s\\" }) { items { _id } } }" }
                """.formatted(id);

        String response = postGraphQL(query, token);
        // Si "items" viene vacío ([]), el device no existe.
        return response.contains("\"_id\"") && !response.contains("\"items\":[]") && !response.contains("\"items\": []");
    }

    /** Mutation de borrado: deleteDevice(_id, hardDelete). Devuelve el Device borrado; éxito = trae _id y no hay errores. */
    private static boolean deleteDevice(String id, String token) throws IOException, InterruptedException {
        String mutation = """
                { "query": "mutation { deleteDevice(_id: \\"%s\\", hardDelete: %s) { _id } }" }
                """.formatted(id, HARD_DELETE);

        String response = postGraphQL(mutation, token);
        if (response.contains("\"errors\"")) {
            throw new IOException("GraphQL devolvió errores: " + response);
        }
        return response.contains("\"_id\":\"" + id + "\"") || response.contains("\"_id\": \"" + id + "\"");
    }

    private static String postGraphQL(String jsonBody, String token) throws IOException, InterruptedException {
        HttpRequest request = HttpRequest.newBuilder()
                .uri(URI.create(GRAPHQL_ENDPOINT))
                .timeout(Duration.ofSeconds(20))
                .header("Content-Type", "application/json")
                .header("Authorization", "Bearer " + token)
                .POST(HttpRequest.BodyPublishers.ofString(jsonBody))
                .build();

        HttpResponse<String> response = httpClient.send(request, HttpResponse.BodyHandlers.ofString());
        if (response.statusCode() / 100 != 2) {
            throw new IOException("HTTP " + response.statusCode() + ": " + response.body());
        }
        return response.body();
    }

    /** Obtiene un token de Keycloak vía grant_type=password, igual que metadata-migrator. */
    private static String fetchKeycloakToken() throws IOException, InterruptedException {
        String form = "grant_type=" + urlEncode("password")
                + "&client_id=" + urlEncode(CLIENT_ID)
                + "&username=" + urlEncode(USERNAME)
                + "&password=" + urlEncode(PASSWORD);

        HttpRequest request = HttpRequest.newBuilder()
                .uri(URI.create(KEYCLOAK_TOKEN_URL))
                .timeout(Duration.ofSeconds(15))
                .header("Content-Type", "application/x-www-form-urlencoded")
                .POST(HttpRequest.BodyPublishers.ofString(form))
                .build();

        HttpResponse<String> response = httpClient.send(request, HttpResponse.BodyHandlers.ofString());
        if (response.statusCode() / 100 != 2) {
            throw new IOException("No se pudo obtener token de Keycloak: HTTP " + response.statusCode() + " " + response.body());
        }

        // Extracción simple del access_token sin librería JSON adicional.
        String body = response.body();
        int idx = body.indexOf("\"access_token\"");
        int start = body.indexOf('"', idx + 15) + 1;
        int end = body.indexOf('"', start);
        return body.substring(start, end);
    }

    private static List<String> readIdsFromExcel(String path, int columnIndex, int headerRows) throws IOException {
        List<String> ids = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(path);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            for (int r = headerRows; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                Cell cell = row.getCell(columnIndex);
                if (cell == null) continue;

                cell.setCellType(CellType.STRING);
                String value = cell.getStringCellValue().trim();
                if (!value.isEmpty()) {
                    ids.add(value);
                }
            }
        }
        return ids;
    }

    private static void writeResultsCsv(String path, List<DeletionResult> results) throws IOException {
        try (FileWriter writer = new FileWriter(path)) {
            writer.write("id,verificado,borrado,mensaje,timestamp\n");
            for (DeletionResult r : results) {
                writer.write(String.format("%s,%s,%s,\"%s\",%s%n",
                        r.id, r.verified, r.deleted, r.message.replace("\"", "'"), LocalDateTime.now()));
            }
        }
    }

    private static <T> List<List<T>> partition(List<T> list, int size) {
        List<List<T>> result = new ArrayList<>();
        for (int i = 0; i < list.size(); i += size) {
            result.add(list.subList(i, Math.min(i + size, list.size())));
        }
        return result;
    }

    private record DeletionResult(String id, boolean verified, boolean deleted, String message) {
    }
}