package ar.com.leo.ml;

import ar.com.leo.HttpRetryHandler;
import ar.com.leo.ml.model.MLCredentials;
import ar.com.leo.ml.model.Producto;
import ar.com.leo.ml.model.TokensML;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.ObjectMapper;

import java.io.File;
import java.io.IOException;
import java.net.URI;
import java.net.URLEncoder;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;
import java.util.function.Supplier;

import static ar.com.leo.HttpRetryHandler.BASE_SECRET_DIR;

public class MercadoLibreAPI {

    private static final Logger logger = LogManager.getLogger(MercadoLibreAPI.class);
    private static final Path MERCADOLIBRE_FILE = BASE_SECRET_DIR.resolve("ml_credentials.json");
    private static final Path TOKEN_FILE = BASE_SECRET_DIR.resolve("ml_tokens.json");
    private static final Object TOKEN_LOCK = new Object();
    private static final ObjectMapper mapper = new ObjectMapper();
    private static final HttpClient httpClient = HttpClient.newHttpClient();
    private static final HttpRetryHandler retryHandler = new HttpRetryHandler(httpClient, 30000L, 5); // 5 requests por
                                                                                                      // segundo
    private static MLCredentials mlCredentials;
    private static TokensML tokens;

    public static void main(String[] args) throws IOException {
        MercadoLibreAPI.inicializar();
        // String userId = MercadoLibreAPI.getUserId();

        JsonNode itemNode = MercadoLibreAPI.getItemNodeByMLA("MLA1435717113");
        System.out.println(itemNode.toPrettyString());

        // JsonNode variations = MercadoLibreAPI.obtenerVariaciones("MLA1435717113");
        // System.out.println(variations.toPrettyString());

        // JsonNode itemNodeU = MercadoLibreAPI.getItemNodeByMLAU("MLAU398170646");
        // System.out.println(itemNodeU.toPrettyString());

        // JsonNode performance = MercadoLibreAPI.getItemPerformance("MLA1100122612");
        // System.out.println(performance.toPrettyString());
    }

    public static String getUserId() throws IOException {
        MercadoLibreAPI.verificarTokens();
        final String url = "https://api.mercadolibre.com/users/me";

        final Supplier<HttpRequest> requestBuilder = () -> HttpRequest.newBuilder()
                .uri(URI.create(url))
                .header("Authorization", "Bearer " + tokens.accessToken)
                .GET()
                .build();

        HttpResponse<String> response = retryHandler.sendWithRetry(requestBuilder);

        if (response.statusCode() != 200) {
            throw new IOException("Error al obtener el user ID de ML: " + response.body());
        }

        return mapper.readTree(response.body()).get("id").asString();
    }

    public static JsonNode obtenerDatosAplicacion(String appId) {
        Supplier<HttpRequest> requestBuilder = () -> HttpRequest.newBuilder()
                .uri(URI.create("https://api.mercadolibre.com/applications/" + appId))
                .header("Authorization", "Bearer " + tokens.accessToken)
                .GET()
                .build();

        HttpResponse<String> response = retryHandler.sendWithRetry(requestBuilder);
        if (response.statusCode() != 200) {
            logger.warn("Error obteniendo datos de la aplicación: " + response.body());
        }

        JsonNode datos = mapper.readTree(response.body());

        return datos;
    }

    public static JsonNode obtenerVariaciones(String itemId) {

        MercadoLibreAPI.verificarTokens();

        final String url = "https://api.mercadolibre.com/items/" + itemId + "/variations";

        final Supplier<HttpRequest> requestBuilder = () -> HttpRequest.newBuilder()
                .uri(URI.create(url))
                .header("Authorization", "Bearer " + tokens.accessToken)
                .GET()
                .build();

        HttpResponse<String> response = retryHandler.sendWithRetry(requestBuilder);

        if (response.statusCode() != 200) {
            logger.warn("Error al obtener las variaciones item: " + itemId + ": " + response.body());
            return null;
        }

        return mapper.readTree(response.body());
    }

    public static List<String> obtenerTodosLosItemsId(String userId) throws InterruptedException {
        final List<String> items = new ArrayList<>();
        String scrollId = null;
        boolean continuar = true;

        do {
            // Construir URL con search_type=scan
            String url = String.format("https://api.mercadolibre.com/users/%s/items/search?search_type=scan", userId);
            if (scrollId != null) {
                url += "&scroll_id=" + URLEncoder.encode(scrollId, StandardCharsets.UTF_8);
            }

            final String finalUrl = url;
            Supplier<HttpRequest> requestBuilder = () -> HttpRequest.newBuilder()
                    .uri(URI.create(finalUrl))
                    .header("Authorization", "Bearer " + tokens.accessToken)
                    .GET()
                    .build();

            HttpResponse<String> response = retryHandler.sendWithRetry(requestBuilder);

            if (response.statusCode() != 200) {
                logger.warn("ML - Error al obtener items: " + response.body());
                return null;
            }

            JsonNode root = mapper.readTree(response.body());
            JsonNode results = root.path("results");

            // Agregar los IDs de los ítems a la lista
            if (results.isArray()) {
                for (JsonNode item : results) {
                    items.add(item.asString());
                }
            }

            // Obtener el siguiente scroll_id para continuar
            if (root.has("scroll_id") && !root.get("scroll_id").isNull()) {
                scrollId = root.get("scroll_id").asString();
            } else {
                continuar = false; // No hay más resultados
            }

            // Si ya no hay resultados, detenemos el bucle
            if (results.isEmpty()) {
                continuar = false;
            }

            // Pausa corta para evitar rate limit
            // Thread.sleep(200);

        } while (continuar);

        return items;
    }

    public static Producto getItemByMLA(String itemId) {
        MercadoLibreAPI.verificarTokens();
        final String url = "https://api.mercadolibre.com/items/" + itemId;

        final Supplier<HttpRequest> requestBuilder = () -> HttpRequest.newBuilder()
                .uri(URI.create(url))
                .header("Authorization", "Bearer " + tokens.accessToken)
                .GET()
                .build();

        HttpResponse<String> response = retryHandler.sendWithRetry(requestBuilder);

        if (response.statusCode() != 200) {
            logger.warn("ML - No se pudo obtener item: " + itemId + ": " + response.body());
            // throw new IOException("Error al obtener el producto: " + itemId +
            // response.body());
        }

        // Convertir JSON → objeto MeliItem
        return mapper.readValue(response.body(), Producto.class);
    }

    public static JsonNode getItemNodeByMLA(String itemId) {

        final String url = "https://api.mercadolibre.com/items/" + itemId;

        final Supplier<HttpRequest> requestBuilder = () -> HttpRequest.newBuilder()
                .uri(URI.create(url))
                .header("Authorization", "Bearer " + tokens.accessToken)
                .GET()
                .build();

        HttpResponse<String> response = retryHandler.sendWithRetry(requestBuilder);

        if (response.statusCode() != 200) {
            logger.warn("ML - Error al obtener el producto: " + response.body());
            return null;
        }

        return mapper.readTree(response.body());
    }

    public static JsonNode getItemNodeByMLAU(String mlau) {

        final String url = "https://api.mercadolibre.com/user-products/" + mlau;

        final Supplier<HttpRequest> requestBuilder = () -> HttpRequest.newBuilder()
                .uri(URI.create(url))
                .header("Authorization", "Bearer " + tokens.accessToken)
                .GET()
                .build();

        HttpResponse<String> response = retryHandler.sendWithRetry(requestBuilder);

        if (response.statusCode() != 200) {
            logger.warn("ML - Error al obtener el producto: " + response.body());
            return null;
        }

        return mapper.readTree(response.body());
    }

    /**
     * Obtiene la calidad/performance de una publicación de MercadoLibre.
     * 
     * @param itemId ID del item (MLA) de MercadoLibre
     * @return JsonNode con los datos de performance, o null si hay error
     */
    public static JsonNode getItemPerformance(String itemId) {
        MercadoLibreAPI.verificarTokens();
        final String url = "https://api.mercadolibre.com/item/" + itemId + "/performance";

        final Supplier<HttpRequest> requestBuilder = () -> HttpRequest.newBuilder()
                .uri(URI.create(url))
                .header("Authorization", "Bearer " + tokens.accessToken)
                .GET()
                .build();

        HttpResponse<String> response = retryHandler.sendWithRetry(requestBuilder);

        if (response.statusCode() != 200) {
            logger.warn("ML - Error al obtener performance del item " + itemId + ": " + response.body());
            return null;
        }

        return mapper.readTree(response.body());
    }

    // TOKENS
    // -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    // --- MÉTODO PRINCIPAL ---
    public static boolean inicializar() {
        mlCredentials = cargarMLCredentials();
        if (mlCredentials == null) {
            logger.warn("ML - No se encontró el archivo de credenciales.");
            return false;
        }

        tokens = cargarTokens();
        if (tokens == null) {
            // No hay tokens → pedir autorización
            logger.info("ML - No hay tokens de ML, solicitando autorización...");
            final String code = pedirCodeManual();
            tokens = obtenerAccessToken(code);
            guardarTokens(tokens);
        }

        return true;
    }

    // --- MÉTODO DE VERIFICACIÓN (centralizado) ---
    public static void verificarTokens() {
        // 1️⃣ Chequeo rápido SIN bloqueo
        if (!tokens.isExpired()) {
            return;
        }

        // 2️⃣ Solo bloquear si realmente está vencido
        synchronized (TOKEN_LOCK) {
            // 3️⃣ Chequeo nuevamente dentro del lock (doble chequeo)
            if (!tokens.isExpired()) {
                return; // otro thread ya lo renovó
            }

            logger.info("ML - Access token expirado, renovando...");
            try {
                tokens = refreshAccessToken(tokens.refreshToken);
                tokens.issuedAt = System.currentTimeMillis();
                guardarTokens(tokens);
                logger.info("ML - Token renovado correctamente.");
            } catch (Exception e) {
                logger.warn("ML - Error al renovar token: " + e.getMessage());
            }
        }
    }

    // --- MÉTODOS AUXILIARES ---
    private static MLCredentials cargarMLCredentials() {
        try {
            File f = MERCADOLIBRE_FILE.toFile();
            return f.exists() ? mapper.readValue(f, MLCredentials.class) : null;
        } catch (Exception e) {
            logger.warn("Error cargando credenciales ML: " + e.getMessage());
            return null;
        }
    }

    private static TokensML cargarTokens() {
        try {
            File f = TOKEN_FILE.toFile();
            return f.exists() ? mapper.readValue(f, TokensML.class) : null;
        } catch (Exception e) {
            logger.warn("Error cargando tokens ML: " + e.getMessage());
            return null;
        }
    }

    private static void guardarTokens(TokensML tokens) {
        try {
            mapper.writerWithDefaultPrettyPrinter().writeValue(TOKEN_FILE.toFile(), tokens);
            logger.info("ML - Tokens guardados en " + TOKEN_FILE);
        } catch (Exception e) {
            logger.warn("Error guardando tokens ML: " + e.getMessage());
        }
    }

    private static String pedirCodeManual() {
        String authURL = "https://auth.mercadolibre.com.ar/authorization?response_type=code"
                + "&client_id=" + mlCredentials.clientId
                + "&redirect_uri=" + mlCredentials.redirectUri;

        logger.info("Abrí esta URL en tu navegador y autorizá la app:");
        logger.info(authURL);
        logger.info("Pegá el code que recibiste:");

        Scanner scanner = new Scanner(System.in);
        String code = scanner.nextLine().trim();
        scanner.close();
        return code;
    }

    private static TokensML obtenerAccessToken(String code) {

        Supplier<HttpRequest> requestBuilder = () -> HttpRequest.newBuilder()
                .uri(URI.create("https://api.mercadolibre.com/oauth/token"))
                .header("Content-Type", "application/x-www-form-urlencoded")
                .POST(HttpRequest.BodyPublishers.ofString(
                        "grant_type=authorization_code" +
                                "&client_id=" + mlCredentials.clientId +
                                "&client_secret=" + mlCredentials.clientSecret +
                                "&code=" + code +
                                "&redirect_uri=" + mlCredentials.redirectUri))
                .build();

        HttpResponse<String> response = retryHandler.sendWithRetry(requestBuilder);
        if (response.statusCode() != 200) {
            throw new RuntimeException("Error al obtener access_token: " + response.body());
        }

        TokensML tokens = mapper.readValue(response.body(), TokensML.class);
        tokens.issuedAt = System.currentTimeMillis();
        return tokens;
    }

    private static TokensML refreshAccessToken(String refreshToken) {
        Supplier<HttpRequest> requestBuilder = () -> HttpRequest.newBuilder()
                .uri(URI.create("https://api.mercadolibre.com/oauth/token"))
                .header("Content-Type", "application/x-www-form-urlencoded")
                .POST(HttpRequest.BodyPublishers.ofString(
                        "grant_type=refresh_token" +
                                "&client_id=" + mlCredentials.clientId +
                                "&client_secret=" + mlCredentials.clientSecret +
                                "&refresh_token=" + refreshToken))
                .build();

        HttpResponse<String> response = retryHandler.sendWithRetry(requestBuilder);
        if (response.statusCode() != 200) {
            throw new RuntimeException("Error al refrescar access_token: " + response.body());
        }

        TokensML tokens = mapper.readValue(response.body(), TokensML.class);
        tokens.issuedAt = System.currentTimeMillis();
        return tokens;
    }

}
