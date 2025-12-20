package ar.com.leo;

import ar.com.leo.ml.MercadoLibreAPI;
import com.google.common.util.concurrent.RateLimiter;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.io.IOException;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.ZonedDateTime;
import java.util.concurrent.ThreadLocalRandom;
import java.util.function.Supplier;

public class HttpRetryHandler {

    public static final Path BASE_SECRET_DIR = Paths.get(
            System.getenv("PROGRAMDATA") != null ? System.getenv("PROGRAMDATA")
                    : System.getProperty("java.io.tmpdir"),
            "SuperMaster", "secrets");
    private static final Logger logger = LogManager.getLogger(HttpRetryHandler.class);
    private static final int MAX_RETRIES = 3; // cantidad máxima de reintentos
    private static final int MAX_RETRIES_429 = 10; // más reintentos para 429 (rate limiting)
    private final long BASE_WAIT_MS; // espera inicial
    private final RateLimiter rateLimiter; // ✅ limitador

    private final HttpClient client;

    public HttpRetryHandler(HttpClient client, long BASE_WAIT_MS, double permitsPerSecond) {
        this.client = client;
        this.BASE_WAIT_MS = BASE_WAIT_MS;
        this.rateLimiter = RateLimiter.create(permitsPerSecond); // Requests por segundo
    }

    public HttpResponse<String> sendWithRetry(Supplier<HttpRequest> requestSupplier) {
        HttpResponse<String> response = null;

        for (int attempt = 1; attempt <= MAX_RETRIES; attempt++) {
            try {
                rateLimiter.acquire();

                HttpRequest request = requestSupplier.get(); // request actualizado
                response = client.send(request, HttpResponse.BodyHandlers.ofString());
                int status = response.statusCode();

                // ---- OK ----
                if (status >= 200 && status < 300)
                    return response;

                // ---- Token expirado ----
                if (status == 401) {
                    logger.warn("401 Unauthorized → actualizando tokens...");
                    MercadoLibreAPI.verificarTokens();
                    continue; // volverá a crear el request con token nuevo
                }

                // ---- Error de concurrencia ----
                if (status == 409 || status == 423) {
                    long waitMs = BASE_WAIT_MS + ThreadLocalRandom.current().nextInt(200, 800);
                    logger.warn("409 Conflict (KVS). Retry en " + waitMs + " ms...");
                    Thread.sleep(waitMs);
                    continue;
                }

                // ---- Too Many Requests ----
                if (status == 429) {
                    // Manejar 429 con más reintentos
                    response = handle429WithRetries(requestSupplier, response);
                    if (response != null && response.statusCode() == 429) {
                        // Si después de todos los reintentos sigue siendo 429, retornar
                        logger.error("429 Too Many Requests: máximo de reintentos alcanzado");
                        return response;
                    }
                    // Si se resolvió el 429, continuar con el flujo normal
                    if (response != null && response.statusCode() >= 200
                            && response.statusCode() < 300) {
                        return response;
                    }
                    // Si hay otro error, continuar con el loop normal
                    continue;
                }

                // ---- Errores de servidor ----
                if (status >= 500 && status < 600) {
                    long waitMs = BASE_WAIT_MS * (long) Math.pow(2, attempt - 1);
                    logger.warn("5xx Error. Retry en " + waitMs + " ms...");
                    Thread.sleep(waitMs);
                    continue;
                }

                // ---- Errores 400-499 no recuperables ----
                return response;

            } catch (IOException e) {
                long waitMs = BASE_WAIT_MS * (long) Math.pow(2, attempt - 1);
                logger.warn("IOException. Retry en " + waitMs + " ms... (" + attempt + "/"
                        + MAX_RETRIES + ")");
                try {
                    Thread.sleep(waitMs);
                } catch (InterruptedException ex) {
                    Thread.currentThread().interrupt();
                    return response;
                }
            } catch (InterruptedException e) {
                Thread.currentThread().interrupt();
                return response;
            }
        }

        return response;
    }

    private HttpResponse<String> handle429WithRetries(Supplier<HttpRequest> requestSupplier,
            HttpResponse<String> lastResponse) {
        for (int retry429 = 1; retry429 <= MAX_RETRIES_429; retry429++) {
            try {
                long waitMs = parseRetryAfter(lastResponse, BASE_WAIT_MS);
                logger.warn("429 Too Many Requests. Retry " + retry429 + "/" + MAX_RETRIES_429
                        + " en " + waitMs
                        + " ms...");
                Thread.sleep(waitMs);

                rateLimiter.acquire();
                HttpRequest request = requestSupplier.get();
                HttpResponse<String> response =
                        client.send(request, HttpResponse.BodyHandlers.ofString());

                if (response.statusCode() >= 200 && response.statusCode() < 300) {
                    return response; // Éxito
                }

                if (response.statusCode() != 429) {
                    return response; // Otro error, dejar que el loop principal lo maneje
                }

                lastResponse = response; // Seguir reintentando

            } catch (InterruptedException e) {
                Thread.currentThread().interrupt();
                return lastResponse;
            } catch (IOException e) {
                logger.warn("IOException durante reintento 429: " + e.getMessage());
                return lastResponse;
            }
        }
        return lastResponse; // Máximo de reintentos alcanzado
    }

    private long parseRetryAfter(HttpResponse<String> response, long defaultMs) {
        return response.headers().firstValue("Retry-After").map(value -> {
            try {
                // si es número → segundos
                return Long.parseLong(value) * 1000;
            } catch (NumberFormatException e) {
                try {
                    // si es fecha → calcular diferencia
                    long epoch = ZonedDateTime
                            .parse(value, java.time.format.DateTimeFormatter.RFC_1123_DATE_TIME)
                            .toInstant()
                            .toEpochMilli();

                    return Math.max(epoch - System.currentTimeMillis(), defaultMs);
                } catch (Exception ignored2) {
                    return defaultMs;
                }
            }
        }).orElse(defaultMs);
    }

}
