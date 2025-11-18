package ar.com.leo;

import ar.com.leo.ml.MercadoLibreAPI;
import com.google.common.util.concurrent.RateLimiter;

import java.io.IOException;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.util.concurrent.ThreadLocalRandom;
import java.util.function.Supplier;

public class HttpRetryHandler {

    private static final int MAX_RETRIES = 3; // cantidad máxima de reintentos
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
                    AppLogger.warn("401 Unauthorized → actualizando tokens...");
                    MercadoLibreAPI.verificarTokens();
                    continue; // volverá a crear el request con token nuevo
                }

                // ---- Error de concurrencia ----
                if (status == 409 || status == 423) {
                    long waitMs = BASE_WAIT_MS + ThreadLocalRandom.current().nextInt(200, 800);
                    AppLogger.warn("409 Conflict (KVS). Retry en " + waitMs + " ms...");
                    Thread.sleep(waitMs);
                    continue;
                }

                // ---- Too Many Requests ----
                if (status == 429) {
                    long waitMs = parseRetryAfter(response, BASE_WAIT_MS);
                    AppLogger.warn("429 Too Many Requests. Retry en " + waitMs + " ms...");
                    Thread.sleep(waitMs);
                    continue;
                }

                // ---- Errores de servidor ----
                if (status >= 500 && status < 600) {
                    long waitMs = BASE_WAIT_MS * (long) Math.pow(2, attempt - 1);
                    AppLogger.warn("5xx Error. Retry en " + waitMs + " ms...");
                    Thread.sleep(waitMs);
                    continue;
                }

                // ---- Errores 400-499 no recuperables ----
                return response;

            } catch (IOException e) {
                long waitMs = BASE_WAIT_MS * (long) Math.pow(2, attempt - 1);
                AppLogger.warn("IOException. Retry en " + waitMs + " ms... (" + attempt + "/" + MAX_RETRIES + ")");
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

    private long parseRetryAfter(HttpResponse<String> response, long defaultMs) {
        return response.headers().firstValue("Retry-After").map(value -> {
            try {
                // si es número → segundos
                return Long.parseLong(value) * 1000;
            } catch (NumberFormatException e) {
                try {
                    // si es fecha → calcular diferencia
                    long epoch = java.time.ZonedDateTime.parse(value, java.time.format.DateTimeFormatter.RFC_1123_DATE_TIME).toInstant().toEpochMilli();

                    return Math.max(epoch - System.currentTimeMillis(), defaultMs);
                } catch (Exception ignored2) {
                    return defaultMs;
                }
            }
        }).orElse(defaultMs);
    }

}
