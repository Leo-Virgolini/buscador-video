package ar.com.leo.ml;

import ar.com.leo.AppLogger;
import ar.com.leo.ml.excel.ExcelManager;
import ar.com.leo.ml.excel.ExcelStyleManager;
import ar.com.leo.ml.excel.ExcelUpdater;
import ar.com.leo.ml.excel.ExcelWriter;
import ar.com.leo.ml.model.Producto;
import ar.com.leo.ml.model.ProductoData;
import javafx.concurrent.Service;
import javafx.concurrent.Task;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.ObjectMapper;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import com.google.common.util.concurrent.RateLimiter;
import java.io.File;
import java.io.IOException;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.file.AccessDeniedException;
import java.nio.file.FileSystemException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.Duration;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.concurrent.*;
import java.util.stream.Stream;

public class ProductReportService extends Service<Void> {

    public static final int POOL_SIZE = 10;
    private static final ExecutorService executor = Executors.newFixedThreadPool(POOL_SIZE);
    private static final ObjectMapper mapper = new tools.jackson.databind.ObjectMapper();

    private static final int TIMEOUT_SECONDS = 15;
    private static final String BUSQUEDA = "alt=\"clip-icon\"";
    private static final HttpClient httpClient =
            HttpClient.newBuilder().connectTimeout(Duration.ofSeconds(TIMEOUT_SECONDS)).build();

    private final File excelFile;
    private final File carpetaImagenes;
    private final File carpetaVideos;
    private final String cookieHeader;
    private final double requestsPorSegundo;
    private RateLimiter videoRateLimiter; // Rate limiter dinámico

    public ProductReportService(File excelFile, File carpetaImagenes, File carpetaVideos,
            String cookieHeader, double requestsPorSegundo) {
        this.excelFile = excelFile;
        this.carpetaImagenes = carpetaImagenes;
        this.carpetaVideos = carpetaVideos;
        this.cookieHeader = cookieHeader;
        this.requestsPorSegundo = requestsPorSegundo;
        this.videoRateLimiter = RateLimiter.create(requestsPorSegundo);
    }

    @Override
    protected Task<Void> createTask() {
        return new Task<>() {
            @Override
            protected Void call() throws Exception {
                AppLogger.setUiLogger(message -> updateMessage(message));
                ProductReportService.this.run();
                return null;
            }
        };
    }

    private static void ejecutarBloque(List<Callable<Void>> tasks) throws Exception {
        for (Future<Void> f : executor.invokeAll(tasks)) {
            try {
                f.get();
            } catch (ExecutionException e) {
                AppLogger.error("Error en una tarea: " + e.getCause().getMessage(), e.getCause());
            } catch (Exception e) {
                AppLogger.error("Error inesperado en una tarea.", e);
            }
        }
    }

    public static void shutdownExecutors() {
        executor.shutdown();
        try {
            if (!executor.awaitTermination(5, TimeUnit.SECONDS)) {
                executor.shutdownNow();
            }
        } catch (InterruptedException e) {
            executor.shutdownNow();
        }
    }

    public void run() throws Exception {
        // Validar cookies primero (antes de validar archivos/carpetas para fallar rápido)
        if (!cookiesValidas(cookieHeader)) {
            throw new IllegalArgumentException(
                    "Cookies inválidas. Por favor verifica que estés logueado en MercadoLibre.");
        }
        AppLogger.info("Cookies válidas.");

        // Verificar que el archivo no esté en uso
        ExcelManager.verificarArchivoDisponible(excelFile);

        // Obtener rutas absolutas (funciona con rutas locales y de red UNC)
        String carpetaImagenesPath = carpetaImagenes.getAbsolutePath();
        String carpetaVideosPath = carpetaVideos.getAbsolutePath();

        // Validar que las carpetas sean accesibles antes de continuar
        validarAccesoCarpeta(carpetaImagenesPath, "imágenes");
        validarAccesoCarpeta(carpetaVideosPath, "videos");

        final List<ProductoData> productoList = obtenerDatos();

        AppLogger.info("Verificando videos en " + productoList.size() + " productos...");
        List<Callable<Void>> tasks = new ArrayList<>();
        for (ProductoData productoData : productoList) {
            tasks.add(() -> {
                // Obtener datos de performance (score, nivel y video)
                String videoResult = obtenerDatosDePerformance(productoData);
                productoData.tieneVideo = videoResult;
                return null;
            });
        }
        ejecutarBloque(tasks);
        AppLogger.info("Verificación de videos completada.");

        // Ordenamiento
        productoList.sort(Comparator
                .comparing((ProductoData p) -> p.status, Comparator.nullsFirst(String::compareTo))
                .thenComparing(p -> p.mla, Comparator.nullsFirst(String::compareTo))
                .thenComparing(p -> p.cantidadImagenes)
                .thenComparing(p -> p.tieneVideo, Comparator.nullsFirst(String::compareTo))
                .thenComparing(p -> p.sku, Comparator.nullsFirst(String::compareTo)));

        // Validar que el archivo Excel exista
        ExcelManager.validarArchivoExiste(excelFile);

        // Abrir workbook
        Workbook workbook = ExcelManager.abrirWorkbook(excelFile);
        try {
            Sheet scanSheet = ExcelManager.obtenerHojaEscaneo(workbook);

            // Limpiar datos existentes (excepto encabezado)
            int lastRowNum = scanSheet.getLastRowNum();
            if (lastRowNum > 0) {
                AppLogger.info("Limpiando " + lastRowNum + " filas existentes...");
            }
            ExcelWriter.limpiarDatosExistentes(scanSheet);

            // Crear estilos
            CellStyle headerStyle = ExcelStyleManager.crearHeaderStyle(workbook);
            CellStyle centeredStyle = ExcelStyleManager.crearCenteredStyle(workbook);

            // Crear encabezados
            ExcelWriter.crearEncabezados(scanSheet, headerStyle);

            // Escribir productos
            ExcelWriter.escribirProductos(scanSheet, productoList, workbook, centeredStyle);

            // Buscar archivos en carpetas y actualizar Excel
            AppLogger.info("Buscando archivos en carpetas...");
            ExcelUpdater.actualizarExcelConArchivos(workbook, scanSheet, carpetaImagenesPath,
                    carpetaVideosPath, headerStyle, centeredStyle);

            // Ajustar ancho de columnas
            ExcelWriter.ajustarAnchoColumnas(scanSheet);

            // Guardar archivo
            ExcelManager.guardarWorkbook(workbook, excelFile);

            String fechaHoraFin =
                    LocalDateTime.now().format(DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm:ss"));
            AppLogger.info("[" + fechaHoraFin + "] Proceso finalizado exitosamente.");
        } finally {
            workbook.close();
        }
    }

    /**
     * Obtiene los datos de performance de un producto: score, nivel y verificación
     * de video.
     * Usa la API de performance de MercadoLibre.
     * Busca la regla "UP_HAS_SHORTS" en el JSON de performance para determinar si
     * tiene video.
     * 
     * @param productoData El producto a procesar (se actualizan score, nivel y
     *                     tieneVideo)
     * @return "SI" si tiene video (status COMPLETED), "NO" si no tiene (status
     *         PENDING), o "ERROR" si hay problemas
     */
    private String obtenerDatosDePerformance(ProductoData productoData) {
        try {
            // La API de performance solo funciona si el status del producto es "active"
            // Si no está activo o es catálogo, usar scraping
            boolean usarScraping = false;
            String razon = "";

            if (productoData.status == null || !"active".equalsIgnoreCase(productoData.status)) {
                usarScraping = true;
                razon = "no está activo (status: " + productoData.status + ")";
            } else if (productoData.tipoPublicacion != null
                    && "CATALOGO".equalsIgnoreCase(productoData.tipoPublicacion)) {
                usarScraping = true;
                razon = "es de tipo catálogo";
            }

            if (usarScraping) {
                AppLogger.info("Producto " + productoData.mla + " " + razon
                        + ", usando scraping para verificar video. No se obtendrán datos de performance.");
                // Usar scraping para verificar video
                // No se obtienen datos de performance (score, nivel, corregir) porque la API no funciona
                // Establecer valores como "N/A" para indicar que no están disponibles
                productoData.score = null; // Se mantiene null para que ExcelWriter lo maneje
                productoData.nivel = "N/A";
                productoData.corregir = "N/A";

                if (productoData.permalink != null && !productoData.permalink.isEmpty()) {
                    return verificarVideo(productoData.permalink, cookieHeader);
                } else {
                    AppLogger.warn("Producto " + productoData.mla
                            + " no tiene permalink, no se puede verificar video por scraping");
                    return "NO";
                }
            }

            // Obtener el ID del producto para el performance
            // Para productos normales: usar mlaParaPerformance (puede ser de item_relations
            // si es catálogo)
            // Para variaciones: usar el MLA del padre (mla)
            String itemId = productoData.esVariacion ? productoData.mla // Para variaciones, usar el MLA del padre
                    : productoData.mlaParaPerformance;

            // Obtener performance del producto usando siempre getItemPerformanceByMLA
            JsonNode performance = MercadoLibreAPI.getItemPerformanceByMLA(itemId);

            if (performance == null) {
                AppLogger.warn("No se pudo obtener performance para " + itemId);
                return "NO";
            }

            // Extraer score y level_wording del JSON
            JsonNode scoreNode = performance.path("score");
            if (!scoreNode.isNull() && scoreNode.isNumber()) {
                productoData.score = scoreNode.asInt();
            }

            JsonNode levelWordingNode = performance.path("level_wording");
            if (!levelWordingNode.isNull()) {
                productoData.nivel = levelWordingNode.asString();
            }

            // Extraer títulos de wordings.title de rules dentro de variables con status PENDING
            List<String> titulosPendientes = new ArrayList<>();
            JsonNode buckets = performance.path("buckets");
            if (buckets.isArray()) {
                for (JsonNode bucket : buckets) {
                    // Verificar variables dentro del bucket
                    JsonNode variables = bucket.path("variables");
                    if (variables.isArray()) {
                        for (JsonNode variable : variables) {
                            JsonNode variableStatusNode = variable.path("status");
                            String variableStatus = variableStatusNode.isNull() ? ""
                                    : variableStatusNode.asString();
                            if ("PENDING".equals(variableStatus)) {
                                // Buscar en las rules de la variable
                                JsonNode rules = variable.path("rules");
                                if (rules.isArray()) {
                                    for (JsonNode rule : rules) {
                                        // Extraer el title de wordings.title
                                        JsonNode wordingsNode = rule.path("wordings");
                                        if (!wordingsNode.isNull() && wordingsNode.isObject()) {
                                            JsonNode wordingTitleNode = wordingsNode.path("title");
                                            if (!wordingTitleNode.isNull()) {
                                                String wordingTitle = wordingTitleNode.asString("");
                                                if (wordingTitle != null
                                                        && !wordingTitle.isEmpty()) {
                                                    // Evitar duplicados
                                                    if (!titulosPendientes.contains(wordingTitle)) {
                                                        titulosPendientes.add(wordingTitle);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // Unir los títulos con " | " y guardar en productoData
            if (!titulosPendientes.isEmpty()) {
                productoData.corregir = String.join(" | ", titulosPendientes);
            } else {
                productoData.corregir = "";
            }

            // Buscar en buckets para verificar video
            // El video puede aparecer en:
            // 1. Bucket "USER_PRODUCT" → variable "UP_SHORTS" → rule "UP_HAS_SHORTS"
            // 2. Bucket "CHARACTERISTICS" → variable "SHORTS" → rule "HAS_SHORTS"
            if (!buckets.isArray()) {
                return "NO";
            }

            for (JsonNode bucket : buckets) {
                JsonNode variables = bucket.path("variables");
                if (!variables.isArray()) {
                    continue;
                }

                // Buscar en todas las variables del bucket
                for (JsonNode variable : variables) {
                    JsonNode variableKeyNode = variable.path("key");
                    String variableKey = variableKeyNode.isNull() ? "" : variableKeyNode.asString();

                    // Buscar variable "SHORTS" o "UP_SHORTS"
                    if ("SHORTS".equals(variableKey) || "UP_SHORTS".equals(variableKey)) {
                        JsonNode rules = variable.path("rules");
                        if (!rules.isArray()) {
                            continue;
                        }

                        // Buscar la regla "HAS_SHORTS" o "UP_HAS_SHORTS"
                        for (JsonNode rule : rules) {
                            JsonNode ruleKeyNode = rule.path("key");
                            String ruleKey = ruleKeyNode.isNull() ? "" : ruleKeyNode.asString();

                            if ("HAS_SHORTS".equals(ruleKey) || "UP_HAS_SHORTS".equals(ruleKey)) {
                                JsonNode statusNode = rule.path("status");
                                String status = statusNode.isNull() ? "" : statusNode.asString();

                                // Si el status es "COMPLETED", tiene video
                                if ("COMPLETED".equals(status)) {
                                    return "SI";
                                } else {
                                    // PENDING u otro estado = no tiene video
                                    return "NO";
                                }
                            }
                        }
                    }
                }
            }

            // Si no se encontró la regla, asumir que no tiene video
            return "NO";

        } catch (Exception e) {
            AppLogger.error("Error al verificar video por performance para " + productoData.mla
                    + ": " + e.getMessage(), e);
            return "ERROR";
        }
    }

    public static List<ProductoData> obtenerDatos()
            throws Exception, InterruptedException, IOException {

        if (!MercadoLibreAPI.inicializar()) {
            throw new IllegalStateException(
                    "No se pudo inicializar la API de MercadoLibre. Verificar credenciales y tokens.");
        }

        final String userId = MercadoLibreAPI.getUserId();
        AppLogger.info("User ID: " + userId);

        AppLogger.info("Obteniendo MLAs de todos los productos...");
        final List<String> productos = MercadoLibreAPI.obtenerTodosLosItemsId(userId);
        AppLogger.info("Total de Productos encontrados: " + productos.size());

        final List<ProductoData> productoList = Collections.synchronizedList(new ArrayList<>());

        AppLogger.info("Obteniendo datos de todos los productos...");
        List<Callable<Void>> tasks = new ArrayList<>();
        for (String mla : productos) {
            tasks.add(() -> {
                Producto producto = MercadoLibreAPI.getItemByMLA(mla);
                if (producto != null) {
                    // Verificar si tiene variaciones (ya vienen en producto.variations)
                    if (producto.variations != null && !producto.variations.isEmpty()) {
                        AppLogger.info("ML - Item " + producto.id + " tiene "
                                + producto.variations.size() + " variaciones");

                        // Recorrer cada variación
                        for (Object variationObj : producto.variations) {
                            // Convertir Object a JsonNode para acceder a los campos
                            JsonNode variation = mapper.valueToTree(variationObj);

                            // Obtener user_product_id de la variación
                            JsonNode userProductIdNode = variation.path("user_product_id");
                            String userProductId = userProductIdNode.isNull() ? null
                                    : userProductIdNode.asString("");

                            if (userProductId != null && !userProductId.isEmpty()) {
                                // Obtener datos de la variación usando getItemNodeByMLAU
                                JsonNode variacionNode =
                                        MercadoLibreAPI.getItemNodeByMLAU(userProductId);
                                if (variacionNode != null) {
                                    // Buscar el atributo SELLER_SKU en attributes
                                    String sku = extraerSkuDeVariacion(variacionNode);
                                    if (sku != null && !sku.isEmpty()) {
                                        // Obtener cantidad de imágenes de picture_ids de la variación
                                        JsonNode pictureIdsNode = variation.path("picture_ids");
                                        int cantidadImagenes = 0;
                                        if (pictureIdsNode.isArray()) {
                                            cantidadImagenes = pictureIdsNode.size();
                                        }

                                        AppLogger
                                                .info("ML - Variación " + userProductId + " - SKU: "
                                                        + sku + " - Imágenes: " + cantidadImagenes);
                                        // Agregar la variación como ProductoData
                                        productoList.add(new ProductoData(producto, userProductId,
                                                sku, cantidadImagenes));
                                    }
                                }
                            }
                        }
                    } else {
                        // Agregar el producto principal (sin variaciones)
                        String sku = getSku(producto.attributes);
                        productoList.add(new ProductoData(producto, sku));
                    }
                }
                return null;
            });
        }
        ejecutarBloque(tasks);

        AppLogger.info("Total de productos y variaciones: " + productoList.size());
        return productoList;
    }

    /**
     * Extrae el SKU de los primeros 7 dígitos del atributo name en SELLER_SKU
     */
    private static String extraerSkuDeVariacion(JsonNode variacionNode) {
        try {
            JsonNode attributes = variacionNode.path("attributes");
            if (attributes.isArray()) {
                for (JsonNode attribute : attributes) {
                    JsonNode idNode = attribute.path("id");
                    String id = idNode.isNull() ? "" : idNode.asString();
                    if ("SELLER_SKU".equals(id)) {
                        // El SKU está en values[0].name, no en attribute.name
                        JsonNode values = attribute.path("values");
                        if (values.isArray() && values.size() > 0) {
                            JsonNode firstValue = values.get(0);
                            JsonNode nameNode = firstValue.path("name");
                            String name = nameNode.isNull() ? "" : nameNode.asString();
                            if (name != null && name.length() >= 7) {
                                // Obtener los primeros 7 dígitos
                                String primeros7Digitos = name.substring(0, 7);
                                // Verificar que sean dígitos
                                if (primeros7Digitos.matches("\\d{7}")) {
                                    return primeros7Digitos;
                                }
                            }
                        }
                    }
                }
            }
        } catch (Exception e) {
            AppLogger.warn("Error al extraer SKU de variación: " + e.getMessage());
        }
        return null;
    }

    public static String getSku(List<Producto.Attribute> attributes) {
        if (attributes == null) {
            return null;
        }
        for (Producto.Attribute a : attributes) {
            if ("SELLER_SKU".equals(a.id) && a.valueName != null) {
                return a.valueName.length() >= 7 ? a.valueName.substring(0, 7) : a.valueName;
            }
        }
        return null;
    }

    /**
     * Valida que una carpeta sea accesible (funciona con rutas locales y de red
     * UNC).
     * Lanza excepción si la carpeta no es accesible.
     */
    private static void validarAccesoCarpeta(String rutaCarpeta, String tipoCarpeta) {
        try {
            Path carpetaPath = Paths.get(rutaCarpeta).normalize();

            if (!Files.exists(carpetaPath)) {
                throw new IllegalArgumentException(
                        "La carpeta de " + tipoCarpeta + " no existe: " + carpetaPath);
            }

            if (!Files.isDirectory(carpetaPath)) {
                throw new IllegalArgumentException(
                        "La ruta de " + tipoCarpeta + " no es un directorio: " + carpetaPath);
            }

            if (!Files.isReadable(carpetaPath)) {
                throw new IllegalArgumentException(
                        "No se tienen permisos de lectura en la carpeta de " + tipoCarpeta + ": "
                                + carpetaPath);
            }

            // Intentar listar el directorio para verificar acceso real (esto puede lanzar
            // AccessDeniedException)
            try (Stream<Path> test = Files.list(carpetaPath)) {
                test.limit(1).count(); // Solo verificar que se puede acceder
            }

        } catch (AccessDeniedException e) {
            throw new IllegalArgumentException(
                    "Acceso denegado a la carpeta de " + tipoCarpeta + ": " + rutaCarpeta, e);
        } catch (FileSystemException e) {
            throw new IllegalArgumentException(
                    "Error del sistema de archivos al acceder a la carpeta de " + tipoCarpeta
                            + " (verifica conectividad de red si es una ruta UNC): " + rutaCarpeta,
                    e);
        } catch (IOException e) {
            throw new IllegalArgumentException(
                    "Error de I/O al validar la carpeta de " + tipoCarpeta + ": " + rutaCarpeta, e);
        } catch (Exception e) {
            throw new IllegalArgumentException(
                    "Error al validar la carpeta de " + tipoCarpeta + ": " + rutaCarpeta, e);
        }
    }


    private String verificarVideo(String url, String cookieHeader) {
        return verificarVideo(url, cookieHeader, 0);
    }

    private String verificarVideo(String url, String cookieHeader, int intentos) {
        // Límite de recursión para evitar StackOverflowError
        if (intentos >= 5) {
            AppLogger.warn("Límite de reintentos alcanzado para: " + url);
            return "ERROR: Límite de reintentos alcanzado";
        }

        int status = 0;
        try {
            // Aplicar rate limiting para evitar bloqueos
            videoRateLimiter.acquire();

            final HttpRequest request = HttpRequest.newBuilder().uri(URI.create(url))
                    .timeout(Duration.ofSeconds(TIMEOUT_SECONDS))
                    .header("User-Agent", "Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N)")
                    .header("Accept",
                            "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7")
                    .header("Accept-Language", "es-AR,es;q=0.9,en;q=0.8")
                    .header("Referer", "https://www.mercadolibre.com.ar/")
                    .header("Cookie", cookieHeader) // 👈 cookie
                    .GET().build();

            HttpResponse<String> response =
                    httpClient.send(request, HttpResponse.BodyHandlers.ofString());
            final String html = response.body();
            status = response.statusCode();

            switch (status) {
                case 200:
                    if (html.contains(BUSQUEDA)) {
                        return "SI";
                    } else {
                        return "NO";
                    }
                case 301:
                case 302:
                case 303:
                case 307:
                case 308:
                    if (html.startsWith("Moved Permanently") || html.contains("Redirecting to ")) {
                        int idx = html.indexOf("Redirecting to ");
                        if (idx != -1) {
                            String newUrl = html.substring(idx + "Redirecting to ".length())
                                    .replace("</p>", "").trim();
                            // Reintentar con la nueva URL (si cambió)
                            if (!newUrl.equals(request.uri().toString())) {
                                AppLogger.info(
                                        "URL vieja: " + url + " - URL actualizada: " + newUrl);
                                return verificarVideo(newUrl, cookieHeader, intentos + 1);
                            }
                        }
                    }
                    break;
                case 404:
                case 410:
                    return "NO EXISTE";
                case 403:
                case 424:
                    AppLogger.info("Too many requests.");
                    Thread.sleep(60000);
                    return verificarVideo(url, cookieHeader, intentos + 1);
                case 500:
                case 502:
                case 503:
                case 504:
                    AppLogger.info("Internal server error.");
                    Thread.sleep(5000);
                    return verificarVideo(url, cookieHeader, intentos + 1);
                default:
                    return "STATUS: " + status;
            }
        } catch (Exception e) {
            AppLogger.error(
                    "Error en url: " + url + " - status: " + status + " -> " + e.getMessage(), e);
            return "ERROR: " + e.getMessage();
        }

        return "ERROR: " + status;
    }

    public static boolean cookiesValidas(String cookieHeader) {
        try {
            HttpClient client = HttpClient.newBuilder().followRedirects(HttpClient.Redirect.NEVER) // importante
                    .build();

            HttpRequest request = HttpRequest.newBuilder()
                    .uri(URI.create("https://www.mercadolibre.com.ar/pampa/profile"))
                    .header("User-Agent", "Mozilla/5.0").header("Cookie", cookieHeader).GET()
                    .build();

            HttpResponse<String> response =
                    client.send(request, HttpResponse.BodyHandlers.ofString());

            int status = response.statusCode();

            // ============================
            // VALIDACIONES DE LOGIN
            // ============================

            // 302 → redirige al login → NO logeado
            if (status == 302)
                return false;

            // 401 o 403 → NO autorizado → NO logeado
            if (status == 401 || status == 403)
                return false;

            // 200 → verificar contenido
            if (status == 200) {
                String body = response.body();
                // si contiene datos personales → usuario logeado
                if (body.contains("myaccount") || body.contains("Mi cuenta")
                        || body.contains("profile")) {
                    return true;
                }
            }

            return false;

        } catch (Exception e) {
            AppLogger.error("Error verificando cookies: " + e.getMessage(), e);
            return false;
        }
    }

}
