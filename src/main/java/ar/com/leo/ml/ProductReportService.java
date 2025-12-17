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

import java.io.File;
import java.io.IOException;
import java.nio.file.AccessDeniedException;
import java.nio.file.FileSystemException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.concurrent.*;
import java.util.stream.Stream;

public class ProductReportService extends Service<Void> {

    public static final int POOL_SIZE = 10;
    private static final ExecutorService executor = Executors.newFixedThreadPool(POOL_SIZE);
    private static final ObjectMapper mapper = new tools.jackson.databind.ObjectMapper();

    private final File excelFile;
    private final File carpetaImagenes;
    private final File carpetaVideos;

    public ProductReportService(File excelFile, File carpetaImagenes, File carpetaVideos) {
        this.excelFile = excelFile;
        this.carpetaImagenes = carpetaImagenes;
        this.carpetaVideos = carpetaVideos;
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
            ExcelUpdater.actualizarExcelConArchivos(workbook, scanSheet, carpetaImagenesPath, carpetaVideosPath,
                    headerStyle, centeredStyle);

            // Ajustar ancho de columnas
            ExcelWriter.ajustarAnchoColumnas(scanSheet);

            // Guardar archivo
            ExcelManager.guardarWorkbook(workbook, excelFile);

            String fechaHoraFin = LocalDateTime.now().format(DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm:ss"));
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
            // Obtener el ID del producto para el performance
            // Para productos normales: usar mlaParaPerformance (puede ser de item_relations
            // si es catálogo)
            // Para variaciones: usar el MLA del padre (mla)
            String itemId = productoData.esVariacion
                    ? productoData.mla // Para variaciones, usar el MLA del padre
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

            // Extraer títulos de variables con status PENDING (solo de variables, no de
            // buckets)
            List<String> titulosPendientes = new ArrayList<>();
            JsonNode buckets = performance.path("buckets");
            if (buckets.isArray()) {
                for (JsonNode bucket : buckets) {
                    // Verificar variables dentro del bucket
                    JsonNode variables = bucket.path("variables");
                    if (variables.isArray()) {
                        for (JsonNode variable : variables) {
                            JsonNode variableStatusNode = variable.path("status");
                            String variableStatus = variableStatusNode.isNull() ? "" : variableStatusNode.asString();
                            if ("PENDING".equals(variableStatus)) {
                                JsonNode variableTitleNode = variable.path("title");
                                if (!variableTitleNode.isNull()) {
                                    String variableTitle = variableTitleNode.asString("");
                                    if (variableTitle != null && !variableTitle.isEmpty()) {
                                        titulosPendientes.add(variableTitle);
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

            // Buscar en buckets → buscar bucket con type "USER_PRODUCT" para verificar
            // video
            if (!buckets.isArray()) {
                return "NO";
            }

            for (JsonNode bucket : buckets) {
                JsonNode bucketTypeNode = bucket.path("type");
                String bucketType = bucketTypeNode.isNull() ? "" : bucketTypeNode.asString();

                // Buscar el bucket de tipo USER_PRODUCT
                if ("USER_PRODUCT".equals(bucketType)) {
                    JsonNode variables = bucket.path("variables");
                    if (!variables.isArray()) {
                        continue;
                    }

                    // Buscar la variable con key "UP_SHORTS"
                    for (JsonNode variable : variables) {
                        JsonNode variableKeyNode = variable.path("key");
                        String variableKey = variableKeyNode.isNull() ? "" : variableKeyNode.asString();

                        if ("UP_SHORTS".equals(variableKey)) {
                            JsonNode rules = variable.path("rules");
                            if (!rules.isArray()) {
                                continue;
                            }

                            // Buscar la regla con key "UP_HAS_SHORTS"
                            for (JsonNode rule : rules) {
                                JsonNode ruleKeyNode = rule.path("key");
                                String ruleKey = ruleKeyNode.isNull() ? "" : ruleKeyNode.asString();

                                if ("UP_HAS_SHORTS".equals(ruleKey)) {
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
            }

            // Si no se encontró la regla, asumir que no tiene video
            return "NO";

        } catch (Exception e) {
            AppLogger.error("Error al verificar video por performance para " + productoData.mla + ": " + e.getMessage(),
                    e);
            return "ERROR";
        }
    }

    public static List<ProductoData> obtenerDatos() throws Exception, InterruptedException, IOException {

        MercadoLibreAPI.inicializar();

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
                        AppLogger.info(
                                "ML - Item " + producto.id + " tiene " + producto.variations.size() + " variaciones");

                        // Recorrer cada variación
                        for (Object variationObj : producto.variations) {
                            // Convertir Object a JsonNode para acceder a los campos
                            JsonNode variation = mapper.valueToTree(variationObj);

                            // Obtener user_product_id de la variación
                            JsonNode userProductIdNode = variation.path("user_product_id");
                            String userProductId = userProductIdNode.isNull() ? null : userProductIdNode.asString("");

                            if (userProductId != null && !userProductId.isEmpty()) {
                                // Obtener datos de la variación usando getItemNodeByMLAU
                                JsonNode variacionNode = MercadoLibreAPI.getItemNodeByMLAU(userProductId);
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

                                        AppLogger.info("ML - Variación " + userProductId + " - SKU: " + sku
                                                + " - Imágenes: " + cantidadImagenes);
                                        // Agregar la variación como ProductoData
                                        productoList
                                                .add(new ProductoData(producto, userProductId, sku, cantidadImagenes));
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
                        "No se tienen permisos de lectura en la carpeta de " + tipoCarpeta + ": " + carpetaPath);
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
                    "Error del sistema de archivos al acceder a la carpeta de " + tipoCarpeta +
                            " (verifica conectividad de red si es una ruta UNC): " + rutaCarpeta,
                    e);
        } catch (IOException e) {
            throw new IllegalArgumentException(
                    "Error de I/O al validar la carpeta de " + tipoCarpeta + ": " + rutaCarpeta, e);
        } catch (Exception e) {
            throw new IllegalArgumentException(
                    "Error al validar la carpeta de " + tipoCarpeta + ": " + rutaCarpeta, e);
        }
    }

}
