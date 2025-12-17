package ar.com.leo.ml;

import ar.com.leo.AppLogger;
import ar.com.leo.Util;
import ar.com.leo.ml.model.Producto;
import ar.com.leo.ml.model.ProductoData;
import javafx.concurrent.Service;
import javafx.concurrent.Task;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.ObjectMapper;

import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.*;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.concurrent.*;
import java.util.stream.Stream;

public class ScrapperService extends Service<Void> {

    public static final int POOL_SIZE = 10;
    private static final ExecutorService executor = Executors.newFixedThreadPool(POOL_SIZE);
    private static final ObjectMapper mapper = new tools.jackson.databind.ObjectMapper();

    // Sets para búsqueda más rápida de extensiones
    private static final Set<String> IMAGE_EXTENSIONS_SET = Set.of(".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp");
    private static final Set<String> VIDEO_EXTENSIONS_SET = Set.of(".mp4", ".avi", ".mov", ".mkv", ".wmv", ".flv",
            ".webm", ".m4v");

    private final File excelFile;
    private final File carpetaImagenes;
    private final File carpetaVideos;

    public ScrapperService(File excelFile, File carpetaImagenes, File carpetaVideos) {
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
                ScrapperService.this.run();
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
        try (@SuppressWarnings("unused")
        RandomAccessFile raf = new RandomAccessFile(excelFile, "rw")) {
            // Si se puede abrir, está disponible
        } catch (Exception ex) {
            throw new IllegalStateException("El excel está en uso. Cerralo antes de continuar.");
        }

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

        // Excel
        final Path excelPath = excelFile.toPath();

        // Validar que el archivo Excel exista
        if (!Files.exists(excelPath)) {
            throw new IllegalArgumentException("El archivo Excel no existe: " + excelPath +
                    ". Por favor, crea el archivo Excel antes de ejecutar el proceso.");
        }

        AppLogger.info("Abriendo archivo Excel...");

        // Configurar límite de detección de Zip bomb para archivos con alta compresión
        ZipSecureFile.setMinInflateRatio(0.001);

        try (FileInputStream fis = new FileInputStream(excelFile);
                Workbook workbook = new XSSFWorkbook(fis)) {

            // Verificar que tenga al menos 2 hojas
            if (workbook.getNumberOfSheets() < 2) {
                throw new IllegalArgumentException("El archivo Excel debe tener al menos 2 hojas. " +
                        "Hojas encontradas: " + workbook.getNumberOfSheets());
            }

            Sheet scanSheet = workbook.getSheetAt(1); // 2da hoja

            // ==========================
            // Limpiar datos existentes (excepto encabezado)
            // ==========================
            int lastRowNum = scanSheet.getLastRowNum();
            if (lastRowNum > 0) {
                AppLogger.info("Limpiando " + lastRowNum + " filas existentes...");
                // Eliminar filas desde la última hasta la primera (excepto fila 0 que es el
                // encabezado)
                // Usar removeRow en lugar de shiftRows para evitar problemas con muchas filas
                for (int i = lastRowNum; i > 0; i--) {
                    Row row = scanSheet.getRow(i);
                    if (row != null) {
                        scanSheet.removeRow(row);
                    }
                }
            }

            // Estilos (se reutilizarán más adelante)
            CellStyle headerStyle = crearHeaderStyle(workbook);
            CellStyle centeredStyle = crearCenteredStyle(workbook);

            // ==========================
            // Encabezados
            // ==========================
            Row header = scanSheet.getRow(0);
            if (header == null) {
                header = scanSheet.createRow(0);
            }
            header.createCell(0).setCellValue("ESTADO");
            header.createCell(1).setCellValue("MLA");
            header.createCell(2).setCellValue("IMAGENES");
            header.createCell(3).setCellValue("VIDEOS");
            header.createCell(4).setCellValue("SKU");
            header.createCell(5).setCellValue("URL");
            header.createCell(6).setCellValue("TIPO PUBLICACION");
            header.createCell(7).setCellValue("IMAGENES EN CARPETA");
            header.createCell(8).setCellValue("VIDEOS EN CARPETA");
            header.createCell(9).setCellValue("CONCLUSION IMAGENES");
            header.createCell(10).setCellValue("CONCLUSION VIDEOS");
            header.createCell(11).setCellValue("SCORE");
            header.createCell(12).setCellValue("NIVEL");
            header.createCell(13).setCellValue("CORREGIR");

            aplicarStyleFila(header, headerStyle);

            // ==========================
            // Cargar productos y variaciones
            // ==========================
            int rowNum = 1;
            for (ProductoData p : productoList) {
                Row row = scanSheet.createRow(rowNum++);

                // Formatear MLA: si es variación, mostrar "MLA (MLAU)"
                String mlaDisplay = p.esVariacion && p.userProductId != null
                        ? (p.mla + " (" + p.userProductId + ")")
                        : p.mla;

                row.createCell(0).setCellValue(p.status);
                row.createCell(1).setCellValue(mlaDisplay);
                row.createCell(2).setCellValue(p.cantidadImagenes);
                row.createCell(3).setCellValue(p.tieneVideo);
                row.createCell(4).setCellValue(p.sku);
                row.createCell(5).setCellValue(p.permalink);
                row.createCell(6).setCellValue(p.tipoPublicacion);
                // Crear celdas vacías para las columnas que se llenarán más adelante
                // para mantener la alineación correcta
                row.createCell(7).setCellValue(""); // IMAGENES EN CARPETA
                row.createCell(8).setCellValue(""); // VIDEOS EN CARPETA
                row.createCell(9).setCellValue(""); // CONCLUSION IMAGENES
                row.createCell(10).setCellValue(""); // CONCLUSION VIDEOS

                // SCORE: si es null, dejar celda vacía, si no, poner el número
                Cell cellScore = row.createCell(11);
                if (p.score != null) {
                    cellScore.setCellValue(p.score);
                } else {
                    cellScore.setCellValue("");
                }
                // Aplicar estilo según el nivel
                cellScore.setCellStyle(obtenerEstiloPorNivel(workbook, p.nivel));

                // NIVEL: si es null, dejar celda vacía
                Cell cellNivel = row.createCell(12);
                cellNivel.setCellValue(p.nivel != null ? p.nivel : "");
                // Aplicar estilo según el nivel
                cellNivel.setCellStyle(obtenerEstiloPorNivel(workbook, p.nivel));

                // CORREGIR: títulos de keys con status PENDING (última columna)
                row.createCell(13).setCellValue(p.corregir != null ? p.corregir : "");

                // Aplicar estilo a todas las celdas excepto SCORE y NIVEL (que ya tienen su
                // estilo)
                aplicarStyleFilaExcluyendo(row, centeredStyle, 11, 12);
            }

            // ==========================
            // Buscar archivos en carpetas y actualizar Excel (sin guardar aún)
            // ==========================
            AppLogger.info("Buscando archivos en carpetas...");
            actualizarExcelConArchivos(workbook, scanSheet, carpetaImagenesPath, carpetaVideosPath, headerStyle,
                    centeredStyle);

            // ==========================
            // Guardar archivo una sola vez al final
            // ==========================
            AppLogger.info("Guardando archivo Excel...");
            try (FileOutputStream fos = new FileOutputStream(excelFile)) {
                workbook.write(fos);
                fos.flush();
            } catch (Exception ex) {
                AppLogger.error("Error al guardar Excel: " + ex.getMessage(), ex);
                throw ex;
            } finally {
                // Limpiar caché de estilos después de usar el workbook
                limpiarCacheEstilos(workbook);
            }
            String fechaHoraFin = LocalDateTime.now().format(DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm:ss"));
            AppLogger.info("[" + fechaHoraFin + "] Proceso finalizado exitosamente.");
        } catch (Exception e) {
            throw e;
        }
    }

    /**
     * Obtiene los datos de performance de un producto: score, nivel y verificación
     * de video.
     * Usa la API de performance de MercadoLibre.
     * Para productos catálogo usa el MLA de item_relations, para variaciones usa el
     * MLAU.
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
     * Normaliza el SKU para búsquedas: extrae los primeros 7 caracteres.
     * Si el SKU tiene menos de 7 caracteres, retorna null (no se puede buscar).
     */
    private static String normalizarSkuParaBusqueda(String sku) {
        if (sku == null || sku.isEmpty()) {
            return null;
        }
        String skuUpper = sku.toUpperCase().trim();
        if (skuUpper.length() < 7) {
            return null; // SKU muy corto, no se puede buscar
        }
        return skuUpper.substring(0, 7);
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

    private static CellStyle crearHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();

        // Negrita
        XSSFFont font = ((XSSFWorkbook) workbook).createFont();
        font.setBold(true);
        style.setFont(font);

        // Fondo gris claro
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // Centrados
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        // Bordes finos
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        return style;
    }

    private static CellStyle crearCenteredStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        // Bordes finos
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        return style;
    }

    private static CellStyle crearCenteredStyleWithColor(Workbook workbook, IndexedColors color) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillForegroundColor(color.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // Bordes finos
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        return style;
    }

    // Caché de estilos para evitar crear estilos duplicados
    private static final Map<String, CellStyle> estiloCache = new HashMap<>();

    /**
     * Obtiene el estilo de celda según el nivel del producto.
     * Profesional → verde, Estándar → amarillo, Básica → rojo
     */
    private static CellStyle obtenerEstiloPorNivel(Workbook workbook, String nivel) {
        if (nivel == null || nivel.isEmpty()) {
            return obtenerEstiloCached(workbook, "DEFAULT", () -> crearCenteredStyle(workbook));
        }

        String nivelUpper = nivel.toUpperCase();
        String cacheKey;
        CellStyle style;

        if (nivelUpper.contains("PROFESIONAL")) {
            // Verde para Profesional
            cacheKey = "NIVEL_PROFESIONAL";
            style = obtenerEstiloCached(workbook, cacheKey,
                    () -> crearCenteredStyleWithColor(workbook, IndexedColors.LIGHT_GREEN));
        } else if (nivelUpper.contains("ESTÁNDAR") || nivelUpper.contains("ESTANDAR")) {
            // Amarillo para Estándar
            cacheKey = "NIVEL_ESTANDAR";
            style = obtenerEstiloCached(workbook, cacheKey,
                    () -> crearCenteredStyleWithColor(workbook, IndexedColors.LIGHT_YELLOW));
        } else if (nivelUpper.contains("BÁSICA") || nivelUpper.contains("BASICA")) {
            // Rojo para Básica
            cacheKey = "NIVEL_BASICA";
            style = obtenerEstiloCached(workbook, cacheKey,
                    () -> crearCenteredStyleWithColor(workbook, IndexedColors.ROSE));
        } else {
            // Estilo normal para otros casos
            cacheKey = "DEFAULT";
            style = obtenerEstiloCached(workbook, cacheKey, () -> crearCenteredStyle(workbook));
        }

        return style;
    }

    private static CellStyle obtenerEstiloConclusion(Workbook workbook, String conclusion) {
        if (conclusion == null) {
            return obtenerEstiloCached(workbook, "DEFAULT", () -> crearCenteredStyle(workbook));
        }

        String conclusionUpper = conclusion.toUpperCase();
        String cacheKey;
        CellStyle style;

        if (conclusionUpper.equals("OK")) {
            // Verde claro para OK
            cacheKey = "OK";
            style = obtenerEstiloCached(workbook, cacheKey,
                    () -> crearCenteredStyleWithColor(workbook, IndexedColors.LIGHT_GREEN));
        } else if (conclusionUpper.contains("CREAR")) {
            // Rojo claro para CREAR
            cacheKey = "CREAR";
            style = obtenerEstiloCached(workbook, cacheKey,
                    () -> crearCenteredStyleWithColor(workbook, IndexedColors.ROSE));
        } else if (conclusionUpper.contains("SUBIR")) {
            // Amarillo claro para SUBIR
            cacheKey = "SUBIR";
            style = obtenerEstiloCached(workbook, cacheKey,
                    () -> crearCenteredStyleWithColor(workbook, IndexedColors.LIGHT_YELLOW));
        } else {
            // Estilo normal para otros casos (ERROR, etc.)
            cacheKey = "DEFAULT";
            style = obtenerEstiloCached(workbook, cacheKey, () -> crearCenteredStyle(workbook));
        }

        return style;
    }

    private static CellStyle obtenerEstiloCached(Workbook workbook, String key,
            java.util.function.Supplier<CellStyle> styleCreator) {
        // Usar el workbook como parte de la clave para evitar conflictos entre
        // diferentes workbooks
        String fullKey = workbook.hashCode() + "_" + key;
        return estiloCache.computeIfAbsent(fullKey, k -> styleCreator.get());
    }

    private static void limpiarCacheEstilos(Workbook workbook) {
        // Limpiar estilos del workbook actual del caché
        int workbookHash = workbook.hashCode();
        estiloCache.entrySet().removeIf(entry -> entry.getKey().startsWith(workbookHash + "_"));
    }

    private static void aplicarStyleFila(Row row, CellStyle style) {
        for (int i = 0; i < row.getLastCellNum(); i++) {
            if (row.getCell(i) != null) {
                row.getCell(i).setCellStyle(style);
            }
        }
    }

    /**
     * Aplica estilo a todas las celdas de una fila excepto las columnas
     * especificadas
     */
    private static void aplicarStyleFilaExcluyendo(Row row, CellStyle style, int... columnasExcluidas) {
        Set<Integer> excluidas = new HashSet<>();
        for (int col : columnasExcluidas) {
            excluidas.add(col);
        }

        for (int i = 0; i < row.getLastCellNum(); i++) {
            if (!excluidas.contains(i) && row.getCell(i) != null) {
                row.getCell(i).setCellStyle(style);
            }
        }
    }

    private static void actualizarExcelConArchivos(Workbook workbook, Sheet scanSheet, String carpetaImagenes,
            String carpetaVideos, CellStyle headerStyle, CellStyle centeredStyle) {
        try {
            // Verificar si ya existen las columnas, si no, agregarlas
            Row header = scanSheet.getRow(0);
            int colImagenesCarpeta = 7; // Columna IMAGENES EN CARPETA
            int colVideosCarpeta = 8; // Columna VIDEOS EN CARPETA

            // Aplicar estilo al header (reutilizando el estilo ya creado)
            aplicarStyleFila(header, headerStyle);

            // Procesar cada fila (empezando desde la 1, la 0 es el header)
            int colSku = 4; // Columna SKU
            int colImagenes = 2; // Columna IMAGENES (cantidad en ML)
            int colVideos = 3; // Columna VIDEOS (SI/NO)
            int colConclusionImagenes = 9; // Columna CONCLUSION IMAGENES
            int colConclusionVideos = 10; // Columna CONCLUSION VIDEOS
            int colCorregir = 13; // Columna CORREGIR (última)
            int totalFilas = scanSheet.getLastRowNum() + 1;
            AppLogger.info("Procesando " + totalFilas + " productos en carpetas...");

            // Indexar archivos una sola vez para optimizar búsquedas (especialmente útil
            // con Google Drive)
            AppLogger.info("Indexando archivos de imágenes...");
            Map<String, Integer> cacheImagenes = indexarArchivosPorSku(carpetaImagenes, IMAGE_EXTENSIONS_SET);
            AppLogger.info("Indexando archivos de videos...");
            Map<String, Integer> cacheVideos = indexarVideosPorSku(carpetaVideos);
            AppLogger.info("Búsqueda indexada completada. Procesando productos...");

            for (int rowNum = 1; rowNum < totalFilas; rowNum++) {
                Row row = scanSheet.getRow(rowNum);
                if (row == null)
                    continue;

                Cell skuCell = row.getCell(colSku);
                if (skuCell == null)
                    continue;

                String sku;
                try {
                    sku = Util.getCellValue(skuCell);
                } catch (Exception e) {
                    AppLogger.warn("Error al leer SKU en fila " + rowNum + ": " + e.getMessage());
                    continue;
                }

                // Normalizar SKU para búsqueda (debe tener al menos 7 caracteres)
                String skuNormalizado = normalizarSkuParaBusqueda(sku);
                if (skuNormalizado == null) {
                    // SKU inválido o muy corto, continuar sin buscar
                    continue;
                }

                // Buscar imágenes usando caché (mucho más rápido)
                int cantidadImagenes = cacheImagenes.getOrDefault(skuNormalizado, 0);

                // Buscar videos usando caché (mucho más rápido)
                int cantidadVideos = cacheVideos.getOrDefault(skuNormalizado, 0);

                // Actualizar celdas de archivos en carpetas
                Cell cellImagenes = row.getCell(colImagenesCarpeta);
                if (cellImagenes == null) {
                    cellImagenes = row.createCell(colImagenesCarpeta);
                }
                cellImagenes.setCellValue(cantidadImagenes);
                cellImagenes.setCellStyle(centeredStyle);

                Cell cellVideos = row.getCell(colVideosCarpeta);
                if (cellVideos == null) {
                    cellVideos = row.createCell(colVideosCarpeta);
                }
                cellVideos.setCellValue(cantidadVideos);
                cellVideos.setCellStyle(centeredStyle);

                // Generar conclusiones separadas para imágenes y videos
                String conclusionImagenes = generarConclusionImagenes(row, colImagenes, colImagenesCarpeta);
                String conclusionVideos = generarConclusionVideos(row, colVideos, colVideosCarpeta);

                // Escribir conclusión de imágenes
                Cell cellConclusionImagenes = row.getCell(colConclusionImagenes);
                if (cellConclusionImagenes == null) {
                    cellConclusionImagenes = row.createCell(colConclusionImagenes);
                }
                cellConclusionImagenes.setCellValue(conclusionImagenes);
                cellConclusionImagenes.setCellStyle(obtenerEstiloConclusion(workbook, conclusionImagenes));

                // Escribir conclusión de videos
                Cell cellConclusionVideos = row.getCell(colConclusionVideos);
                if (cellConclusionVideos == null) {
                    cellConclusionVideos = row.createCell(colConclusionVideos);
                }
                cellConclusionVideos.setCellValue(conclusionVideos);
                cellConclusionVideos.setCellStyle(obtenerEstiloConclusion(workbook, conclusionVideos));

                if ((rowNum % 100) == 0) {
                    AppLogger.info("Procesados " + rowNum + " de " + totalFilas + " productos...");
                }
            }

            // Ajustar ancho de columnas (solo las que pueden haber cambiado)
            scanSheet.autoSizeColumn(colImagenesCarpeta);
            scanSheet.autoSizeColumn(colVideosCarpeta);
            scanSheet.autoSizeColumn(colConclusionImagenes);
            scanSheet.autoSizeColumn(colConclusionVideos);
            scanSheet.autoSizeColumn(colCorregir);
            // Re-ajustar todas las columnas al final para asegurar que todo esté bien
            for (int i = 0; i <= 13; i++) {
                scanSheet.autoSizeColumn(i);
            }

            AppLogger.info("Excel actualizado con información de archivos en carpetas.");

        } catch (Exception e) {
            AppLogger.error("Error al actualizar Excel con archivos: " + e.getMessage(), e);
            throw e;
        }
    }

    private static String generarConclusionImagenes(Row row, int colImagenes, int colImagenesCarpeta) {
        try {
            // Leer cantidad de imágenes en ML (optimizado: leer directamente como número si
            // es posible)
            int cantidadImagenesML = 0;
            Cell cellImagenes = row.getCell(colImagenes);
            if (cellImagenes != null) {
                if (cellImagenes.getCellType() == CellType.NUMERIC) {
                    cantidadImagenesML = (int) cellImagenes.getNumericCellValue();
                } else {
                    try {
                        String valorImagenes = Util.getCellValue(cellImagenes);
                        if (valorImagenes != null && !valorImagenes.isEmpty()) {
                            cantidadImagenesML = Integer.parseInt(valorImagenes);
                        }
                    } catch (Exception ignored) {
                        // Si falla, queda en 0
                    }
                }
            }

            // Leer cantidad de imágenes en carpeta (optimizado)
            int cantidadImagenesCarpeta = 0;
            Cell cellImagenesCarpeta = row.getCell(colImagenesCarpeta);
            if (cellImagenesCarpeta != null) {
                if (cellImagenesCarpeta.getCellType() == CellType.NUMERIC) {
                    cantidadImagenesCarpeta = (int) cellImagenesCarpeta.getNumericCellValue();
                } else {
                    try {
                        String valorImagenesCarpeta = Util.getCellValue(cellImagenesCarpeta);
                        if (valorImagenesCarpeta != null && !valorImagenesCarpeta.isEmpty()) {
                            cantidadImagenesCarpeta = Integer.parseInt(valorImagenesCarpeta);
                        }
                    } catch (Exception ignored) {
                        // Si falla, queda en 0
                    }
                }
            }

            // Verificar imágenes
            if (cantidadImagenesML < 6) {
                int imagenesFaltantes = 6 - cantidadImagenesML; // Cuántas faltan para llegar a 6 en ML

                if (cantidadImagenesCarpeta < 6) {
                    if (cantidadImagenesCarpeta > cantidadImagenesML) {
                        // Hay más imágenes en carpeta que en ML (pueden que no estén subidas)
                        int imagenesACrear = 6 - cantidadImagenesCarpeta;
                        return "CREAR " + imagenesACrear + " " + (imagenesACrear == 1 ? "imagen" : "imágenes");
                    } else {
                        // No hay más imágenes en carpeta que en ML, solo crear las faltantes
                        return "CREAR " + imagenesFaltantes + " " + (imagenesFaltantes == 1 ? "imagen" : "imágenes");
                    }
                } else if (cantidadImagenesCarpeta == 6) {
                    // Ya hay 6 en carpeta, solo subir las faltantes
                    return "SUBIR " + imagenesFaltantes + " " + (imagenesFaltantes == 1 ? "imagen" : "imágenes");
                } else {
                    // Hay más de 6 imágenes en carpeta
                    return "SUBIR " + imagenesFaltantes + " " + (imagenesFaltantes == 1 ? "imagen" : "imágenes")
                            + "  (se pueden subir hasta " + (cantidadImagenesCarpeta - cantidadImagenesML) + " más)";
                }
            }

            // Si todo está OK
            return "OK";

        } catch (Exception e) {
            AppLogger.warn("Error al generar conclusión de imágenes: " + e.getMessage());
            return "ERROR";
        }
    }

    private static String generarConclusionVideos(Row row, int colVideos, int colVideosCarpeta) {
        try {
            // Leer si tiene video en ML (optimizado)
            String tieneVideoStr = "NO";
            Cell cellVideos = row.getCell(colVideos);
            if (cellVideos != null) {
                if (cellVideos.getCellType() == CellType.STRING) {
                    tieneVideoStr = cellVideos.getStringCellValue();
                } else {
                    try {
                        tieneVideoStr = Util.getCellValue(cellVideos);
                    } catch (Exception ignored) {
                        // Si falla, queda "NO"
                    }
                }
            }
            boolean tieneVideoML = "SI".equalsIgnoreCase(tieneVideoStr);

            // Leer cantidad de videos en carpeta (optimizado)
            int cantidadVideosCarpeta = 0;
            Cell cellVideosCarpeta = row.getCell(colVideosCarpeta);
            if (cellVideosCarpeta != null) {
                if (cellVideosCarpeta.getCellType() == CellType.NUMERIC) {
                    cantidadVideosCarpeta = (int) cellVideosCarpeta.getNumericCellValue();
                } else {
                    try {
                        String valorVideosCarpeta = Util.getCellValue(cellVideosCarpeta);
                        if (valorVideosCarpeta != null && !valorVideosCarpeta.isEmpty()) {
                            cantidadVideosCarpeta = Integer.parseInt(valorVideosCarpeta);
                        }
                    } catch (Exception ignored) {
                        // Si falla, queda en 0
                    }
                }
            }

            // Verificar video
            if (!tieneVideoML) {
                if (cantidadVideosCarpeta > 0) {
                    return "SUBIR video";
                } else {
                    return "CREAR video";
                }
            }

            // Si todo está OK
            return "OK";

        } catch (Exception e) {
            AppLogger.warn("Error al generar conclusión de videos: " + e.getMessage());
            return "ERROR";
        }
    }

    /**
     * Indexa todos los archivos por SKU en un Map para búsquedas rápidas.
     * Recorre la carpeta una sola vez en lugar de hacerlo por cada SKU.
     */
    private static Map<String, Integer> indexarArchivosPorSku(String carpeta, Set<String> extensionesSet) {
        Map<String, Integer> index = new HashMap<>();

        if (carpeta == null || carpeta.isEmpty()) {
            return index;
        }

        try {
            Path carpetaPath = Paths.get(carpeta).normalize();

            if (!Files.exists(carpetaPath) || !Files.isDirectory(carpetaPath) || !Files.isReadable(carpetaPath)) {
                AppLogger.warn("No se puede acceder a la carpeta de imágenes para indexar: " + carpetaPath);
                return index;
            }

            try (Stream<Path> paths = Files.walk(carpetaPath)) {
                paths.filter(Files::isRegularFile)
                        .forEach(path -> {
                            try {
                                String nombreArchivoCompleto = path.getFileName().toString().toUpperCase();

                                int lastDot = nombreArchivoCompleto.lastIndexOf('.');
                                if (lastDot == -1 || lastDot == nombreArchivoCompleto.length() - 1) {
                                    return;
                                }

                                String extension = nombreArchivoCompleto.substring(lastDot).toLowerCase();
                                if (!extensionesSet.contains(extension)) {
                                    return;
                                }

                                String nombreSinExtension = nombreArchivoCompleto.substring(0, lastDot);
                                if (nombreSinExtension.length() < 7) {
                                    return;
                                }

                                String skuKey = nombreSinExtension.substring(0, 7);
                                index.put(skuKey, index.getOrDefault(skuKey, 0) + 1);
                            } catch (Exception e) {
                                // Ignorar errores en archivos individuales
                            }
                        });
            }
        } catch (Exception e) {
            AppLogger.warn("Error al indexar archivos de imágenes: " + e.getMessage());
        }

        return index;
    }

    /**
     * Indexa todos los videos por SKU en un Map para búsquedas rápidas.
     * Recorre la carpeta una sola vez en lugar de hacerlo por cada SKU.
     */
    private static Map<String, Integer> indexarVideosPorSku(String carpetaVideos) {
        Map<String, Integer> index = new HashMap<>();

        if (carpetaVideos == null || carpetaVideos.isEmpty()) {
            return index;
        }

        try {
            Path carpetaPath = Paths.get(carpetaVideos).normalize();

            if (!Files.exists(carpetaPath) || !Files.isDirectory(carpetaPath) || !Files.isReadable(carpetaPath)) {
                AppLogger.warn("No se puede acceder a la carpeta de videos para indexar: " + carpetaPath);
                return index;
            }

            try (Stream<Path> carpetas = Files.list(carpetaPath)) {
                carpetas.filter(Files::isDirectory)
                        .forEach(carpeta -> {
                            try {
                                String nombreCarpeta = carpeta.getFileName().toString().toUpperCase();
                                if (nombreCarpeta.length() < 7) {
                                    return;
                                }

                                String skuKey = nombreCarpeta.substring(0, 7);

                                try (Stream<Path> archivos = Files.list(carpeta)) {
                                    long count = archivos
                                            .filter(Files::isRegularFile)
                                            .filter(archivo -> {
                                                String nombreArchivo = archivo.getFileName().toString().toUpperCase();
                                                int lastDot = nombreArchivo.lastIndexOf('.');
                                                if (lastDot == -1 || lastDot == nombreArchivo.length() - 1) {
                                                    return false; // No tiene extensión válida
                                                }
                                                String extension = nombreArchivo.substring(lastDot).toLowerCase();
                                                return VIDEO_EXTENSIONS_SET.contains(extension);
                                            })
                                            .count();

                                    if (count > 0) {
                                        index.put(skuKey, index.getOrDefault(skuKey, 0) + (int) count);
                                    }
                                }
                            } catch (Exception e) {
                                // Ignorar errores en carpetas individuales
                            }
                        });
            }
        } catch (Exception e) {
            AppLogger.warn("Error al indexar videos: " + e.getMessage());
        }

        return index;
    }

}
