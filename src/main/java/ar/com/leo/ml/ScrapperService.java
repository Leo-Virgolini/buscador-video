package ar.com.leo.ml;

import ar.com.leo.AppLogger;
import ar.com.leo.Util;
import ar.com.leo.ml.model.Producto;
import javafx.concurrent.Service;
import java.util.concurrent.Future;
import java.util.concurrent.Callable;
import java.util.List;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.openxml4j.util.ZipSecureFile;

import java.util.concurrent.ExecutionException;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.RandomAccessFile;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.Duration;
import java.util.*;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;
import java.util.stream.Collectors;
import java.util.stream.Stream;
import javafx.concurrent.Task;

public class ScrapperService extends Service<Void> {

    public static final int POOL_SIZE = 10;
    private static final ExecutorService executor = Executors.newFixedThreadPool(POOL_SIZE);

    private static final int TIMEOUT_SECONDS = 15;
    private static final String BUSQUEDA = "alt=\"clip-icon\"";
    private static final HttpClient httpClient = HttpClient.newBuilder()
            .connectTimeout(Duration.ofSeconds(TIMEOUT_SECONDS)).build();

    // Sets para búsqueda más rápida de extensiones
    private static final Set<String> IMAGE_EXTENSIONS_SET = Set.of(".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp");
    private static final Set<String> VIDEO_EXTENSIONS_SET = Set.of(".mp4", ".avi", ".mov", ".mkv", ".wmv", ".flv",
            ".webm", ".m4v");

    private final File excelFile;
    private final File carpetaImagenes;
    private final File carpetaVideos;
    private final String cookieHeader;

    public ScrapperService(File excelFile, File carpetaImagenes, File carpetaVideos, String cookieHeader) {
        this.excelFile = excelFile;
        this.carpetaImagenes = carpetaImagenes;
        this.carpetaVideos = carpetaVideos;
        this.cookieHeader = cookieHeader;
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

    public void run() throws Exception, IOException, InterruptedException {

        // Verificar que el archivo no esté en uso
        try (RandomAccessFile raf = new RandomAccessFile(excelFile, "rw")) {
            // Si se puede abrir, está disponible
        } catch (Exception ex) {
            throw new IllegalStateException("El excel está en uso. Cerralo antes de continuar.");
        }
        String carpetaImagenesPath = carpetaImagenes.getAbsolutePath();
        String carpetaVideosPath = carpetaVideos.getAbsolutePath();

        if (cookiesValidas(cookieHeader)) {
            AppLogger.info("Cookies válidas.");

            final List<Producto> productoList = obtenerDatos();

            AppLogger.info("Verificando videos en " + productoList.size() + " productos...");
            List<Callable<Void>> tasks = new ArrayList<>();
            for (Producto producto : productoList) {
                tasks.add(() -> {
                    producto.videoId = verificarVideo(producto.permalink, cookieHeader);
                    return null;
                });
            }
            ejecutarBloque(tasks);
            AppLogger.info("Verificación de videos completada.");

            // Ordenamiento
            productoList.sort(Comparator
                    .comparing((Producto p) -> p.status, Comparator.nullsFirst(String::compareTo))
                    .thenComparing(p -> p.id, Comparator.nullsFirst(String::compareTo))
                    .thenComparing(p -> p.pictures == null ? 0 : p.pictures.size())
                    .thenComparing(p -> p.videoId != null ? p.videoId.toString() : "NO",
                            Comparator.nullsFirst(String::compareTo))
                    .thenComparing(p -> getSku(p.attributes), Comparator.nullsFirst(String::compareTo)));

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
                    // Usar shiftRows para eliminar todas las filas de una vez (más eficiente)
                    scanSheet.shiftRows(1, lastRowNum, -lastRowNum);
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

                aplicarStyleFila(header, headerStyle);

                // ==========================
                // Cargar productos
                // ==========================
                int rowNum = 1;
                for (Producto p : productoList) {
                    Row row = scanSheet.createRow(rowNum++);
                    int cantidadImagenes = p.pictures != null ? p.pictures.size() : 0;
                    String tieneVideo = p.videoId != null ? p.videoId.toString() : "NO";

                    row.createCell(0).setCellValue(p.status);
                    row.createCell(1).setCellValue(p.id);
                    row.createCell(2).setCellValue(cantidadImagenes);
                    row.createCell(3).setCellValue(tieneVideo);
                    row.createCell(4).setCellValue(getSku(p.attributes)); // SKU
                    row.createCell(5).setCellValue(p.permalink);
                    row.createCell(6).setCellValue(p.catalogListing ? "CATALOGO" : "TRADICIONAL");

                    aplicarStyleFila(row, centeredStyle);
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
            } catch (Exception e) {
                throw e;
            }
        } else {
            throw new IllegalArgumentException(
                    "Cookies inválidas. Por favor verifica que estés logueado en MercadoLibre.");
        }
    }

    private static String verificarVideo(String url, String cookieHeader) {
        return verificarVideo(url, cookieHeader, 0);
    }

    private static String verificarVideo(String url, String cookieHeader, int intentos) {
        // Límite de recursión para evitar StackOverflowError
        if (intentos >= 5) {
            AppLogger.warn("Límite de reintentos alcanzado para: " + url);
            return "ERROR: Límite de reintentos alcanzado";
        }

        int status = 0;
        try {
            final HttpRequest request = HttpRequest.newBuilder()
                    .uri(URI.create(url))
                    .timeout(Duration.ofSeconds(TIMEOUT_SECONDS))
                    .header("User-Agent", "Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N)")
                    .header("Accept",
                            "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7")
                    .header("Accept-Language", "es-AR,es;q=0.9,en;q=0.8")
                    .header("Referer", "https://www.mercadolibre.com.ar/")
                    .header("Cookie", cookieHeader) // 👈 cookie
                    .GET()
                    .build();

            HttpResponse<String> response = httpClient.send(request, HttpResponse.BodyHandlers.ofString());
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
                            String newUrl = html.substring(idx + "Redirecting to ".length()).replace("</p>", "").trim();
                            // Reintentar con la nueva URL (si cambió)
                            if (!newUrl.equals(request.uri().toString())) {
                                AppLogger.info("URL vieja: " + url + " - URL actualizada: " + newUrl);
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
            AppLogger.error("Error en url: " + url + " - status: " + status + " -> " + e.getMessage(), e);
            return "ERROR: " + e.getMessage();
        }

        return "ERROR: " + status;
    }

    public static List<Producto> obtenerDatos() throws Exception, InterruptedException, IOException {

        MercadoLibreAPI.inicializar();

        final String userId = MercadoLibreAPI.getUserId();
        AppLogger.info("User ID: " + userId);

        AppLogger.info("Obteniendo MLAs de todos los productos...");
        final List<String> productos = MercadoLibreAPI.obtenerTodosLosItemsId(userId);
        AppLogger.info("Total de Productos encontrados: " + productos.size());

        final List<Producto> productoList = Collections.synchronizedList(new ArrayList<>());

        AppLogger.info("Obteniendo datos de todos los productos...");
        List<Callable<Void>> tasks = new ArrayList<>();
        for (String mla : productos) {
            tasks.add(() -> {
                Producto producto = MercadoLibreAPI.getItemByMLA(mla);
                if (producto != null) {
                    productoList.add(producto);
                }
                return null;
            });
        }
        ejecutarBloque(tasks);

        return productoList;
    }

    public static boolean cookiesValidas(String cookieHeader) {
        try {
            HttpClient client = HttpClient.newBuilder()
                    .followRedirects(HttpClient.Redirect.NEVER) // importante
                    .build();

            HttpRequest request = HttpRequest.newBuilder()
                    .uri(URI.create("https://www.mercadolibre.com.ar/pampa/profile"))
                    .header("User-Agent", "Mozilla/5.0")
                    .header("Cookie", cookieHeader)
                    .GET()
                    .build();

            HttpResponse<String> response = client.send(request, HttpResponse.BodyHandlers.ofString());

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
                if (body.contains("myaccount") || body.contains("Mi cuenta") || body.contains("profile")) {
                    return true;
                }
            }

            return false;

        } catch (Exception e) {
            AppLogger.error("Error verificando cookies: " + e.getMessage(), e);
            return false;
        }
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

    private static void actualizarExcelConArchivos(Workbook workbook, Sheet scanSheet, String carpetaImagenes,
            String carpetaVideos, CellStyle headerStyle, CellStyle centeredStyle) {
        try {

            // Verificar si ya existen las columnas, si no, agregarlas
            Row header = scanSheet.getRow(0);
            int colImagenesCarpeta = 7; // Columnas conocidas (ya están definidas arriba)
            int colVideosCarpeta = 8;

            // Aplicar estilo al header (reutilizando el estilo ya creado)
            aplicarStyleFila(header, headerStyle);

            // Procesar cada fila (empezando desde la 1, la 0 es el header)
            int colSku = 4; // Columna SKU
            int colImagenes = 2; // Columna IMAGENES (cantidad en ML)
            int colVideos = 3; // Columna VIDEOS (SI/NO)
            int colConclusionImagenes = 9; // Columna CONCLUSION IMAGENES
            int colConclusionVideos = 10; // Columna CONCLUSION VIDEOS
            int totalFilas = scanSheet.getLastRowNum() + 1;
            AppLogger.info("Procesando " + totalFilas + " productos en carpetas...");

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

                // Buscar imágenes
                int cantidadImagenes = 0;
                if (!carpetaImagenes.isEmpty()) {
                    cantidadImagenes = contarArchivosPorSku(carpetaImagenes, skuNormalizado, IMAGE_EXTENSIONS_SET);
                }

                // Buscar videos (en carpetas nombradas con el SKU)
                int cantidadVideos = 0;
                if (!carpetaVideos.isEmpty()) {
                    cantidadVideos = contarVideosPorSku(carpetaVideos, skuNormalizado);
                }

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
            // Re-ajustar todas las columnas al final para asegurar que todo esté bien
            for (int i = 0; i <= 10; i++) {
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
                int imagenesFaltantes = 6 - cantidadImagenesML;
                if (cantidadImagenesCarpeta < 6) {
                    return "CREAR " + imagenesFaltantes + " " + (imagenesFaltantes == 1 ? "imagen" : "imágenes");
                } else if (cantidadImagenesCarpeta == 6) {
                    return "SUBIR " + imagenesFaltantes + " " + (imagenesFaltantes == 1 ? "imagen" : "imágenes");
                } else {
                    // Solo mostrar "se pueden subir hasta" si hay más de 6 imágenes en carpeta
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

    private static int contarVideosPorSku(String carpetaVideos, String sku) {
        if (sku == null || sku.isEmpty() || carpetaVideos == null || carpetaVideos.isEmpty()) {
            return 0;
        }

        try {
            Path carpetaPath = Paths.get(carpetaVideos);
            if (!Files.exists(carpetaPath) || !Files.isDirectory(carpetaPath)) {
                AppLogger.warn("La carpeta de videos no existe o no es un directorio: " + carpetaVideos);
                return 0;
            }

            int contador = 0;
            String skuUpper = sku.toUpperCase();

            // Buscar carpetas que coincidan con el SKU (primeros 7 caracteres)
            try (Stream<Path> carpetas = Files.list(carpetaPath)) {
                List<Path> carpetasSku = carpetas
                        .filter(Files::isDirectory)
                        .filter(carpeta -> {
                            String nombreCarpeta = carpeta.getFileName().toString().toUpperCase();
                            // Extraer los primeros 7 caracteres del nombre de la carpeta
                            if (nombreCarpeta.length() < 7) {
                                return false;
                            }
                            String primeros7Caracteres = nombreCarpeta.substring(0, 7);
                            // Comparar con el SKU
                            return primeros7Caracteres.equals(skuUpper);
                        })
                        .collect(Collectors.toList());

                // Dentro de cada carpeta encontrada, contar los archivos de video
                for (Path carpetaSku : carpetasSku) {
                    try (Stream<Path> archivos = Files.list(carpetaSku)) {
                        long videosEnCarpeta = archivos
                                .filter(Files::isRegularFile)
                                .filter(archivo -> {
                                    String nombreArchivo = archivo.getFileName().toString().toUpperCase();
                                    // Usar Set para búsqueda más rápida
                                    String extension = nombreArchivo
                                            .substring(Math.max(0, nombreArchivo.lastIndexOf('.')));
                                    return VIDEO_EXTENSIONS_SET.contains(extension.toLowerCase());
                                })
                                .count();
                        contador += (int) videosEnCarpeta;
                    }
                }
            }

            return contador;

        } catch (IOException e) {
            AppLogger.warn("Error al buscar videos en carpeta " + carpetaVideos + ": " + e.getMessage());
            return 0;
        }
    }

    private static int contarArchivosPorSku(String carpeta, String sku, Set<String> extensionesSet) {
        if (sku == null || sku.isEmpty() || carpeta == null || carpeta.isEmpty()) {
            return 0;
        }

        try {
            Path carpetaPath = Paths.get(carpeta);
            if (!Files.exists(carpetaPath) || !Files.isDirectory(carpetaPath)) {
                AppLogger.warn("La carpeta no existe o no es un directorio: " + carpeta);
                return 0;
            }

            int contador = 0;
            String skuUpper = sku.toUpperCase();

            try (Stream<Path> paths = Files.walk(carpetaPath)) {
                contador = (int) paths
                        .filter(Files::isRegularFile)
                        .filter(path -> {
                            String nombreArchivoCompleto = path.getFileName().toString().toUpperCase();

                            // Verificar extensión primero usando Set para búsqueda más rápida
                            int lastDot = nombreArchivoCompleto.lastIndexOf('.');
                            if (lastDot == -1 || lastDot == nombreArchivoCompleto.length() - 1) {
                                return false; // Sin extensión válida
                            }

                            String extension = nombreArchivoCompleto.substring(lastDot).toLowerCase();
                            if (!extensionesSet.contains(extension)) {
                                return false; // Extensión no válida
                            }

                            // Remover la extensión para obtener solo el nombre
                            String nombreSinExtension = nombreArchivoCompleto.substring(0, lastDot);

                            // Extraer los primeros 7 caracteres del nombre del archivo (sin extensión)
                            if (nombreSinExtension.length() < 7) {
                                return false;
                            }

                            String primeros7Caracteres = nombreSinExtension.substring(0, 7);

                            // Comparar con el SKU
                            return primeros7Caracteres.equals(skuUpper);
                        })
                        .count();
            }

            return contador;

        } catch (IOException e) {
            AppLogger.warn("Error al buscar archivos en carpeta " + carpeta + ": " + e.getMessage());
            return 0;
        }
    }

}
