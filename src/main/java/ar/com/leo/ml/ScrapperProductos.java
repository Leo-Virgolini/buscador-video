package ar.com.leo.ml;

import ar.com.leo.Util;
import ar.com.leo.ml.model.Producto;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URI;
import java.net.URISyntaxException;
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

public class ScrapperProductos {

    private static final int MAX_THREADS = 5; // controla cuántas requests simultáneas
    private static final int TIMEOUT_SECONDS = 15;
    private static final String BUSQUEDA = "alt=\"clip-icon\"";
    private static final HttpClient httpClient = HttpClient.newBuilder().connectTimeout(Duration.ofSeconds(TIMEOUT_SECONDS)).build();
    // Poner cookie de ML: obtener desde el browser al estar logeado
    private static String COOKIE_HEADER;

    public static void main(String[] args) throws IOException, URISyntaxException, InterruptedException {

//        System.out.println(verificarVideo("https://www.mercadolibre.com.ar/pinza-multiuso-cocina-punta-silicona-30-cm-parrillera-acero/p/MLA26223658?pdp_filters=item_id:MLA1385111157"));

        final String jarDir = Util.getJarFolder();
        System.setProperty("logPath", jarDir + File.separator + "logs");

        Scanner scanner = new Scanner(System.in);
        System.out.println("Abre el navegador -> F12 -> Ir a Mercado Libre (logeado) -> Presionar en Network -> Click en el primer resultado -> Headers -> Abrir Request Headers -> Buscar Cookie -> Copiar todo");
        System.out.print("Pegá tus cookies de Mercado Libre: ");
        COOKIE_HEADER = scanner.nextLine();

        if (!COOKIE_HEADER.isBlank()) {
            System.out.println("Cookies capturadas: " + COOKIE_HEADER);

            final List<Producto> productoList = obtenerDatos();

            ExecutorService executor = Executors.newFixedThreadPool(MAX_THREADS);
            System.out.println("Obteniendo videos...");
            for (Producto producto : productoList) {
                executor.submit(() -> {
                    producto.videoId = verificarVideo(producto.permalink);
                    System.out.println("URL: " + producto.permalink + " - Tiene video: " + producto.videoId);
                });
            }
            executor.shutdown();
            if (!executor.awaitTermination(1, TimeUnit.HOURS)) {
                System.err.println("Timeout del executor.");
            }

            // Ordenamiento
            productoList.sort(Comparator
                    .comparing((Producto p) -> p.status, Comparator.nullsFirst(String::compareTo))
                    .thenComparing(p -> p.id, Comparator.nullsFirst(String::compareTo))
                    .thenComparing(p -> p.pictures == null ? 0 : p.pictures.size())
                    .thenComparing(p -> p.videoId.toString(), Comparator.nullsFirst(String::compareTo))
                    .thenComparing(p -> getSku(p.attributes), Comparator.nullsFirst(String::compareTo))
            );

            // Excel
            String jarFolder;
            try {
                jarFolder = Util.getJarFolder();
            } catch (URISyntaxException e) {
                throw new RuntimeException("No se pudo resolver la ruta del .jar", e);
            }

            Path excelPath = Paths.get(jarFolder, "productos.xlsx");

            // Si existe, lo borro
            if (Files.exists(excelPath)) {
                Files.delete(excelPath);
            }

            // Creo nuevo Excel vacío
            System.out.println("Generando excel...");
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Productos");

            // ==========================
            //  Encabezados
            // ==========================
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("ESTADO");
            header.createCell(1).setCellValue("MLA");
            header.createCell(2).setCellValue("IMAGENES");
            header.createCell(3).setCellValue("VIDEOS");
            header.createCell(4).setCellValue("SKU");
            header.createCell(5).setCellValue("URL");

            // ==========================
            //  Cargar productos
            // ==========================
            int rowNum = 1;
            for (Producto p : productoList) {
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(p.status);
                row.createCell(1).setCellValue(p.id);
                row.createCell(2).setCellValue(p.pictures != null ? p.pictures.size() : 0);
                row.createCell(3).setCellValue(p.videoId.toString());
                row.createCell(4).setCellValue(getSku(p.attributes)); // SKU
                row.createCell(5).setCellValue(p.permalink);
            }

            // Ajuste automático
            for (int i = 0; i <= 4; i++) {
                sheet.autoSizeColumn(i);
            }

            // ==========================
            // Guardar archivo
            // ==========================
            try (FileOutputStream fos = new FileOutputStream(excelPath.toFile())) {
                workbook.write(fos);
            }

            workbook.close();

            System.out.println("Archivo generado: " + excelPath);
        } else {
            System.out.println("Ingresa Cookies válidas.");
        }
    }

    private static String verificarVideo(String url) {
        int status = 0;
        try {
            final HttpRequest request = HttpRequest.newBuilder()
                    .uri(URI.create(url))
                    .timeout(Duration.ofSeconds(TIMEOUT_SECONDS))
                    .header("User-Agent", "Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N)")
                    .header("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7")
                    .header("Accept-Language", "es-AR,es;q=0.9,en;q=0.8")
                    .header("Referer", "https://www.mercadolibre.com.ar/")
                    .header("Cookie", COOKIE_HEADER) // 👈 cookie
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
                                System.out.println("URL vieja: " + url + " - URL actualizada: " + newUrl);
                                return verificarVideo(newUrl);
                            }
                        }
                    }
                    break;
                case 404:
                case 410:
                    return "NO EXISTE";
                case 403:
                case 424:
                    System.out.println("Too many requests.");
                    Thread.sleep(60000);
                    return verificarVideo(url);
                case 500:
                case 502:
                case 503:
                case 504:
                    System.out.println("Internal server error.");
                    Thread.sleep(5000);
                    return verificarVideo(url);
                default:
                    return "STATUS: " + status;
            }
        } catch (Exception e) {
            System.err.println("Error en url: " + url + " - status: " + status + " -> " + e.getMessage());
            return "ERROR: " + e.getMessage();
        }

        return "ERROR: " + status;
    }

    public static List<Producto> obtenerDatos() throws InterruptedException, IOException {

        MercadoLibreAPI.inicializar();

        final String userId = MercadoLibreAPI.getUserId();
        System.out.println("User ID: " + userId);

        System.out.println("Obteniendo MLAs de todos los productos...");
        final List<String> productos = MercadoLibreAPI.obtenerTodosLosItemsId(userId);
        System.out.println("Total de Productos: " + productos.size());

        final List<Producto> productoList = Collections.synchronizedList(new ArrayList<>());

        System.out.println("Obteniendo estado, sku, cantidad de imágenes y urls de los productos...");
        ExecutorService executor = Executors.newFixedThreadPool(10);
        for (String mla : productos) {
            executor.submit(() -> {
                Producto producto = MercadoLibreAPI.getItemByMLA(mla);
                if (producto != null) {
                    productoList.add(producto);
                }
            });
        }
        executor.shutdown();
        if (!executor.awaitTermination(1, TimeUnit.HOURS)) {
            System.err.println("Timeout del executor.");
        }

        return productoList;
    }

    public static String getSku(List<Producto.Attribute> attributes) {
        for (Producto.Attribute a : attributes) {
            if ("SELLER_SKU".equals(a.id)) {
                if (a.valueName != null) {
                    return a.valueName.length() >= 7
                            ? a.valueName.substring(0, 7)
                            : a.valueName;
                }
            }
        }
        return null;
    }

}
