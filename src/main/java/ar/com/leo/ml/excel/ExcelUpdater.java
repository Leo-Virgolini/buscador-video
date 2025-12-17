package ar.com.leo.ml.excel;

import ar.com.leo.AppLogger;
import ar.com.leo.Util;
import org.apache.poi.ss.usermodel.*;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;
import java.util.stream.Stream;

/**
 * Actualiza el Excel con información de archivos encontrados en las carpetas.
 */
public class ExcelUpdater {

    // Sets para búsqueda más rápida de extensiones
    private static final Set<String> IMAGE_EXTENSIONS_SET = Set.of(".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp");
    private static final Set<String> VIDEO_EXTENSIONS_SET = Set.of(".mp4", ".avi", ".mov", ".mkv", ".wmv", ".flv",
            ".webm", ".m4v");

    /**
     * Actualiza el Excel con información de archivos encontrados en las carpetas.
     */
    public static void actualizarExcelConArchivos(Workbook workbook, Sheet scanSheet, String carpetaImagenes,
            String carpetaVideos, CellStyle headerStyle, CellStyle centeredStyle) {
        try {
            // Verificar si ya existen las columnas, si no, agregarlas
            Row header = scanSheet.getRow(0);
            int colImagenesCarpeta = 7; // Columna IMAGENES EN CARPETA
            int colVideosCarpeta = 8; // Columna VIDEOS EN CARPETA

            // Aplicar estilo al header (reutilizando el estilo ya creado)
            ExcelStyleManager.aplicarStyleFila(header, headerStyle);

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
                String conclusionImagenes = ConclusionGenerator.generarConclusionImagenes(row, colImagenes, colImagenesCarpeta);
                String conclusionVideos = ConclusionGenerator.generarConclusionVideos(row, colVideos, colVideosCarpeta);

                // Escribir conclusión de imágenes
                Cell cellConclusionImagenes = row.getCell(colConclusionImagenes);
                if (cellConclusionImagenes == null) {
                    cellConclusionImagenes = row.createCell(colConclusionImagenes);
                }
                cellConclusionImagenes.setCellValue(conclusionImagenes);
                cellConclusionImagenes.setCellStyle(ExcelStyleManager.obtenerEstiloConclusion(workbook, conclusionImagenes));

                // Escribir conclusión de videos
                Cell cellConclusionVideos = row.getCell(colConclusionVideos);
                if (cellConclusionVideos == null) {
                    cellConclusionVideos = row.createCell(colConclusionVideos);
                }
                cellConclusionVideos.setCellValue(conclusionVideos);
                cellConclusionVideos.setCellStyle(ExcelStyleManager.obtenerEstiloConclusion(workbook, conclusionVideos));

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
