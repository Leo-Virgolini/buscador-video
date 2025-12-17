package ar.com.leo.ml.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.HashMap;
import java.util.Map;

/**
 * Maneja la creación y caché de estilos para celdas de Excel.
 * Evita crear estilos duplicados para mejorar el rendimiento.
 */
public class ExcelStyleManager {
    
    // Caché de estilos para evitar crear estilos duplicados
    private static final Map<String, CellStyle> estiloCache = new HashMap<>();

    /**
     * Crea un estilo para encabezados con fondo gris, texto en negrita y centrado.
     */
    public static CellStyle crearHeaderStyle(Workbook workbook) {
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

    /**
     * Crea un estilo centrado con bordes.
     */
    public static CellStyle crearCenteredStyle(Workbook workbook) {
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

    /**
     * Crea un estilo centrado con color de fondo.
     */
    public static CellStyle crearCenteredStyleWithColor(Workbook workbook, IndexedColors color) {
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

    /**
     * Obtiene el estilo de celda según el nivel del producto.
     * Profesional → verde, Estándar → amarillo, Básica → rojo
     */
    public static CellStyle obtenerEstiloPorNivel(Workbook workbook, String nivel) {
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

    /**
     * Obtiene el estilo según la conclusión.
     * OK → verde, CREAR → rojo, SUBIR → amarillo
     */
    public static CellStyle obtenerEstiloConclusion(Workbook workbook, String conclusion) {
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

    /**
     * Obtiene un estilo del caché o lo crea si no existe.
     */
    private static CellStyle obtenerEstiloCached(Workbook workbook, String key,
            java.util.function.Supplier<CellStyle> styleCreator) {
        // Usar el workbook como parte de la clave para evitar conflictos entre
        // diferentes workbooks
        String fullKey = workbook.hashCode() + "_" + key;
        return estiloCache.computeIfAbsent(fullKey, k -> styleCreator.get());
    }

    /**
     * Limpia el caché de estilos para un workbook específico.
     */
    public static void limpiarCacheEstilos(Workbook workbook) {
        // Limpiar estilos del workbook actual del caché
        int workbookHash = workbook.hashCode();
        estiloCache.entrySet().removeIf(entry -> entry.getKey().startsWith(workbookHash + "_"));
    }

    /**
     * Aplica un estilo a todas las celdas de una fila.
     */
    public static void aplicarStyleFila(Row row, CellStyle style) {
        for (int i = 0; i < row.getLastCellNum(); i++) {
            if (row.getCell(i) != null) {
                row.getCell(i).setCellStyle(style);
            }
        }
    }

    /**
     * Aplica estilo a todas las celdas de una fila excepto las columnas especificadas.
     */
    public static void aplicarStyleFilaExcluyendo(Row row, CellStyle style, int... columnasExcluidas) {
        java.util.Set<Integer> excluidas = new java.util.HashSet<>();
        for (int col : columnasExcluidas) {
            excluidas.add(col);
        }

        for (int i = 0; i < row.getLastCellNum(); i++) {
            if (!excluidas.contains(i) && row.getCell(i) != null) {
                row.getCell(i).setCellStyle(style);
            }
        }
    }
}
