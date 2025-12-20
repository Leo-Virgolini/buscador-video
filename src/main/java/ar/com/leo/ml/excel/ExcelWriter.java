package ar.com.leo.ml.excel;

import ar.com.leo.ml.model.ProductoData;
import org.apache.poi.ss.usermodel.*;

import java.util.*;

/**
 * Maneja la escritura de datos de productos al archivo Excel.
 */
public class ExcelWriter {

    /**
     * Limpia las filas existentes en la hoja, incluyendo el encabezado.
     * Los headers se regeneran dinámicamente basados en los datos de los productos.
     */
    public static void limpiarDatosExistentes(Sheet sheet) {
        int lastRowNum = sheet.getLastRowNum();
        if (lastRowNum >= 0) {
            // Eliminar todas las filas, incluyendo el encabezado (fila 0)
            // Usar removeRow en lugar de shiftRows para evitar problemas con muchas filas
            for (int i = lastRowNum; i >= 0; i--) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    sheet.removeRow(row);
                }
            }
        }
    }

    /**
     * Recolecta todas las variable keys de todos los productos para crear columnas dinámicas.
     */
    private static List<String> recolectarVariableKeys(List<ProductoData> productoList) {
        Set<String> keysSet = new LinkedHashSet<>(); // LinkedHashSet para mantener orden
        for (ProductoData p : productoList) {
            if (p.corregir != null && !p.corregir.isEmpty()) {
                keysSet.addAll(p.corregir.keySet());
            }
        }
        return new ArrayList<>(keysSet);
    }

    /**
     * Crea los encabezados de las columnas en la hoja.
     * Incluye columnas fijas y columnas dinámicas basadas en variable keys.
     */
    public static void crearEncabezados(Sheet sheet, CellStyle headerStyle,
            List<ProductoData> productoList) {
        Row header = sheet.getRow(0);
        if (header == null) {
            header = sheet.createRow(0);
        }

        Workbook workbook = sheet.getWorkbook();
        CellStyle headerStyleRojo = ExcelStyleManager.crearHeaderStyleRojo(workbook);

        // Columnas fijas
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

        // Aplicar estilo normal a columnas fijas (0-12)
        for (int i = 0; i <= 12; i++) {
            Cell cell = header.getCell(i);
            if (cell != null) {
                cell.setCellStyle(headerStyle);
            }
        }

        // Columnas dinámicas basadas en variable keys (con texto rojo)
        List<String> variableKeys = recolectarVariableKeys(productoList);
        int colIndex = 13; // Empezar después de NIVEL
        for (String variableKey : variableKeys) {
            Cell cell = header.createCell(colIndex++);
            cell.setCellValue(variableKey);
            cell.setCellStyle(headerStyleRojo);
        }
    }

    /**
     * Escribe los datos de los productos en la hoja de Excel.
     */
    public static void escribirProductos(Sheet sheet, List<ProductoData> productoList,
            Workbook workbook, CellStyle centeredStyle) {
        // Recolectar todas las variable keys para mantener el mismo orden que en los headers
        List<String> variableKeys = recolectarVariableKeys(productoList);

        int rowNum = 1;
        for (ProductoData p : productoList) {
            Row row = sheet.createRow(rowNum++);

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

            // SCORE: si es null, poner "N/A", si no, poner el número
            Cell cellScore = row.createCell(11);
            if (p.score != null) {
                cellScore.setCellValue(p.score);
            } else {
                cellScore.setCellValue("N/A");
            }
            // Aplicar estilo según el nivel
            cellScore.setCellStyle(ExcelStyleManager.obtenerEstiloPorNivel(workbook, p.nivel));

            // NIVEL: si es null, poner "N/A"
            Cell cellNivel = row.createCell(12);
            cellNivel.setCellValue(p.nivel != null ? p.nivel : "N/A");
            // Aplicar estilo según el nivel
            cellNivel.setCellStyle(ExcelStyleManager.obtenerEstiloPorNivel(workbook, p.nivel));

            // Columnas dinámicas: escribir el valor para cada variable key
            int colIndex = 13; // Empezar después de NIVEL
            for (String variableKey : variableKeys) {
                String value = "";
                if (p.corregir != null && p.corregir.containsKey(variableKey)) {
                    value = p.corregir.get(variableKey);
                }
                row.createCell(colIndex++).setCellValue(value);
            }

            // Aplicar estilo a todas las celdas excepto SCORE y NIVEL (que ya tienen su estilo)
            ExcelStyleManager.aplicarStyleFilaExcluyendo(row, centeredStyle, 11, 12);
        }
    }

    /**
     * Ajusta el ancho de todas las columnas automáticamente.
     * Ajusta las columnas fijas (0-12) más las columnas dinámicas.
     */
    public static void ajustarAnchoColumnas(Sheet sheet) {
        Row header = sheet.getRow(0);
        if (header != null) {
            int lastCellNum = header.getLastCellNum();
            for (int i = 0; i < lastCellNum; i++) {
                sheet.autoSizeColumn(i);
            }
        }
    }
}
