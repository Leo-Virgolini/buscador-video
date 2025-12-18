package ar.com.leo.ml.excel;

import ar.com.leo.ml.model.ProductoData;
import org.apache.poi.ss.usermodel.*;

import java.util.List;

/**
 * Maneja la escritura de datos de productos al archivo Excel.
 */
public class ExcelWriter {

    /**
     * Limpia las filas existentes en la hoja (excepto el encabezado).
     */
    public static void limpiarDatosExistentes(Sheet sheet) {
        int lastRowNum = sheet.getLastRowNum();
        if (lastRowNum > 0) {
            // Eliminar filas desde la última hasta la primera (excepto fila 0 que es el encabezado)
            // Usar removeRow en lugar de shiftRows para evitar problemas con muchas filas
            for (int i = lastRowNum; i > 0; i--) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    sheet.removeRow(row);
                }
            }
        }
    }

    /**
     * Crea los encabezados de las columnas en la hoja.
     */
    public static void crearEncabezados(Sheet sheet, CellStyle headerStyle) {
        Row header = sheet.getRow(0);
        if (header == null) {
            header = sheet.createRow(0);
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

        ExcelStyleManager.aplicarStyleFila(header, headerStyle);
    }

    /**
     * Escribe los datos de los productos en la hoja de Excel.
     */
    public static void escribirProductos(Sheet sheet, List<ProductoData> productoList,
            Workbook workbook, CellStyle centeredStyle) {
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

            // CORREGIR: títulos de keys con status PENDING (última columna), si es null poner "N/A"
            row.createCell(13).setCellValue(p.corregir != null ? p.corregir : "N/A");

            // Aplicar estilo a todas las celdas excepto SCORE y NIVEL (que ya tienen su estilo)
            ExcelStyleManager.aplicarStyleFilaExcluyendo(row, centeredStyle, 11, 12);
        }
    }

    /**
     * Ajusta el ancho de todas las columnas automáticamente.
     */
    public static void ajustarAnchoColumnas(Sheet sheet) {
        for (int i = 0; i <= 13; i++) {
            sheet.autoSizeColumn(i);
        }
    }
}
