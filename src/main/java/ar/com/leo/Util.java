package ar.com.leo;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;

import java.net.URI;
import java.nio.file.Path;
import java.nio.file.Paths;

public class Util {

    public static String getJarFolder() {
        try {
            URI uri = Util.class.getProtectionDomain()
                    .getCodeSource()
                    .getLocation()
                    .toURI();

            Path path = Paths.get(uri).toAbsolutePath();
            return path.getParent().toString();

        } catch (Exception e) {
            throw new RuntimeException("No se pudo resolver la ruta del JAR", e);
        }
    }

    // Función que detecta si una fila está vacía
    public static boolean isEmptyRow(Row row) {
        if (row == null) {
            return true;
        }
        for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
            final Cell cell = row.getCell(i);
            if (cell != null && cell.getCellType() != CellType.BLANK) {
                // Si es string, comprobar que no sea solo espacios
                if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().isBlank()) {
                    continue; // sigue buscando
                }
                return false; // hay contenido real
            }
        }
        return true;
    }

    // Función que devuelve el valor de una celda
    public static String getCellValue(Cell cell) throws Exception {
        if (cell == null) {
            return "";
        }

        final CellType cellType = cell.getCellType();
        switch (cellType) {
            case STRING:
                return cell.getStringCellValue().trim();

            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    // Fecha
                    return cell.getDateCellValue().toString();
                } else {
                    // Numérico: eliminar ".0" si es entero
                    double num = cell.getNumericCellValue();
                    if (num == Math.floor(num)) {
                        return String.valueOf((long) num);
                    } else {
                        return String.valueOf(num);
                    }
                }

            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());

            case FORMULA:
                switch (cell.getCachedFormulaResultType()) {
                    case STRING:
                        return cell.getStringCellValue().trim();
                    case NUMERIC:
                        if (DateUtil.isCellDateFormatted(cell)) {
                            return cell.getDateCellValue().toString();
                        } else {
                            double num = cell.getNumericCellValue();
                            if (num == Math.floor(num)) {
                                return String.valueOf((long) num);
                            } else {
                                return String.valueOf(num);
                            }
                        }
                    case BOOLEAN:
                        return String.valueOf(cell.getBooleanCellValue());
                    case ERROR:
                        return "0";
                    default:
                        return "";
                }

            case ERROR:
                throw new Exception("Excel - Error en la celda fila: " + (cell.getAddress().getRow() + 1)
                        + " columna: " + (cell.getAddress().getColumn() + 1));

            case BLANK:
            default:
                return "";
        }
    }

}