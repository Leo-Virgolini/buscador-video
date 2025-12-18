package ar.com.leo.ml.excel;

import ar.com.leo.AppLogger;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.RandomAccessFile;
import java.nio.file.Files;
import java.nio.file.Path;

/**
 * Maneja operaciones generales con archivos Excel: validación, apertura y guardado.
 */
public class ExcelManager {

    /**
     * Verifica que el archivo Excel no esté en uso.
     * Lanza excepción si el archivo está abierto en otro programa.
     */
    public static void verificarArchivoDisponible(File excelFile) throws Exception {
        try (@SuppressWarnings("unused")
        RandomAccessFile raf = new RandomAccessFile(excelFile, "rw")) {
            // Si se puede abrir, está disponible
        } catch (Exception ex) {
            throw new IllegalStateException("El excel está en uso. Cerralo antes de continuar.");
        }
    }

    /**
     * Valida que el archivo Excel exista.
     */
    public static void validarArchivoExiste(File excelFile) {
        Path excelPath = excelFile.toPath();
        if (!Files.exists(excelPath)) {
            throw new IllegalArgumentException("El archivo Excel no existe: " + excelPath +
                    ". Por favor, crea el archivo Excel antes de ejecutar el proceso.");
        }
    }

    /**
     * Abre un archivo Excel y retorna el Workbook.
     * Configura el límite de detección de Zip bomb.
     */
    public static Workbook abrirWorkbook(File excelFile) throws Exception {
        AppLogger.info("Abriendo archivo Excel...");

        // Configurar límite de detección de Zip bomb para archivos con alta compresión
        ZipSecureFile.setMinInflateRatio(0.001);

        // Usar try-with-resources para asegurar que el FileInputStream se cierre
        // incluso si falla la construcción del Workbook
        try (FileInputStream fis = new FileInputStream(excelFile)) {
            Workbook workbook = new XSSFWorkbook(fis);

            // Verificar que el workbook sea válido (debe tener al menos 1 hoja)
            if (workbook.getNumberOfSheets() < 1) {
                workbook.close();
                throw new IllegalArgumentException("El archivo Excel no tiene hojas válidas.");
            }

            // El workbook ahora tiene el contenido en memoria, el stream se cerrará automáticamente
            return workbook;
        }
    }

    /**
     * Obtiene la hoja de escaneo llamada "SCAN".
     * Si no existe, la crea.
     */
    public static Sheet obtenerHojaEscaneo(Workbook workbook) {
        Sheet sheet = workbook.getSheet("SCAN");
        if (sheet == null) {
            // Si no existe la hoja "SCAN", crearla
            AppLogger.info("La hoja 'SCAN' no existe, creándola...");
            sheet = workbook.createSheet("SCAN");
        }
        return sheet;
    }

    /**
     * Guarda el Workbook en el archivo Excel.
     */
    public static void guardarWorkbook(Workbook workbook, File excelFile) throws Exception {
        AppLogger.info("Guardando archivo Excel...");
        try (FileOutputStream fos = new FileOutputStream(excelFile)) {
            workbook.write(fos);
            fos.flush();
        } catch (Exception ex) {
            AppLogger.error("Error al guardar Excel: " + ex.getMessage(), ex);
            throw ex;
        } finally {
            // Limpiar caché de estilos después de usar el workbook
            ExcelStyleManager.limpiarCacheEstilos(workbook);
        }
    }
}
