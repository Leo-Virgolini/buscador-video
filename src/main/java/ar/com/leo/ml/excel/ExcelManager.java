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

        FileInputStream fis = new FileInputStream(excelFile);
        Workbook workbook = new XSSFWorkbook(fis);

        // Verificar que tenga al menos 2 hojas
        if (workbook.getNumberOfSheets() < 2) {
            workbook.close();
            fis.close();
            throw new IllegalArgumentException("El archivo Excel debe tener al menos 2 hojas. " +
                    "Hojas encontradas: " + workbook.getNumberOfSheets());
        }

        return workbook;
    }

    /**
     * Obtiene la hoja de escaneo (segunda hoja, índice 1).
     */
    public static Sheet obtenerHojaEscaneo(Workbook workbook) {
        return workbook.getSheetAt(1); // 2da hoja
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
