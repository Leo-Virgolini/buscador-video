package ar.com.leo.ml.excel;

import ar.com.leo.Util;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

/**
 * Genera conclusiones sobre imágenes y videos basándose en los datos del Excel.
 */
public class ConclusionGenerator {

    /**
     * Genera la conclusión sobre las imágenes de un producto.
     * Compara la cantidad de imágenes en ML con las disponibles en la carpeta.
     */
    public static String generarConclusionImagenes(Row row, int colImagenes, int colImagenesCarpeta) {
        try {
            // Leer cantidad de imágenes en ML (optimizado: leer directamente como número si es posible)
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
            return "ERROR";
        }
    }

    /**
     * Genera la conclusión sobre los videos de un producto.
     * Compara si tiene video en ML con los videos disponibles en la carpeta.
     */
    public static String generarConclusionVideos(Row row, int colVideos, int colVideosCarpeta) {
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
            return "ERROR";
        }
    }
}
