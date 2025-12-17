package fx;

import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.*;
import javafx.stage.FileChooser;
import javafx.stage.DirectoryChooser;

import java.io.File;
import java.net.URL;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ResourceBundle;
import java.util.prefs.Preferences;

import ar.com.leo.AppLogger;
import ar.com.leo.ml.ScrapperService;

public class VentanaController implements Initializable {

    @FXML
    private TextField ubicacionExcel;
    @FXML
    private TextField ubicacionCarpetaImagenes;
    @FXML
    private TextField ubicacionCarpetaVideos;

    @FXML
    private TextArea logTextArea;
    @FXML
    private ProgressIndicator progressIndicator;
    @FXML
    private Button buscarButton;

    private File excelFile; // Archivo Excel de salida
    private File carpetaImagenes; // Carpeta de imágenes
    private File carpetaVideos; // Carpeta de videos

    public void initialize(URL url, ResourceBundle rb) {
        loadPreferences();
        Main.stage.setOnCloseRequest(event -> {
            savePreferences();
        });
    }

    private void loadPreferences() {
        final Preferences prefs = Preferences.userRoot().node("buscadorVideo");

        String pathExcel = prefs.get("ubicacionExcel", "");
        excelFile = new File(pathExcel);
        if (excelFile.isFile()) {
            ubicacionExcel.setText(excelFile.getAbsolutePath());
        } else {
            excelFile = null;
        }

        String pathImagenes = prefs.get("ubicacionCarpetaImagenes", "");
        carpetaImagenes = new File(pathImagenes);
        if (carpetaImagenes.isDirectory()) {
            ubicacionCarpetaImagenes.setText(carpetaImagenes.getAbsolutePath());
        } else {
            carpetaImagenes = null;
        }

        String pathVideos = prefs.get("ubicacionCarpetaVideos", "");
        carpetaVideos = new File(pathVideos);
        if (carpetaVideos.isDirectory()) {
            ubicacionCarpetaVideos.setText(carpetaVideos.getAbsolutePath());
        } else {
            carpetaVideos = null;
        }
    }

    private void savePreferences() {
        final Preferences prefs = Preferences.userRoot().node("buscadorVideo");

        prefs.put("ubicacionExcel", ubicacionExcel.getText());
        prefs.put("ubicacionCarpetaImagenes", ubicacionCarpetaImagenes.getText());
        prefs.put("ubicacionCarpetaVideos", ubicacionCarpetaVideos.getText());
    }

    @FXML
    public void buscarExcel() {
        Preferences prefs = Preferences.userRoot().node("buscadorVideo");

        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Elige archivo Excel de salida");
        fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Archivo XLSX", "*.xlsx"));

        // Obtener último archivo guardado y su carpeta padre
        File lastFile = new File(prefs.get("ubicacionExcel", ""));
        File initialDir = (lastFile.exists() && lastFile.getParentFile() != null) ? lastFile.getParentFile()
                : new File(System.getProperty("user.dir"));

        fileChooser.setInitialDirectory(initialDir);

        excelFile = fileChooser.showOpenDialog(Main.stage);

        if (excelFile != null) {
            ubicacionExcel.setText(excelFile.getAbsolutePath());
            prefs.put("ubicacionExcel", excelFile.getAbsolutePath());
        } else {
            ubicacionExcel.clear();
            excelFile = null;
        }
    }

    @FXML
    public void buscarCarpetaImagenes() {
        File carpeta = seleccionarCarpeta("ubicacionCarpetaImagenes", "Elige carpeta de imágenes");
        if (carpeta != null) {
            carpetaImagenes = carpeta;
            ubicacionCarpetaImagenes.setText(carpeta.getAbsolutePath());
        } else {
            ubicacionCarpetaImagenes.clear();
            carpetaImagenes = null;
        }
    }

    @FXML
    public void buscarCarpetaVideos() {
        File carpeta = seleccionarCarpeta("ubicacionCarpetaVideos", "Elige carpeta de videos/clips");
        if (carpeta != null) {
            carpetaVideos = carpeta;
            ubicacionCarpetaVideos.setText(carpeta.getAbsolutePath());
        } else {
            ubicacionCarpetaVideos.clear();
            carpetaVideos = null;
        }
    }

    private File seleccionarCarpeta(String preferenceKey, String titulo) {
        Preferences prefs = Preferences.userRoot().node("buscadorVideo");
        DirectoryChooser directoryChooser = new DirectoryChooser();
        directoryChooser.setTitle(titulo);

        File lastDir = new File(prefs.get(preferenceKey, ""));
        File initialDir = (lastDir.exists() && lastDir.isDirectory()) ? lastDir
                : new File(System.getProperty("user.dir"));

        directoryChooser.setInitialDirectory(initialDir);

        File carpeta = directoryChooser.showDialog(Main.stage);
        if (carpeta != null) {
            prefs.put(preferenceKey, carpeta.getAbsolutePath());
        }
        return carpeta;
    }

    @FXML
    public void buscarImagenesYVideos() {
        logTextArea.clear();
        logTextArea.setStyle("-fx-text-fill: firebrick;");

        // Validar Excel (requerido)
        if (excelFile == null || !excelFile.isFile()) {
            logTextArea.appendText("❌ Error: Debes seleccionar un archivo Excel válido.\n");
            logTextArea.appendText("Haz click en 'Buscar Excel' para seleccionar el archivo.\n");
            return;
        }

        // Validar carpeta de imágenes (requerido)
        if (carpetaImagenes == null || !carpetaImagenes.isDirectory()) {
            logTextArea.appendText("❌ Error: Debes seleccionar una carpeta de imágenes válida.\n");
            logTextArea.appendText("Haz click en 'Buscar Carpeta' junto a 'Carpeta de Imágenes'.\n");
            return;
        }

        // Validar carpeta de videos (requerido)
        if (carpetaVideos == null || !carpetaVideos.isDirectory()) {
            logTextArea.appendText("❌ Error: Debes seleccionar una carpeta de videos/clips válida.\n");
            logTextArea.appendText("Haz click en 'Buscar Carpeta' junto a 'Carpeta de Videos/Clips'.\n");
            return;
        }

        ScrapperService service = new ScrapperService(excelFile, carpetaImagenes, carpetaVideos);

        service.messageProperty().addListener((obs, old, nuevo) -> {
            if (nuevo != null && !nuevo.isBlank()) {
                logTextArea.appendText(nuevo + "\n");
            }
        });

        service.setOnRunning(e -> {
            buscarButton.setDisable(true);
            progressIndicator.setVisible(true);
            logTextArea.setStyle("-fx-text-fill: darkblue;");
            String fechaHora = LocalDateTime.now().format(DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm:ss"));
            AppLogger.info("[" + fechaHora + "] Iniciando proceso...");
        });

        service.setOnSucceeded(e -> {
            logTextArea.setStyle("-fx-text-fill: darkgreen;");
            String fechaHora = LocalDateTime.now().format(DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm:ss"));
            AppLogger.info("[" + fechaHora + "] Proceso finalizado exitosamente.");
            buscarButton.setDisable(false);
            progressIndicator.setVisible(false);
        });

        service.setOnFailed(e -> {
            logTextArea.setStyle("-fx-text-fill: firebrick;");
            AppLogger.error("Error: " + service.getException().getLocalizedMessage(), service.getException());
            buscarButton.setDisable(false);
            progressIndicator.setVisible(false);
        });

        service.start();
    }

}