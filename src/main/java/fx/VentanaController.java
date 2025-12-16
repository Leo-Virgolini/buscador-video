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
    private TextField requestsPorSegundo;

    @FXML
    private TextArea logTextArea;
    @FXML
    private TextArea cookiesTextArea;
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

        String cookies = prefs.get("cookies", "");
        if (!cookies.isBlank()) {
            cookiesTextArea.setText(cookies);
        }

        String requestsPorSeg = prefs.get("requestsPorSegundo", "5");
        requestsPorSegundo.setText(requestsPorSeg);
    }

    private void savePreferences() {
        final Preferences prefs = Preferences.userRoot().node("buscadorVideo");

        prefs.put("ubicacionExcel", ubicacionExcel.getText());
        prefs.put("ubicacionCarpetaImagenes", ubicacionCarpetaImagenes.getText());
        prefs.put("ubicacionCarpetaVideos", ubicacionCarpetaVideos.getText());
        prefs.put("cookies", cookiesTextArea.getText());
        prefs.put("requestsPorSegundo", requestsPorSegundo.getText());
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

        // Validar cookies
        String cookies = cookiesTextArea.getText().trim();
        if (cookies.isBlank()) {
            logTextArea.appendText("❌ Error: Debes ingresar las cookies de MercadoLibre.\n\n");
            logTextArea.appendText("Instrucciones para obtener las cookies:\n");
            logTextArea.appendText("1. Abre el navegador -> Presiona F12\n");
            logTextArea.appendText("2. Ve a Mercado Libre (debes estar logueado)\n");
            logTextArea.appendText("3. Presiona la pestaña 'Network' o 'Red'\n");
            logTextArea.appendText("4. Recarga la página o navega por el sitio\n");
            logTextArea.appendText("5. Click en el primer resultado de la lista\n");
            logTextArea.appendText("6. Ve a 'Headers' -> 'Request Headers'\n");
            logTextArea.appendText("7. Busca 'Cookie' y copia TODO su valor\n");
            logTextArea.appendText("8. Pega el valor completo en el campo de cookies\n");
            return;
        }

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

        // Validar y obtener requests por segundo
        double requestsPorSeg = 5.0; // Valor por defecto
        try {
            String requestsText = requestsPorSegundo.getText().trim();
            if (!requestsText.isBlank()) {
                requestsPorSeg = Double.parseDouble(requestsText);
                if (requestsPorSeg <= 0 || requestsPorSeg > 20) {
                    logTextArea.appendText(
                            "⚠️ Advertencia: Requests por segundo debe estar entre 0.1 y 20. Usando valor por defecto: 5\n");
                    requestsPorSeg = 5.0;
                }
            }
        } catch (NumberFormatException e) {
            logTextArea.appendText(
                    "⚠️ Advertencia: Valor inválido en 'Requests por segundo'. Usando valor por defecto: 5\n");
            requestsPorSeg = 5.0;
        }

        ScrapperService service = new ScrapperService(excelFile, carpetaImagenes, carpetaVideos, cookies,
                requestsPorSeg);

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