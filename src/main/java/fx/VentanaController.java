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
import ar.com.leo.ml.ProductReportService;
import javafx.geometry.Insets;
import javafx.scene.layout.HBox;
import javafx.scene.layout.VBox;
import javafx.application.Platform;
import java.util.Optional;
import java.util.concurrent.CountDownLatch;

public class VentanaController implements Initializable {

    @FXML
    private TextField ubicacionExcel;
    @FXML
    private TextField ubicacionCarpetaImagenes;
    @FXML
    private TextField ubicacionCarpetaVideos;

    @FXML
    private TextArea cookiesTextArea;
    @FXML
    private TextField requestsPorSegundo;

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
        fileChooser.getExtensionFilters()
                .add(new FileChooser.ExtensionFilter("Archivo XLSX", "*.xlsx"));

        // Obtener último archivo guardado y su carpeta padre
        File lastFile = new File(prefs.get("ubicacionExcel", ""));
        File initialDir =
                (lastFile.exists() && lastFile.getParentFile() != null) ? lastFile.getParentFile()
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
        File carpeta =
                seleccionarCarpeta("ubicacionCarpetaVideos", "Elige carpeta de videos/clips");
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
            logTextArea
                    .appendText("Haz click en 'Buscar Carpeta' junto a 'Carpeta de Imágenes'.\n");
            return;
        }

        // Validar carpeta de videos (requerido)
        if (carpetaVideos == null || !carpetaVideos.isDirectory()) {
            logTextArea
                    .appendText("❌ Error: Debes seleccionar una carpeta de videos/clips válida.\n");
            logTextArea.appendText(
                    "Haz click en 'Buscar Carpeta' junto a 'Carpeta de Videos/Clips'.\n");
            return;
        }

        // Validar requests por segundo
        if (requestsPorSegundo.getText().trim().isBlank()) {
            logTextArea
                    .appendText("❌ Error: Debes ingresar la cantidad de requests por segundo.\n");
            return;
        }
        double requestsPorSeg = Double.parseDouble(requestsPorSegundo.getText().trim());

        ProductReportService service = new ProductReportService(excelFile, carpetaImagenes,
                carpetaVideos, cookies, requestsPorSeg);

        service.messageProperty().addListener((obs, old, nuevo) -> {
            if (nuevo != null && !nuevo.isBlank()) {
                logTextArea.appendText(nuevo + "\n");
            }
        });

        service.setOnRunning(e -> {
            buscarButton.setDisable(true);
            progressIndicator.setVisible(true);
            logTextArea.setStyle("-fx-text-fill: darkblue;");
            String fechaHora =
                    LocalDateTime.now().format(DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm:ss"));
            AppLogger.info("[" + fechaHora + "] Iniciando proceso...");
        });

        service.setOnSucceeded(e -> {
            logTextArea.setStyle("-fx-text-fill: darkgreen;");
            String fechaHora =
                    LocalDateTime.now().format(DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm:ss"));
            AppLogger.info("[" + fechaHora + "] Proceso finalizado exitosamente.");
            buscarButton.setDisable(false);
            progressIndicator.setVisible(false);
        });

        service.setOnFailed(e -> {
            logTextArea.setStyle("-fx-text-fill: firebrick;");
            AppLogger.error("Error: " + service.getException().getLocalizedMessage(),
                    service.getException());
            buscarButton.setDisable(false);
            progressIndicator.setVisible(false);
        });

        service.start();
    }

    /**
     * Muestra un diálogo para solicitar el código de autorización de MercadoLibre.
     * 
     * @param authURL La URL de autorización que el usuario debe abrir
     * @return El código de autorización ingresado por el usuario, o null si canceló
     */
    public static String mostrarDialogoAutorizacion(String authURL) {
        // Crear TextInputDialog
        TextInputDialog dialog = new TextInputDialog();
        dialog.setTitle("Autorización de MercadoLibre");
        dialog.setHeaderText("Autorización requerida");

        // Configurar el diálogo para que no se cierre automáticamente si el campo está vacío
        // Esto permite que el usuario cancele explícitamente
        dialog.getDialogPane().lookupButton(ButtonType.OK).setDisable(false);

        // Crear contenido personalizado con URL seleccionable
        VBox vbox = new VBox(10);
        vbox.setPadding(new Insets(10));

        // Label con instrucciones
        Label label1 = new Label("1. Abra esta URL en su navegador:");
        vbox.getChildren().add(label1);

        // HBox para URL y botón copiar
        HBox urlBox = new HBox(5);
        TextField urlField = new TextField(authURL);
        urlField.setEditable(false);
        urlField.setStyle("-fx-background-color: #f0f0f0;");

        Button copyButton = new Button("Copiar");
        copyButton.setOnAction(e -> {
            try {
                java.awt.Toolkit.getDefaultToolkit()
                        .getSystemClipboard()
                        .setContents(new java.awt.datatransfer.StringSelection(authURL), null);
                AppLogger.info("URL copiada al portapapeles");
            } catch (Exception ex) {
                AppLogger.warn("No se pudo copiar al portapapeles: " + ex.getMessage());
            }
        });

        urlBox.getChildren().addAll(urlField, copyButton);
        vbox.getChildren().add(urlBox);

        // Label con instrucción para el código
        Label label2 = new Label("2. Después de autorizar, pegue el código aquí:");
        vbox.getChildren().add(label2);

        // Configurar el diálogo para usar el contenido personalizado
        dialog.setGraphic(vbox);
        dialog.getEditor().setPromptText("Ingrese el código de autorización...");

        // Mostrar diálogo y obtener resultado
        Optional<String> result = dialog.showAndWait();

        // Verificar resultado
        // Si el usuario cerró el diálogo sin ingresar código, result.isPresent() será false
        // Si el usuario hizo clic en OK pero dejó el campo vacío, result.isPresent() será true pero el string estará vacío
        if (!result.isPresent()) {
            // Usuario cerró el diálogo (X o Cancelar)
            return null;
        }

        String code = result.get();
        if (code != null && !code.trim().isEmpty()) {
            return code.trim();
        }

        // Usuario hizo clic en OK pero no ingresó código
        return null;
    }

    /**
     * Solicita el código de autorización usando JavaFX.
     * Maneja la ejecución en el hilo correcto de JavaFX.
     * 
     * @param authURL La URL de autorización
     * @return El código de autorización ingresado por el usuario
     * @throws RuntimeException Si el usuario canceló o hubo un error
     */
    public static String pedirCodeConJavaFX(String authURL) {
        // Verificar si ya estamos en el hilo de JavaFX
        if (Platform.isFxApplicationThread()) {
            // Ya estamos en el hilo de JavaFX, mostrar el diálogo directamente
            String code = mostrarDialogoAutorizacion(authURL);
            if (code == null || code.trim().isEmpty()) {
                throw new RuntimeException("El usuario canceló la autorización");
            }
            return code.trim();
        } else {
            // No estamos en el hilo de JavaFX, usar Platform.runLater
            final String[] codeResult = new String[1];
            final RuntimeException[] exceptionResult = new RuntimeException[1];
            final CountDownLatch latch = new CountDownLatch(1);

            Platform.runLater(() -> {
                try {
                    codeResult[0] = mostrarDialogoAutorizacion(authURL);
                } catch (RuntimeException e) {
                    exceptionResult[0] = e;
                } catch (Exception e) {
                    exceptionResult[0] = new RuntimeException("Error al mostrar diálogo", e);
                } finally {
                    latch.countDown();
                }
            });

            try {
                // Esperar a que se complete el diálogo
                latch.await();

                if (exceptionResult[0] != null) {
                    throw exceptionResult[0];
                }

                if (codeResult[0] == null || codeResult[0].trim().isEmpty()) {
                    throw new RuntimeException("El usuario canceló la autorización");
                }

                return codeResult[0].trim();
            } catch (InterruptedException e) {
                Thread.currentThread().interrupt();
                throw new RuntimeException("Interrupción mientras se esperaba el diálogo", e);
            }
        }
    }

}
