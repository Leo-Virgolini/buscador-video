package ar.com.leo;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.util.concurrent.BlockingQueue;
import java.util.concurrent.LinkedBlockingQueue;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.function.Consumer;

public class AppLogger {

    private static final Logger logger = LogManager.getLogger(AppLogger.class);

    private static volatile Consumer<String> uiLogger; // callback para updateMessage()
    private static final BlockingQueue<String> messageQueue = new LinkedBlockingQueue<>();
    private static final AtomicBoolean processorRunning = new AtomicBoolean(false);
    private static Thread processorThread;

    public static void setUiLogger(Consumer<String> uiLogger) {
        AppLogger.uiLogger = uiLogger;
        startMessageProcessor();
    }

    private static synchronized void startMessageProcessor() {
        if (processorRunning.get() || uiLogger == null) {
            return;
        }

        processorRunning.set(true);
        processorThread = new Thread(() -> {
            while (processorRunning.get() || !messageQueue.isEmpty()) {
                try {
                    // Tomar mensaje de la cola (bloquea hasta que haya uno)
                    String message = messageQueue.take();
                    Consumer<String> logger = uiLogger;
                    if (logger != null && message != null) {
                        logger.accept(message);
                    }
                    // Pequeña pausa para evitar saturar el UI
                    Thread.sleep(10);
                } catch (InterruptedException e) {
                    Thread.currentThread().interrupt();
                    break;
                } catch (Exception e) {
                    // Ignorar errores en el procesador
                }
            }
        });
        processorThread.setDaemon(true);
        processorThread.setName("AppLogger-Processor");
        processorThread.start();
    }

    public static void shutdown() {
        processorRunning.set(false);
        if (processorThread != null) {
            processorThread.interrupt();
        }
    }

    public static void info(String message) {
        logger.info(message);
        ui(message);
    }

    public static void warn(String message) {
        logger.warn(message);
        ui(message);
    }

    public static void error(String message, Throwable t) {
        logger.error(message, t);
        ui(message);
    }

    private static void ui(String message) {
        if (message == null || message.isEmpty()) {
            return;
        }

        // Agregar mensaje a la cola (thread-safe)
        try {
            messageQueue.put(message);
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            // Si falla la cola, intentar llamar directamente como fallback
            Consumer<String> logger = uiLogger;
            if (logger != null) {
                logger.accept(message);
            }
        }
    }

}
