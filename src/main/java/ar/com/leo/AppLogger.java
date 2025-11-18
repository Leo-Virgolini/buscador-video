package ar.com.leo;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.util.function.Consumer;

public class AppLogger {

    private static final Logger logger = LogManager.getLogger(AppLogger.class);

    private static Consumer<String> uiLogger; // callback para updateMessage()

    public static void info(String message) {
        logger.info(message);
    }

    public static void warn(String message) {
        logger.warn(message);
    }

    public static void error(String message, Throwable t) {
        logger.error(message, t);
    }

}
