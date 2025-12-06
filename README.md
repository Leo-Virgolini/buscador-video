# Buscador de Imágenes y Videos de MercadoLibre

Aplicación JavaFX que analiza productos de MercadoLibre y genera un reporte en Excel comparando imágenes y videos publicados con archivos locales.

## Requisitos

- Java 25
- Maven 3.6+
- Archivo Excel con al menos 2 hojas
- Cookies de sesión de MercadoLibre

## Instalación

```bash
mvn clean package
```

## Uso

1. Obtener cookies: F12 → Network → Headers → Request Headers → Copiar valor de "Cookie"
2. Ejecutar: `java -jar target/buscador-video-1.0.jar`
3. Configurar: Seleccionar Excel, carpetas de imágenes/videos y pegar cookies
4. Ejecutar análisis

El reporte se genera en la segunda hoja del Excel.

## Tecnologías

Java 25, JavaFX, Apache POI, Jackson, Log4j2
