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

El reporte se genera en la hoja "SCAN" del Excel.

## Obtención de Datos por Tipo de Producto

### 1. Productos Tradicionales (Activos)

**Imágenes:**
- Se obtienen del campo `pictures` del objeto `Producto` obtenido mediante la API `getItemByMLA(mla)`
- Cantidad: `producto.pictures.size()`

**Videos:**
- Primero se intenta obtener del campo `videoId` del objeto `Producto`
- Si `videoId` es `null`, se verifica mediante scraping usando `verificarVideo(permalink, cookieHeader)`
- Resultado: "SI" o "NO"

**Performance:**
- Se obtiene mediante la API `getItemPerformanceByMLA(mla)`
- Se extrae:
  - `score`: Puntuación del producto
  - `level_wording`: Nivel del producto (ej: "Profesional", "good")
  - Variables con `status: "PENDING"`: Se crean columnas dinámicas con el nombre de la variable (sin prefijo "UP_") y se concatenan todos los `wordings.title` de las rules con `status: "PENDING"` separados por " | "

### 2. Productos Catálogo

**Imágenes:**
- Se obtienen del campo `pictures` del objeto `Producto` obtenido mediante la API `getItemByMLA(mla)`
- Cantidad: `producto.pictures.size()`

**Videos:**
- Se verifica mediante scraping usando `verificarVideo(permalink, cookieHeader)`
- La API no funciona para productos catálogo, por lo que siempre se usa scraping

**Performance:**
- **No se obtiene** mediante API (la API de performance no funciona para productos catálogo)
- Se establecen valores:
  - `score`: `null` (se muestra "N/A" en Excel)
  - `nivel`: "N/A"
  - Variables: Map vacío (no se crean columnas)

**Nota:** Si el producto catálogo tiene `item_relations`, se usa el MLA del primer `item_relation` para intentar obtener performance, pero si el status no es "active" o es catálogo, se usa scraping.

### 3. Variaciones

**Imágenes:**
- Se obtienen del campo `picture_ids` del JSON de la variación obtenido mediante `getItemNodeByMLAU(userProductId)`
- Cantidad: `picture_ids.size()`

**Videos:**
- Se obtienen del campo `videoId` del producto padre
- Si `videoId` es `null`, se verifica mediante scraping usando `verificarVideo(permalink, cookieHeader)`
- El `permalink` es el mismo del producto padre

**Performance:**
- Se obtiene mediante la API `getItemPerformanceByMLA(mlaPadre)` usando el MLA del producto padre
- Se extrae la misma información que para productos tradicionales
- **Importante:** El performance se obtiene del producto padre, no de la variación individual

### 4. Productos No Activos

**Imágenes:**
- Se obtienen del campo `pictures` del objeto `Producto` obtenido mediante la API `getItemByMLA(mla)`
- Cantidad: `producto.pictures.size()`

**Videos:**
- Se verifica mediante scraping usando `verificarVideo(permalink, cookieHeader)`
- La API no funciona para productos no activos

**Performance:**
- **No se obtiene** mediante API (la API de performance solo funciona para productos con `status: "active"`)
- Se establecen valores:
  - `score`: `null` (se muestra "N/A" en Excel)
  - `nivel`: "N/A"
  - Variables: Map vacío (no se crean columnas)

## Resumen de Métodos de Obtención

| Tipo de Producto | Status | Imágenes | Videos | Performance |
|-----------------|--------|----------|--------|-------------|
| Tradicional | active | API (`pictures`) | API (`videoId`) o Scraping | API (`getItemPerformanceByMLA`) |
| Catálogo | active | API (`pictures`) | Scraping | N/A (no disponible) |
| Variación | active | API (`picture_ids` de variación) | API (`videoId` del padre) o Scraping | API (`getItemPerformanceByMLA` del padre) |
| Cualquiera | no active | API (`pictures`) | Scraping | N/A (no disponible) |

## Columnas Dinámicas de Performance

Las columnas dinámicas se generan basándose en las variables con `status: "PENDING"` encontradas en todos los productos:

- Se extrae la `key` de cada variable con `status: "PENDING"`
- Se quita el prefijo "UP_" si existe (para unificar con variables sin prefijo)
- Se concatenan todos los `wordings.title` de las rules con `status: "PENDING"` separados por " | "
- Cada variable genera una columna con su nombre como header
- Si un producto no tiene una variable específica, la celda queda vacía

## Tecnologías

Java 25, JavaFX, Apache POI, Jackson, Log4j2
