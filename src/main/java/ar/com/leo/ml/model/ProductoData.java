package ar.com.leo.ml.model;

/**
 * Clase para almacenar los datos que van al Excel.
 * Representa tanto productos principales como variaciones.
 */
public class ProductoData {
    public String status;
    public String mla; // ID del producto (MLA o MLAU para variaciones)
    public int cantidadImagenes;
    public String tieneVideo; // "SI" o "NO"
    public String sku;
    public String permalink;
    public String tipoPublicacion; // "CATALOGO" o "TRADICIONAL"
    public boolean esVariacion; // Indica si es una variación
    public String userProductId; // Solo para variaciones

    public ProductoData(Producto producto, String sku) {
        this.status = producto.status;
        this.mla = producto.id;
        this.cantidadImagenes = producto.pictures != null ? producto.pictures.size() : 0;
        this.tieneVideo = producto.videoId != null ? producto.videoId.toString() : "NO";
        this.sku = sku;
        this.permalink = producto.permalink;
        this.tipoPublicacion = producto.catalogListing ? "CATALOGO" : "TRADICIONAL";
        this.esVariacion = false;
        this.userProductId = null;
    }

    /**
     * Constructor para variaciones
     */
    public ProductoData(Producto productoPadre, String userProductId, String skuVariacion) {
        this.status = productoPadre.status;
        this.mla = productoPadre.id;
        this.cantidadImagenes = productoPadre.pictures != null ? productoPadre.pictures.size() : 0;
        this.tieneVideo = productoPadre.videoId != null ? productoPadre.videoId.toString() : "NO";
        this.sku = skuVariacion;
        this.permalink = productoPadre.permalink;
        this.tipoPublicacion = productoPadre.catalogListing ? "CATALOGO" : "TRADICIONAL";
        this.esVariacion = true;
        this.userProductId = userProductId;
    }
}
