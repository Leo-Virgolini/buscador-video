package ar.com.leo.ml.model;

import tools.jackson.databind.JsonNode;
import tools.jackson.databind.ObjectMapper;

import java.util.HashMap;
import java.util.Map;

/**
 * Clase para almacenar los datos que van al Excel.
 * Representa tanto productos principales como variaciones.
 */
public class ProductoData {
    private static final ObjectMapper mapper = new ObjectMapper();

    public String status;
    public String mla; // ID del producto (MLA o MLAU para variaciones)
    public int cantidadImagenes;
    public String tieneVideo; // "SI" o "NO"
    public String sku;
    public String permalink;
    public String tipoPublicacion; // "CATALOGO" o "TRADICIONAL"
    public boolean esVariacion; // Indica si es una variación
    public String userProductId; // Solo para variaciones
    public String mlaParaPerformance; // MLA a usar para performance/video (puede ser de item_relations si es
                                      // catálogo)
    public Integer score; // Score del producto (del performance)
    public String nivel; // Nivel del producto (level_wording del performance)
    public Map<String, String> corregir; // Map<variableKey, title> para variables con status PENDING

    /**
     * Extrae el MLA de item_relations si es catálogo, sino devuelve el MLA normal
     */
    private static String obtenerMlaParaPerformance(Producto producto) {
        // Si es catálogo y tiene item_relations, usar el MLA de item_relations
        if (Boolean.TRUE.equals(producto.catalogListing) && producto.itemRelations != null
                && !producto.itemRelations.isEmpty()) {
            try {
                // Convertir el primer elemento de itemRelations a JsonNode
                Object firstRelation = producto.itemRelations.get(0);
                JsonNode relationNode = mapper.valueToTree(firstRelation);
                JsonNode idNode = relationNode.path("id");
                if (!idNode.isNull()) {
                    String mlaFromRelation = idNode.asString("");
                    if (mlaFromRelation != null && !mlaFromRelation.isEmpty()) {
                        return mlaFromRelation;
                    }
                }
            } catch (Exception e) {
                // Si hay error, usar el MLA normal
            }
        }
        // Si no es catálogo o no tiene item_relations, usar el MLA normal
        return producto.id;
    }

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
        this.mlaParaPerformance = obtenerMlaParaPerformance(producto);
        this.score = null;
        this.nivel = null;
        this.corregir = new HashMap<>();
    }

    /**
     * Constructor para variaciones
     * 
     * @param productoPadre             El producto padre
     * @param userProductId             El user_product_id de la variación
     * @param skuVariacion              El SKU de la variación
     * @param cantidadImagenesVariacion La cantidad de imágenes de la variación (de
     *                                  picture_ids)
     */
    public ProductoData(Producto productoPadre, String userProductId, String skuVariacion,
            int cantidadImagenesVariacion) {
        this.status = productoPadre.status;
        this.mla = productoPadre.id;
        this.cantidadImagenes = cantidadImagenesVariacion; // Usar las imágenes de la variación
        this.tieneVideo = productoPadre.videoId != null ? productoPadre.videoId.toString() : "NO";
        this.sku = skuVariacion;
        this.permalink = productoPadre.permalink;
        this.tipoPublicacion = productoPadre.catalogListing ? "CATALOGO" : "TRADICIONAL";
        this.esVariacion = true;
        this.userProductId = userProductId;
        // Para variaciones, usar el MLA del padre para performance
        this.mlaParaPerformance = productoPadre.id;
        this.score = null;
        this.nivel = null;
        this.corregir = new HashMap<>();
    }
}
