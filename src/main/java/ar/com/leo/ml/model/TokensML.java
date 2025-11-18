package ar.com.leo.ml.model;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.annotation.JsonProperty;

@JsonIgnoreProperties(ignoreUnknown = true) // Ignora cualquier campo extra
public class TokensML {

    @JsonProperty("access_token")
    public String accessToken;

    @JsonProperty("refresh_token")
    public String refreshToken;

    @JsonProperty("expires_in")
    public long expiresIn; // en segundos

    public long issuedAt;  // timestamp en milisegundos

    public boolean isExpired() {
        long now = System.currentTimeMillis();
        long elapsed = (now - issuedAt) / 1000;
        return elapsed >= expiresIn;
    }

}
