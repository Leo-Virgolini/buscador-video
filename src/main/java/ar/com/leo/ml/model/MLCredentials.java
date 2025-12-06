package ar.com.leo.ml.model;

import com.fasterxml.jackson.annotation.JsonProperty;

public class MLCredentials {

    @JsonProperty("CLIENT_ID")
    public String clientId;

    @JsonProperty("CLIENT_SECRET")
    public String clientSecret;

    @JsonProperty("REDIRECT_URI")
    public String redirectUri;

}