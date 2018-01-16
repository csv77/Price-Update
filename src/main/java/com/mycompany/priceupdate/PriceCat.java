package com.mycompany.priceupdate;

public enum PriceCat {
    HU_LISTA ("Ft", 1),
    EUR_LISTA ("EUR", 2),
    AGRAM_TR ("EUR", 4),
    AGRAM_LISTA ("HRK", 10),
    OKOVI_TR ("EUR", 5),
    OKOVI_LISTA ("SRD", 12),
    HU_KONF ("Ft", 6),
    EUR_KONF ("EUR", 7),
    AGRAM_KONF_TR ("EUR", 8),
    AGRAM_KONF_LISTA ("HRK", 11),
    OKOVI_KONF_TR ("EUR", 9),
    OKOVI_KONF_LISTA ("SRD", 13);
    
    private final String currency;
    private final int code;
    
    PriceCat(String currency, int code) {
        this.currency = currency;
        this.code = code;
    }

    public String getCurrency() {
        return currency;
    }

    public int getCode() {
        return code;
    }
    
}
