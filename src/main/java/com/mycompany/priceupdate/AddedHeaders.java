package com.mycompany.priceupdate;

public enum AddedHeaders {
    BESZAREUR ("Beszár Eur"),
    BESZARHUF ("Beszár HUF"),
    ARRESEUR ("Árrés EUR"),
    ARRESHUF ("Árrés HUF"),
    ARRESAGRAM ("Árrés Agram"),
    ARRESOKOVI ("Árrés Okovi"),
    EUR_LISTA ("Új listaár (EUR)"),
    HU_LISTA ("Új listaár (Ft)"),
    AGRAM_KONF_LISTA ("Új Agram konfek. lista ár (HRK)"),
    AGRAM_KONF_TR ("Új Agram konfek. transf. ár (EUR)"),
    AGRAM_LISTA ("Új Agram lista ár (HRK)"),
    AGRAM_TR ("Új Agram transfer ár (EUR)"),
    OKOVI_KONF_LISTA ("Új Okovi konfek. lista ár (SRD)"),
    OKOVI_KONF_TR ("Új Okovi konfek. transf. ár (EUR)"),
    OKOVI_LISTA ("Új Okovi lista ár (SRD)"),
    OKOVI_TR ("Új Okovi transfer ár (EUR)"),
    EUR_KONF ("Új konfekcionált ár (EUR)"),
    HU_KONF ("Új konfekcionált ár (Ft)");
    
    private String cat;

    private AddedHeaders(String cat) {
        this.cat = cat;
    }

    public String getCat() {
        return cat;
    }
}
