package com.mycompany.priceupdate;

public enum Headers {
    CIKKSZAM ("Cikkszám"),
    EUR_LISTA ("Listaár (EUR)"),
    HU_LISTA ("Listaár (Ft)"),
    AGRAM_KONF_LISTA ("Agram konfek. lista ár (HRK)"),
    AGRAM_KONF_TR ("Agram konfek. transf. ár (EUR)"),
    AGRAM_LISTA ("Agram lista ár (HRK)"),
    AGRAM_TR ("Agram transfer ár (EUR)"),
    OKOVI_KONF_LISTA ("Okovi konfek. lista ár (SRD)"),
    OKOVI_KONF_TR ("Okovi konfek. transf. ár (EUR)"),
    OKOVI_LISTA ("Okovi lista ár (SRD)"),
    OKOVI_TR ("Okovi transfer ár (EUR)"),
    EUR_KONF ("Konfekcionált ár (EUR)"),
    HU_KONF ("Konfekcionált ár (Ft)"),
    FUVAR ("K00001;0;Fuvar %"),
    VAM ("K00002;0;Vám %"),
    ENGEDMENY ("K00003;0;Engedmény %"),
    EGYEB ("K00004;0;Egyéb %"),
    HULLADEK ("K00005;0;Hulladék / selejt %"),
    SZELESSEG ("K00006;0;Szélesség %"),
    FIXKTG ("K00007;1;Fix költség"),
    SZELESSEG2 ("Szélesség");
    
    private String cat;

    private Headers(String cat) {
        this.cat = cat;
    }

    public String getCat() {
        return cat;
    }
}
