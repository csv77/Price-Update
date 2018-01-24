package com.mycompany.priceupdate;

public enum SchemaCat {
    FUVAR (0, "K00001"),
    VAM (0, "K00002"),
    ENGEDMENY (0, "K00003"),
    EGYEB (0, "K00004"),
    HULLADEK (0, "K00005"),
    SZELESSEG (0, "K00006"),
    FIXKTG (1, "K00007");
    
    private final int type;
    private final String code;
    
    SchemaCat(int type, String code) {
        this.type = type;
        this.code = code;
    }

    public int getType() {
        return type;
    }

    public String getCode() {
        return code;
    }
}
