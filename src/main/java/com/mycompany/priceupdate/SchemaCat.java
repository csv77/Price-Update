package com.mycompany.priceupdate;

public enum SchemaCat {
    FUVAR (0, 1),
    VAM (0, 1),
    EGYEB (0, 1),
    JUTALEK (0, 1),
    HULLADEK (0, 1),
    SZELESSEG (0, 1),
    FIXKTG (1, 1);
    
    private final int type;
    private final int code;
    
    SchemaCat(int type, int code) {
        this.type = type;
        this.code = code;
    }

    public int getType() {
        return type;
    }

    public int getCode() {
        return code;
    }
}
