package com.iesvdc.dam.acceso.modelo;

/**
 * Tipos de datos posibles para los campos
 */
public enum FieldType {
    INTEGER,
    DECIMAL,
    STRING,
    DATE,
    BOOLEAN,
    UNKNOWN;

    /**
     * Determina si el tipo representa un valor numérico.
     * @return true si es INTEGER o DECIMAL.
     */
    public boolean isNumeric() {
        return this == INTEGER || this == DECIMAL;
    }
}