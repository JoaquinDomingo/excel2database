package com.iesvdc.dam.acceso.modelo;
import java.util.Objects;

import org.apache.commons.lang3.builder.EqualsBuilder;

/**
 * El modelo que almacena información de un campo y sus propiedades.
 */

public class FieldModel {
    private String name;
    private FieldType type;

    public FieldModel() {
        this.name = "";
        this.type = FieldType.UNKNOWN; 
    }

    public FieldModel(String name, FieldType type) {
        this.name = name;
        this.type = type;
    }

        public FieldModel(String name) {
        this.name = name;
        this.type = FieldType.UNKNOWN;
    }
    public String getName() {
        return this.name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public FieldType getType() {
        return this.type;
    }

    public void setType(FieldType type) {
        this.type = type;
    }

    public FieldModel name(String name) {
        setName(name);
        return this;
    }

    public FieldModel type(FieldType type) {
        setType(type);
        return this;
    }

    @Override
    public boolean equals(Object o) {
      return EqualsBuilder.reflectionEquals(this, o);
    }

    @Override
    public int hashCode() {
        return Objects.hash(name, type);
    }

    @Override
    public String toString() {
        return "{" +
            " name='" + getName() + "'" +
            ", type='" + getType() + "'" +
            "}";
    }
    
}
