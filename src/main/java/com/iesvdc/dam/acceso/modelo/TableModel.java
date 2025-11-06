package com.iesvdc.dam.acceso.modelo;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

import org.apache.commons.lang3.builder.EqualsBuilder;

/**
 * El modelo que almacena información de una tabla y su lista de campos.
 */
public class TableModel {
    private final String name;
    private final List<FieldModel> fields;

    public TableModel() {
        this.name = "";
        fields = new ArrayList<FieldModel>();
    }

    public TableModel(String name) {
    this.name = name;
    this.fields = new ArrayList<FieldModel>();
    }


    public String getName() {
        return this.name;
    }


    public List<FieldModel> getFields() {
        return this.fields;
    }

    public boolean addField(FieldModel field){
        return fields.add(field);
    }

    @Override
    public boolean equals(Object o) {
      return EqualsBuilder.reflectionEquals(this, o);
    }

    @Override
    public int hashCode() {
        return Objects.hash(name, fields);
    }

    @Override
    public String toString() {
        return "{" +
            " name='" + getName() + "'" +
            ", fields='" + getFields() + "'" +
            "}";
    }

}
