package com.iesvdc.dam.acceso.modelo;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

import org.apache.commons.lang3.builder.EqualsBuilder;


/**
 * El modelo que almacena el libro o lista de tablas.
 */
public class WorkbookModel {
    private final List<TableModel> tables;

    public WorkbookModel() {
        tables = new ArrayList<TableModel>();
    }

    public WorkbookModel(List<TableModel> tables) {
        this.tables = tables;
    }

    public List<TableModel> getTables() {
        return this.tables;
    }

    public boolean addTable(TableModel table){
        return tables.add(table);
    }


    
    @Override
    public boolean equals(Object o) {
      return EqualsBuilder.reflectionEquals(this, o);
    }

    @Override
    public int hashCode() {
        return Objects.hash(tables);
    }

    @Override
    public String toString() {
        return "{" +
            " tables='" + getTables() + "'" +
            "}";
    }

}
