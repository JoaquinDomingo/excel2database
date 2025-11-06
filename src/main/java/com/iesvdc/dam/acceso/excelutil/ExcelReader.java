package com.iesvdc.dam.acceso.excelutil;

import java.io.FileInputStream;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.Statement;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.iesvdc.dam.acceso.modelo.FieldType;
import com.iesvdc.dam.acceso.conexion.Conexion;
import com.iesvdc.dam.acceso.modelo.FieldModel;
import com.iesvdc.dam.acceso.modelo.TableModel;
import com.iesvdc.dam.acceso.modelo.WorkbookModel;

public class ExcelReader {
    private Workbook wb;

    private Connection conexion;

    private WorkbookModel wbm;


    public ExcelReader() {
    }

    private static final double EPSILON = 1e-10;

    /**
     * Devuelve un String indicando el tipo de dato de la celda.
     * Puede ser: Entero, Decimal, Texto, Booleano, Fecha, Vacía, Fórmula, Error
     */
    public static FieldType getTipoDato(Cell cell) {
        if (cell == null) {
            return FieldType.UNKNOWN;
        }

        switch (cell.getCellType()) {
            case STRING:
                return FieldType.STRING;

            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return FieldType.DATE;
                } else {
                    double valor = cell.getNumericCellValue();
                    if (Math.abs(valor - Math.floor(valor)) < EPSILON) {
                        return FieldType.INTEGER;
                    } else {
                        return FieldType.DECIMAL;
                    }
                }

            case BOOLEAN:
                return FieldType.BOOLEAN;

            default:
                return FieldType.UNKNOWN;
        }
    }

    public void loadWorbook(String filename){
        try (FileInputStream fis = new FileInputStream(filename)){
            wb = new XSSFWorkbook(fis);
            wbm = new WorkbookModel();

            int nhojas = wb.getNumberOfSheets();
            for (int i = 0; i < nhojas; i++) {
                //empiezo por la primera hoja del libro
                Sheet hojaactual = wb.getSheetAt(i);
                TableModel tabla = new TableModel(hojaactual.getSheetName());
                //de la primera fila tomo las cabeceras
                Row primerafila = hojaactual.getRow(0);
                Row segundafila = hojaactual.getRow(1);
                //obtengo el numero de columnas
                int ncols = primerafila.getLastCellNum();
                for (int j = 0; j < ncols; j++) {
                    FieldModel campo = 
                    new FieldModel(
                        primerafila.getCell(j).getStringCellValue(),
                        getTipoDato(segundafila.getCell(j)));
                    tabla.addField(campo);
                    System.out.println("Añadiendo campo: " + campo.toString());

                //de la segunda fila tomo los tipos de datos
            
                
            }
            wbm.addTable(tabla);
        }
        System.out.println("Tablas");
        System.out.println(wbm.toString());
    } catch (Exception e) {
            System.out.println("Imposible cargar el archivo excel " + e.getLocalizedMessage());
        }
    }

    public String generateDDL(){

        StringBuilder sqlSB = new StringBuilder();
        
        

        for (TableModel tableModel : wbm.getTables()){
            sqlSB.append("CREATE TABLE ");
            sqlSB.append(tableModel.getName());
            sqlSB.append("(");
            int nCampos = tableModel.getFields().size();
            for (FieldModel fieldModel : tableModel.getFields()){
                nCampos--;
                sqlSB.append("`");
                sqlSB.append(fieldModel.getName());
                sqlSB.append("`");
                sqlSB.append(" ");
                sqlSB.append(fieldModel.getType().toString());
                if (nCampos >0) {
                    sqlSB.append(", ");
                }
            } 
           sqlSB.append(");\n");
        }

        return sqlSB.toString();
    }
    public boolean executeDDL(){
        StringBuilder sb;
        boolean resultado = true;

        conexion = Conexion.getConnection();
        if ((conexion== null)) {
            
        }

        for (TableModel tabla : wbm.getTables()) {
            sb = new StringBuilder();
            sb.append("CREATE TABLE `");
            sb.append(tabla.getName());
            sb.append("` (");
            int nCampos = tabla.getFields().size();
            for (FieldModel campo : tabla.getFields()) { 
                nCampos--;
                sb.append("`");
                sb.append(campo.getName());
                sb.append("` ");
                sb.append(getSQLType(campo.getType()));
                if (nCampos > 0) sb.append(", "); 
            }
            sb.append(");");

            try {
                Statement stm = conexion.createStatement();
                stm.execute(sb.toString());
            } catch (Exception e) {
                e.getLocalizedMessage();
            }
           
        }
        return resultado;
    }

    private String getSQLType(FieldType type) {
    switch (type) {
        case INTEGER:
            return "INT";
        case DECIMAL:
            return "DOUBLE";
        case STRING:
            return "VARCHAR(150)";
        case DATE:
            return "DATE";
        case BOOLEAN:
            return "BOOLEAN";
        default:
            return "TEXT";
    }
}
/*
public boolean executeInserts() {
    boolean resultado = true;
    conexion = Conexion.getConnection();
    if (conexion == null) {
        System.err.println("No se pudo establecer la conexión con la base de datos");
        return false;
    }

    try {
        // Recorremos todas las hojas del workbook original (wb)
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            Sheet hoja = wb.getSheetAt(i);
            String nombreTabla = hoja.getSheetName();

            // La primera fila son los nombres de columnas
            Row encabezado = hoja.getRow(0);
            int numColumnas = encabezado.getLastCellNum();

            // Preparamos la sentencia SQL dinámica con parámetros
            StringBuilder sbCampos = new StringBuilder();
            StringBuilder sbParametros = new StringBuilder();

            for (int col = 0; col < numColumnas; col++) {
                String nombreCampo = encabezado.getCell(col).getStringCellValue();
                sbCampos.append("`").append(nombreCampo).append("`");
                sbParametros.append("?");

                if (col < numColumnas - 1) {
                    sbCampos.append(", ");
                    sbParametros.append(", ");
                }
            }

            String sql = "INSERT INTO `" + nombreTabla + "` (" + sbCampos + ") VALUES (" + sbParametros + ")";
            System.out.println("Plantilla SQL preparada: " + sql);

            // Iteramos desde la tercera fila (índice 2)
            for (int fila = 2; fila <= hoja.getLastRowNum(); fila++) {
                Row filaActual = hoja.getRow(fila);
                if (filaActual == null) continue;

                try (PreparedStatement ps = conexion.prepareStatement(sql)) {
                    for (int col = 0; col < numColumnas; col++) {
                        Cell celda = filaActual.getCell(col);
                        asignarValorParametro(ps, col + 1, celda);
                    }
                    ps.executeUpdate();
                } catch (Exception e) {
                    System.err.println("Error insertando fila " + fila + " en " + nombreTabla + ": " + e.getMessage());
                    resultado = false;
                }
            }
        }
    } catch (Exception e) {
        System.err.println("Error general en inserciones: " + e.getMessage());
        resultado = false;
    }

    return resultado;
}

private void asignarValorParametro(PreparedStatement ps, int indice, Cell celda) throws Exception {
    if (celda == null) {
        ps.setNull(indice, java.sql.Types.NULL);
        return;
    }

    switch (celda.getCellType()) {
        case STRING:
            ps.setString(indice, celda.getStringCellValue());
            break;

        case NUMERIC:
            if (DateUtil.isCellDateFormatted(celda)) {
                java.util.Date fecha = celda.getDateCellValue();
                ps.setDate(indice, new java.sql.Date(fecha.getTime()));
            } else {
                double valor = celda.getNumericCellValue();
                if (Math.abs(valor - Math.floor(valor)) < 1e-10)
                    ps.setInt(indice, (int) valor);
                else
                    ps.setDouble(indice, valor);
            }
            break;

        case BOOLEAN:
            ps.setBoolean(indice, celda.getBooleanCellValue());
            break;

        case BLANK:
            ps.setNull(indice, java.sql.Types.NULL);
            break;

        default:
            ps.setNull(indice, java.sql.Types.NULL);
    }
}
 */
}
