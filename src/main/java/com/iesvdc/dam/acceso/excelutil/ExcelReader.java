package com.iesvdc.dam.acceso.excelutil;

import java.io.FileInputStream;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.SQLException;
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
        boolean resultado = true;

        conexion = Conexion.getConnection();
        if ((conexion== null)) {
            
        }
    
        try (Statement stm = conexion.createStatement()){
            String ddl = generateDDL();
            stm.execute(ddl);
        } catch (Exception e) {
            e.getLocalizedMessage();
        }
        return resultado;
    }
    

public boolean insertDDL(){
    conexion = Conexion.getConnection();

    if ((conexion== null)) {
            
    }
    boolean insertado = true;

    try {
        for (int i = 0; i < wb.getNumberOfSheets(); i++){
            Sheet hoja = wb.getSheetAt(i);
            String tabla = hoja.getSheetName();

            Row cabecera = hoja.getRow(0);
            int columns  = cabecera.getLastCellNum();

            StringBuilder sb = new StringBuilder();
            sb.append("INSERT INTO `").append(tabla).append("` (");

            for (int j = 0; j < columns; j++) {
                String campo = cabecera.getCell(j).getStringCellValue();
                sb.append("`").append(campo).append("`");
                if (j < columns - 1) {
                    sb.append(", ");
                }
            }
            sb.append(") VALUES (");
            
            for (int k = 0; k < columns; k++) {
                sb.append("?");
                if (k < columns - 1) {
                    sb.append(", ");
                }
            }
            
            sb.append(")");
            String sentencia = sb.toString();

            for (int j = 2; j < hoja.getPhysicalNumberOfRows(); j++) {
                Row filactual = hoja.getRow(j);
                if (filactual == null) {
                    System.out.println("Esta fila " + j + " es nula");
                    continue;
                }

                try (PreparedStatement stm = conexion.prepareStatement(sentencia)){
                    for (int k = 0; k < columns; k++) {
                        Cell celda = filactual.getCell(k);
                        asignarValorSQL(stm, k + 1, celda);
                    }
                    stm.executeUpdate();
                } catch (Exception e) {
                    e.getLocalizedMessage();
                    insertado = false;
                }
            }

        }
    } catch (Exception e) {
        e.getLocalizedMessage();
        insertado = false;
    }
    return insertado;
}

private void asignarValorSQL(PreparedStatement ps, int indice, Cell celda) throws SQLException {
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
                if (Math.abs(valor - Math.floor(valor)) < EPSILON)
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

}
