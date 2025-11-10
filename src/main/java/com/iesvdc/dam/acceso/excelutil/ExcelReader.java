package com.iesvdc.dam.acceso.excelutil;

import java.io.FileInputStream;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Date;

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
import java.sql.Types;
/*
 *  Clase ExcelReader
 * 
 *  En esta clase se encuentran varios métodos con los cuales nos ayudaremos para poder
 *  transferir los datos desde un archivo excel a una base de datos.
 * 
 *  En esta nos creamos un workboox, un modelo de workbook y  abrimos la conexion a nuetras base
 *  de datos a la cual le insertameros los datos extraidos desde un archivo .xlsx 
 * 
 * 
 */
public class ExcelReader {

    private Workbook wb;

    private Connection conexion;

    private WorkbookModel wbm;

    /*
     * Constructor de nuestra clase el cual nos permitira instanciarla 
     * en la clase principal Excel2database
     */
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
    /**
     * 
     * Método público con el cual abriremos nuestro archivo excel
     * y mediante un FileInputStream, cargaremos los datos que hay tiene dentro este
     * excel. Posteriormente, crearemos el WorkbookModel que contendrá
     * la estructura de datos que necesitaremos para poder generar la sentencia SQL de
     * inserción de datos 
     * 
     */
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

    /**
     * Método público con el cual mediantes los datos recogidos gracias al método 
     * loadWorkbook, generarelos la sentencia SQL necesaria para crear las tablas
     * en nuestra base de datos
     * 
     * Crearemos la sentencia mediante un StringBuilder, y recorriendo el archivo, 
     * obtedremos lo necesario para la sentencia, gracias a fieldModel.getName() y
     * fieldModel.getType().toString().
     * 
     * Por ultimo, gracias al métoto toString() de StringBuilder, parsearemos este último
     * a String.
     * 
     * @return Nos devuelve un String, el cual será la sentecia SQL
     */
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
    /**
     * Método público con el cual ejecutaremos la sentencia SQL
     * generada en el método generateDDL(), con este método
     * crearemos las tablas en la base de datos
     * 
     * @return Nos devuelve un booleano, en el caso en el que la sentencia
     * se haya ejecutado de manera correcta, obtendremos un true, y un false
     * si ha habido algún error a la hora de la ejecucion.
     */
    public boolean executeDDL() {
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
    
    /**
     * Método público por el cual, insertaremos los datos 
     * extraidos del excel en la base de datos. Para ello, utilizaremos un bucle 
     * for para recorrer las hojas del excel, haremos una serie de bucles for anidados
     * para recorrer completamente el excel. Después mediante los PreparedStatement, 
     * crearemos la sentencia SQL de inserción, y mediante otro bucle for, 
     * obtendremos el numero de columnas con datos que tiene la hoja actual. Y en el caso en 
     * el que esta celda sea nula, mostraremos un mensaje por consola sobre la celda, y contnuaremos
     * con las siguientes filas. 
     * 
     * Por último, mediante un método auxiliar, le asignaremos el valor SQL a cada celda dependiendo
     * del tipo de dato que sea cada una.
     * 
     * @return Nos devuelve un true en el caso en el que la inserción haya sido ejecutada de manera correcta, y un false
     * si no se ha podido realizar la inserción correctamente.
     */
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
                            /*
                             * Aqui utilizamos el método auxiliar para asignar el valor SQL
                             * a cada celda dependiendo del tipo de dato que sea.
                             */
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

/**
 * 
 * Método privado auxiliar con el cual, dependiendo del tipo de dato que obtengamos del excel,
 * le asignaremos el valor SQL que le corresponda a cada datos. Esto lo realizaremos mediante
 * un switch del tipo de dato de la celda. 
 * 
 * Si la celda es de tipo STRING, asignaremos el valor mediante el método setString().
 * Si es de tipo NUMERIC, comprobaremos si es una fecha o un número, y dependiendo de esto, 
 * le asignaremos el valor mediante setDate() o setInt()/setDouble(), siendo este último, 
 * dependiendo de si es un número entero o decimal. 
 * 
 * Si nuestra celda es de tipo BOOLEAN, usaremos el método setBoolean().
 * 
 * Y si la celda estuviera vacía, usaremos el método setNull(). Al igual que en el caso
 * por defecto del switch.
 * 
 * Para ello le pasamos como parámetros el PreparedStatement, el índice de la celda y la propia celda.
 * 
 */
private void asignarValorSQL(PreparedStatement ps, int indice, Cell celda) throws SQLException {
    if (celda == null) {
        ps.setNull(indice, Types.NULL);
        return;
    }

    switch (celda.getCellType()) {
        case STRING:
            ps.setString(indice, celda.getStringCellValue());
            break;

        case NUMERIC:
            if (DateUtil.isCellDateFormatted(celda)) {
                Date fecha = celda.getDateCellValue();
                /*
                 * Aqui convertimos el java.util.Date a java.sql.Date pero al ya 
                 * tener importada la clase Date de java.util, tenemos que poner 
                 * el nombre completo de la clase java.sql.Date, ya que nos da conflicto
                 * al intentar exportar ambas.
                 */
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
            ps.setNull(indice, Types.NULL);
            break;

        default:
            ps.setNull(indice, Types.NULL);
    }
}

}
