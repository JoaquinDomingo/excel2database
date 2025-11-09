package com.iesvdc.dam.acceso.excelutil;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.Statement;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.iesvdc.dam.acceso.conexion.Conexion;

/*
 * Clase ExcelWriter
 * 
 * Mediante esta clase y sus métodos, obtendremos los datos de la base de datos
 * y posteriormente los pasaremos a un nuevo excel distinto desde el que pasamos los datos 
 * a la base de datos desde nuestra clase ExcelReader.
 * 
 * 
 */
public class ExcelWriter {

    private Workbook wb;
    private Sheet hoja;
    private Connection conexion;

    /*
     * Constructor de nuestra clase el cual nos permitira instanciarla 
     * en la clase principal Excel2database
     * 
     */
    public ExcelWriter() {

    }

    /*
     * Metodo público con el cual abriremos nuestros archivo excel 
     * al cual le insertaremos los datos extraidos de nuestra base de datos
     * 
     * En este usaremos el tipo FileInputStream, ya que tenemos un archivo
     * creado con las cabeceras de los datos.
     * 
     * Abrimos nuestro excel mediante XSSFWorkbook(fis) y nos vamos a la hoja
     * en la que queramos trabajar, en este caso sería la primera
     */
    public void loadWorkbook(String filename) {
        try (FileInputStream fis = new FileInputStream(filename)) {
            wb = new XSSFWorkbook(fis);
            hoja = wb.getSheetAt(0);
        } catch (Exception e) {
            System.out.println("Error al abrir el archivo Excel: " + e.getMessage());
        }
    }

    /*
     * Con este método público, una vez que ya hayamos abierto el excel en el 
     * método loadWorkbook, procederemos a escribir los datos extraidos de la base de datos
     * en la hoja seleccionada.
     * 
     * Para esto utilizaremos una consulta SQL mediante un Statement y un ResultSet
     * que sera el cual ejecute ese statement. Para poder saber los tipos de datos que 
     * tenemos que introducir y ademas el número de columnas, usaremos el metodo
     * getMetaData() nuestro ResultSet.
     * 
     * Posteriormente, siempore y cuando nuestro resultSet siga teniendo datos, crearemos
     * las filas donde insertaremos los datos extraidos.
     * 
     * Mediante rs.getObject(i) obtendremos el valor de cada columna y dependiendo del tipo 
     * de dato que sea mediante un instaceof, los iremos insertando en la celda a la que 
     * le corresponda ese dato.
     */
    public void writeExcel() {
        if (wb == null || hoja == null) {
            System.out.println("Workbook o hoja no cargada.");
            return;
        }

        conexion = Conexion.getConnection();
        if (conexion == null){
            System.out.println("No se ha podido establecer la conexión");
            return;
        } 

        try (Statement stmt = conexion.createStatement();
             ResultSet rs = stmt.executeQuery("SELECT * FROM personas")) {

            int filaNum = 1;
            /*
             * Aquí obtenemos el número de columnas que tiene nuestra tabla
             * mediante el metodo getMetaData(), del cual obtenemos los metadatos
             * de la base de datos, y luego con el getColumnCount() obtenemos el
             * número de columnas que tiene la tabla de la cual hagamos el SELECT.  
             */
            int columnas = rs.getMetaData().getColumnCount();

            while (rs.next()) {
                Row fila = hoja.createRow(filaNum++);
                for (int i = 1; i <= columnas; i++) {
                    Cell celda = fila.createCell(i - 1);
                    /*
                     * Aquí comprobamos el tipo de dato que es cada columna, esto
                     * lo hacemos directamente con la clase Object, ya que al no saber
                     * que datos nos vamos a encontrar a la hora de bajarlos
                     * de la base no sabemos que tipo puede ser.
                     * 
                     * Posteriormente mediante varios else if y los instaceof, comprobamos 
                     * el tipo de datos que es.
                     */
                    Object valor = rs.getObject(i);
                    if (valor == null) {
                        celda.setBlank();
                    } else if (valor instanceof Integer){
                        celda.setCellValue((Integer) valor);
                    } else if (valor instanceof Double) {
                        celda.setCellValue((Double) valor);
                    } else if (valor instanceof Boolean) {
                        celda.setCellValue((Boolean) valor);
                    } else if (valor instanceof java.sql.Date) {
                        celda.setCellValue((java.util.Date) valor);
                    } else {
                        celda.setCellValue(valor.toString());
                    }
                }
            }

        } catch (Exception e) {
            System.out.println("Error al escribir datos: " + e.getMessage());
        }
    }

    /**
     * Este método funciona al contrario que el método loadWorkbook.
     * Ya que en este en lugar de cargar el excel, lo que haremos en esta 
     * ocasion sera escribir el excel con los datos obtenidos y guardar 
     * el archivo.
     * 
     * En lugar de usar FileInputStream, usaremos FileOutputStream ya que en 
     * esta ocasion, no estaremos recibiendo datos de la base de datos, sino 
     * que vamos a escribir los datos ya recibidos.
     * 
     */
    public void saveWorkbook(String filename) {
        try (FileOutputStream fos = new FileOutputStream(filename)) {
            wb.write(fos);
            wb.close();
            System.out.println("Archivo guardado correctamente: " + filename);
        } catch (Exception e) {
            System.out.println("Error al guardar el archivo Excel: " + e.getMessage());
        }
    }
}
