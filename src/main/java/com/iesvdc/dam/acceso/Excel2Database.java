package com.iesvdc.dam.acceso;

import java.sql.Connection;
import java.util.Properties;

import com.iesvdc.dam.acceso.conexion.Conexion;
import com.iesvdc.dam.acceso.conexion.Config;
import com.iesvdc.dam.acceso.excelutil.ExcelReader;
import com.iesvdc.dam.acceso.excelutil.ExcelWriter;
import com.iesvdc.dam.acceso.modelo.WorkbookModel;

/**
 * Este programa genérico en java (proyecto Maven) es un ejercicio 
 * simple que vuelca un libro Excel (xlsx) a una base de datos (MySQL) 
 * y viceversa. El programa lee la configuración de la base de datos 
 * de un fichero "properties" de Java y luego, con apache POI, leo 
 * las hojas, el nombre de cada hoja será el nombre de las tablas, 
 * la primera fila de cada hoja será el nombre de los atributos de 
 * cada tabla (hoja) y para saber el tipo de dato, tendré que 
 * preguntar a la segunda fila qué tipo de dato tiene. 
 * 
 * Procesamos el fichero Excel y creamos una estructura de datos 
 * con la información siguiente: La estructura principal es el libro, 
 * que contiene una lista de tablas y cada tabla contiene tuplas 
 * nombre del campo y tipo de dato.
 *
 */
public class Excel2Database 
{
    public static void main( String[] args )
    {               
        /**
         * Cargamos las propiedades nuestro .properties
         * en una variable de tipo Properties mediante
         * el metodo getProperties de la clase Config.
         */ 
        Properties props = Config.getProperties("config.properties");
        System.out.println("La acción es: " + props.getProperty("action"));


        /*
         * Mediante este condicional if - else if, nuestro programa realizará dos acciones
         * distintas. Si en nuestro archivo config.properties, si la propiedad action del
         * archivo es load, nuestro progrmama pasará los datos que haya en el excel
         * a la base de datos, usando los métodos de la clase ExcelReader. 
         * 
         * Por el contrario, si el valor de la propiedad action es save, este usará
         * los métodos de la clase ExcelWriter.
         */
        if (props.getProperty("action").equals("load")) {

            System.out.println("El nombre del archivo desde donde insertaremos datos es: " + props.getProperty("filentrada"));
            System.out.println("Estas insertando datos de un excel en la base de datos");
            ExcelReader reader = new ExcelReader();
            reader.loadWorbook(props.getProperty("filentrada"));
            System.out.println(reader.generateDDL());
            System.out.println(reader.executeDDL());
            System.out.println(reader.insertDDL());

        } else if (props.getProperty("action").equals("save")) {

            System.out.println("El nombre del archivo donde se insertaran datos desde la bd es: " + props.getProperty("filesalida"));
            System.out.println("Esta guardando datos de la base datos en un excel");
            ExcelWriter writer = new ExcelWriter();
            writer.loadWorkbook(props.getProperty("filesalida"));
            writer.writeExcel();;
            writer.saveWorkbook(props.getProperty("filesalida"));

        }

    
        //Test de conexion
        try (Connection conexion = Conexion.getConnection()) {
            if (conexion!=null) 
                System.out.println("Conectado correctamente.");
            else 
                System.out.println("Imposible conectar");
        } catch (Exception e) {
            System.err.println("No se pudo conectar.");            
        }
        
    }


}

