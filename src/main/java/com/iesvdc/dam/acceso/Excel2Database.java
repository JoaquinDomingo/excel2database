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
        Properties props = Config.getProperties("config.properties");
        System.out.println("El nombre del archivo es: " + props.getProperty("file"));
        System.out.println("La acción es: " + props.getProperty("action"));


        //TO-DO si la accion es LOAD:
        if (props.getProperty("action").equals("load")) {
            System.out.println("Estas insertando datos en la base de datos");
            ExcelReader reader = new ExcelReader();
            reader.loadWorbook(props.getProperty("file"));
            System.out.println(reader.generateDDL());
            System.out.println(reader.executeDDL());
            System.out.println(reader.insertDDL());
        }

        //Si la accion es SAVE:
        if (props.getProperty("action").equals("save")) {
            System.out.println("Esta guardando datos de la base datos");
            ExcelWriter writer = new ExcelWriter();
            
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

