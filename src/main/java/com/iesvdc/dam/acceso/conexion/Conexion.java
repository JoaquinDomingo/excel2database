package com.iesvdc.dam.acceso.conexion;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
// import java.sql.SQLException;
import java.util.Properties;

public class Conexion {
    
    /**
      * Establece una conexión con la base de datos MySQL utilizando los parámetros definidos
      * en un fichero de propiedades llamado <b>config.properties</b>.
      * <p>
      * El fichero debe contener las claves:
      * <ul>
      *     <li><b>host</b>: dirección del servidor de base de datos</li>
      *     <li><b>port</b>: puerto del servidor (por ejemplo, 3306)</li>
      *     <li><b>database</b>: nombre de la base de datos</li>
      *     <li><b>user</b>: usuario de conexión</li>
      *     <li><b>password</b>: contraseña del usuario</li>
      * </ul>
      * <p>
      * Si ocurre algún error al leer el fichero o establecer la conexión, se muestra un mensaje
      * descriptivo por consola y se devuelve {@code null}.
      *
      * @return un objeto {@link java.sql.Connection} si la conexión se establece correctamente;
      *         {@code null} si ocurre algún error durante el proceso.
      *
      * @throws SecurityException si el acceso al fichero de propiedades está restringido
      *         por el sistema de seguridad de Java.
      *
      * @see java.util.Properties
      * @see java.sql.Connection
      * @see java.sql.DriverManager
      */
public static Connection getConnection() {
    Properties props = Config.getProperties("config.properties");
    if (props == null) {
        System.err.println("No se pudieron cargar las propiedades de configuración");
        return null;
    }
    
    String cadenaConexion = 
        "jdbc:mysql://" + props.getProperty("host") +
        ":" + props.getProperty("port") +
        "/" + props.getProperty("database") +
        "?useSSL=false&serverTimezone=UTC";
    Connection conn = null;
    try {            
        conn = DriverManager.getConnection(cadenaConexion, props);            
    } catch (SQLException sqle) {
        System.err.println(
            "Error al conectar a la base de datos: " +
            sqle.getLocalizedMessage());
    } 
    
    return conn;
}

public static void crearDatabase(Connection conn, String nombredatabase) {

    if (nombredatabase == null || nombredatabase.trim().isEmpty()) {
        System.out.println("El nombre de tu base de datos no puede estar vacio");
        return;
    }

    if (!nombredatabase.matches("[a-zA-Z0-9]+")) {
        System.out.println("El nombre de tu base datos no es valido");
        return;
    }

    String sql = "CREATE DATABASE " + nombredatabase;

    try (PreparedStatement statement = conn.prepareStatement(sql)) {
        statement.executeUpdate();
        System.out.println("Tu base de datos: " + nombredatabase + " ha sido creada correctamente");
    } catch (SQLException e) {
        e.printStackTrace();
    }
}


}
