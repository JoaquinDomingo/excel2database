# Proyecto excel2database

Programa ejemplo en java que vuelca un archivo Excel a una base de datos MySQL

Este programa genérico en  java (proyecto Maven), nos lee desde un fichero .xlsx y nos convertirá los datos que haya en ese fichero a una base de datos. Este programa nos lee la configuración del fichero .properties de Java, luego con Apache POI leo las hojas, el nombre de cada hoja, la primera fila de cada hoja será el nombre de los atributos. Procesamos el fichero Excel y creamos una estructura de datos con la información siguiente: La estructura principal es el libro, que contiene una lista de tablas y cada tabla contiene tuplas nombre del campo y tipo de dato.


Uso del programa:

```bash
excel2database -f fichero.xlsx -db agenda
```
## Como instalar Java para que funcione el proyecto


##  Como crear un proyecto Maven

Instalamos las dependencias Maven:

* apache poi
* apache poi ooxml
* mysql
## El archivo de propiedades

Creamos el archivo **config.properties**
```properties
user=root
password=s83n38DGB8d72
useUnicode=yes
useJDBCCompliantTimezoneShift=true
port=33307
database=agenda
host=localhost
driver=MySQL
outputFile=datos/salida.xlsx
inputFile=datos/entrada.xlsx
useSSL=false
serverTimezone=Europe/Madrid
allowPublicKeyRetrieval=true
```
Esto **jamas** debe ser usado en la producción

* `useSSL=false`: No encripta la conexion

* `allowPublicKeyRetrieval=true`: 
## Detectando el tipo de dato con Apache POI

Con **Apache POI** puedes inspeccionar el **tipo de dato almacenado en una celda de un Excel (.xlsx) y actuar según corresponda. Cuando trabajas con una celda (`Cell`), puedes preguntar su tipo mediante: 
```java
cell.getCellType()
```

Esto devuelve un valor del enum `CellType`, que puede ser:

| Tipo (`CellType`) | Significado                                                             |
| ----------------- | ----------------------------------------------------------------------- |
| `NUMERIC`         | Número (entero o decimal, o incluso la fecha si se indica en el formato)|
| `STRING`          | Texto                                                                   |
| `BOOLEAN`         | Verdadero/Falso                                                         |
| `FORMULA`         | Celda con una fórmula                                                   |
| `BLANK`           | Celda vacía                                                             |
| `ERROR`           | Celda con error                                                         |

* Excel nos almacena **las fechas y las horas** como números (a partir del 1/1/1900). Para poder distinguirlas se utiliza este método:

```java
DateUtil.isCellDateFormatted(cell)
```
* Si este método nos devuelve un `true`, el contenido es efecto una **fecha o hora**, y se obtiene de esta manera:

```java
Date fecha = cell.getDateCellValue();
```

* En el caso en que no sea `true`, se puede tratar como un número:

```java
double valor = cell.getNumericalCellValue();
```

* En Excel, todos los números (ya sean enteros o decimales), se almacenan como doble, ya que no existe una distinción **formal** entre estos.

```java
double valor = cell.getNumericalCellValue();
if (valor == Math.floor(valor)){
    System.out.println("Nº Entero: " + (int) valor)
} else {
    System.out.println("Nº Decimal: " + valor)
}
```

Un ejemplo de cómo hacer la detección sería:

```java

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.DateUtil;

public class ExcelUtils {

    private static final double EPSILON = 1e-10;

    /**
     * Devuelve un String indicando el tipo de dato de la celda.
     * Puede ser: Entero, Decimal, Texto, Booleano, Fecha, Vacía, Fórmula, Error
     */
    public static String getTipoDato(Cell cell) {
        if (cell == null) {
            return "Vacía";
        }

        switch (cell.getCellType()) {
            case STRING:
                return "Texto";

            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return "Fecha";
                } else {
                    double valor = cell.getNumericCellValue();
                    if (Math.abs(valor - Math.floor(valor)) < EPSILON) {
                        return "Nº Entero";
                    } else {
                        return "Nº Decimal";
                    }
                }

            case BOOLEAN:
                return "Boolean";

            case FORMULA:
                // Puedes decidir si quieres evaluar la fórmula o solo indicar que es fórmula
                return "Fórmula";

            case BLANK:
                return "Vacía";

            case ERROR:
                return "Error";

            default:
                return "Desconocido";
        }
    }
}

```


## Guardar archivos .properties en archivos .xml

```java
        Properties p = getProperties("config.properties");

        try (FileOutputStream fos = new FileOutputStream("propiedades.xml")){
            p.storeToXML(fos, "Fichero de configuracion en xml");
        } catch (Exception e) {
            e.printStackTrace();
        }
```
* Método getProperties()
```java
    static public Properties getProperties(String nombreArchivo){
        Properties props = new Properties();
        try(FileInputStream is = new FileInputStream(nombreArchivo)){
            props.load(is);
        }catch (Exception e) {
            System.out.println("Imposible cargar el archivo de propiedades");
        }
        return props; 
    }
```
## Recorrer archivos .properties
  * También hay que declarar el metodo getProperties() antes mencionado
```java

  Properties p = getProperties("config.properties");

  Enumeration e = p.propertyNames();
   while (e.hasMoreElements()) {
      String nombre = e.nextElement().toString();
      System.out.println("Propiedad: " + nombre + " valor: " + p.getProperty(nombre));
  }
```

## Metodo getConnection()

```java
    public static Connection getConnection(Properties props) throws SQLException{
        
        String cadenaConexion = 
        "jdbc:mysql:"+ 
        props.getProperty("host") + //host
        ":" +  
        props.getProperty("port") + //puerto
        "/"+ 
        props.getProperty("database"); //base de datos

        Connection connection = DriverManager.getConnection(cadenaConexion,props);
        
        return connection;
        
    }
```
## Métodos Crud
* **Insertar Registros**
```java

 private static void  insertarRegistro(Connection connection, String nombre,
    String apellido) throws SQLException {
        String insert = "INSERT INTO personas (nombre, apellido)" +
        "VALUES (Joaquin, Domingo)";
        try (PreparedStatement statement = connection.prepareStatement(insert)){
            prepareStatement.setString(1, nombre);
            prepareStatement.setString(2, apellido);
            int filasactualizadas = prepareStatement.executeUpdate();
            if (filasactualizadas > 0) {
                System.out.println("Registro insertado con exito");
            }
        }
    }
```

* **Consultar Registros**

```java
private static void consultarRegistros(Connection connection) throws SQLException{
        String select = "SELECT * FROM personas";
        try(PreparedStatement statement = connection.prepareStatement(select); 
            ResultSet resultSet = prepareStatement.executeQuery()){
                while (resultSet.next()) {
                    int id = resultSet.getInt("id");
                    String nombre = resultSet.getString("nombre");
                    String apellido = resultSet.getString("apellido");
                    System.out.println("Id: " + id + ", Nombre: " + nombre + ", Apellido: " 
                    + apellido);
                }
            }
    }
```
* **Actualizar Registros**
```java

     private static void actualizarRegistro(Connection connection,int
        id, String nuevoNombre, String nuevoApellido)throws SQLException {
        String updateQuery ="UPDATE tabla_ejemplo SET nombre = ?, apellido =?WHERE id =?";
        try (PreparedStatement preparedStatement =connection.prepareStatement(updateQuery)){
          preparedStatement.setString(1, nuevoNombre);
          preparedStatement.setString(2, nuevoApellido);
          preparedStatement.setInt(3, id);
          int filasAfectadas = preparedStatement.executeUpdate();
          if (filasAfectadas >0) {
            System.out.println("Registroactualizado con éxito.");
             }
            }
        }
```
* **Borrar Registros**
```java
private static void eliminarRegistro(Connection connection, int id)throws SQLException {
        String deleteQuery ="DELETE FROM tabla_ejemplo WHERE id =?";
        try (PreparedStatement preparedStatement =connection.prepareStatement(deleteQuery)){
            preparedStatement.setInt(1, id);
            int filasAfectadas = preparedStatement.executeUpdate();
            if (filasAfectadas >0) {
                System.out.println("Registro eliminado conéxito.");
            }
        }
    }
```

## Apéndice Repaso de SQL

* Para crear una base de datos en MySQL hacemos: 

```sql
create database `agenda`;
```

```sql
create database `agenda` collate 'utf16_spanish_ci';
```

* Para borrar una base de datos en MySQL hacemos:

```sql
drop database `agenda`;
```

* Crear una tabla de personas:

```sql
CREATE TABLE `personas` (
  `nombre` varchar(100) NOT NULL,
  `apellidos` varchar(300) NOT NULL,
  `email` varchar(100),
  `teléfono` varchar(12),
  `género` enum('FEMENINO','MASCULINO','NEUTRO','OTRO') NOT NULL,
  `id` int NOT NULL AUTO_INCREMENT PRIMARY KEY
) ENGINE='InnoDB';
```

* Como insertar valores:

```sql
INSERT INTO `personas` (`nombre`, `apellidos`, `email`, `teléfono`, `género`)
VALUES ('Joaquin', 'Domingo Domingo', 'jdomdom0901@g.educaand.es', '+34987654321', 'MASCULINO');
```


* Como consultar la base de datos:

```sql
SELECT * FROM `personas`;
```
* Como consultar la base de datos con limitaciones:

  * El modificador **limit**  nos limita el numero de datos que se nos motraría por pantalla. 
  * Mientras que el modificador **offset** hace que empecemos directamente por la posicion de la base de datos que le indiquemos
```sql
SELECT * FROM `personas` LIMIT 50 OFFSET 10;
```
## Estructura del proyecto


- `Readme.md` — documentación del proyecto (este fichero).
- `config.properties` — archivo de configuración con parámetros de conexión y rutas (ejemplo, NO usar en producción).
- `pom.xml` — descriptor Maven con dependencias y configuración de build.
- `datos/`
    - `personas.xlsx` — fichero Excel de entrada .
    - `personasdesdebdat.xlsx` — fichero Excel de salida.
- `src/main/java/com/iesvdc/dam/acceso/`
    - `Excel2Database.java` — clase principal que orquesta la lectura del Excel y el volcado a la base de datos y viceversa.
    - `conexion/`
        - `Conexion.java` — utilidades para obtener la conexión JDBC a MySQL.
        - `Config.java` — carga y parseo de `config.properties` en `Properties`.
    - `excelutil/`
        - `ExcelReader.java` — lectura del fichero .xlsx; detecta tipos de campo, hojas y construye modelos.
        - `ExcelWriter.java` — lectura de la base de datos, detecta los datos, y los introduce a un nuevo .xlsx
    - `modelo/`
        - `WorkbookModel.java` — modelo que representa el libro (workbook) con sus tablas/hojas.
        - `TableModel.java` — modelo que representa una hoja/tabla (nombre, lista de campos).
        - `FieldModel.java` — modelo que representa un campo/columna (nombre, tipo, posición, posible metadato).
        - `FieldType.java` — enum o clase que define los tipos de campo reconocidos (Texto, Entero, Decimal, Fecha, Booleano, etc.).
- `test/java/com/iesvdc/dam/acceso/AppTest.java`
- `stack-excelreader/`
    - `docker-compose.yml` — composición Docker (por ejemplo, contenedor MySQL) para desarrollo/pruebas.
    - `db/init.sql` — script SQL para inicializar la base de datos y tablas de ejemplo.
- `target/` — salida del build Maven (artefactos compilados). Esta carpeta se genera automáticamente y no debe versionarse.

## Autor

- Nombre: Joaquín Domingo Domingo
- GitHub: `@JoaquinDomingo`
- Correo de contacto: `dojoaquindo@gmail.com`

