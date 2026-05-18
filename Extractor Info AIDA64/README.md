# Extractor Info AIDA64

Proyecto Java que extrae información de archivos HTML generados por AIDA64, muestra los datos en una interfaz gráfica y permite exportar un reporte a Excel.

## Descripción

`Extractor Info AIDA64` es una aplicación de escritorio Java que:

- Busca archivos `.htm` dentro de una carpeta seleccionada.
- Extrae datos clave de cada archivo AIDA64 (nombre de equipo, marca, modelo, número de serie, MAC, IP, licencia, datos de monitor, etc.).
- Muestra los resultados en una tabla interactiva.
- Genera un archivo Excel `.xlsx` con toda la información recopilada.

## Tecnologías

- Java 11
- Maven
- Swing (interfaz gráfica)
- Apache POI (generación de Excel)
- Jsoup (procesamiento de HTML)
- Log4j (registro)

## Estructura

- `pom.xml` - configuración del proyecto Maven.
- `src/main/java/org/example/Extractor.java` - lógica principal de la aplicación y extracción de datos.
- `src/main/java/org/example/Principal.java` - clase `main` que inicia la aplicación.

## Cómo usar

### Build

Desde la raíz del proyecto, ejecutar:

```powershell
mvn clean package
```

Esto genera un JAR con dependencias en `target/`.

### Ejecutar

Después de compilar, ejecutar:

```powershell
java -jar target/ExtractorInfo-1.0-SNAPSHOT-jar-with-dependencies.jar
```

### Flujo de uso

1. Ejecutar la aplicación.
2. Seleccionar la carpeta que contiene los archivos `.htm` generados por AIDA64.
3. Esperar a que se carguen y muestren los registros en la tabla.
4. Presionar `Generar Reporte Excel` para exportar los datos a un archivo `.xlsx`.

## Campos extraídos

- Chapa (nombre del archivo sin la extensión `.htm`)
- Nombre PC
- Marca PC
- Modelo PC
- Nº Serie PC
- MAC
- IP
- Licencia
- Monitor ID
- Monitor Nombre
- Monitor Modelo

## Notas

- El parser asume que los archivos `.htm` tienen tablas donde las etiquetas y valores están en filas con celdas `td`.
- Si no se encuentran archivos `.htm`, la aplicación muestra un mensaje informativo.

## Licencia

Proyecto de ejemplo sin licencia especificada.
