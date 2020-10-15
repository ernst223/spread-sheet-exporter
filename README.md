# spread-sheet-exporter
Java Maven Project that allows the creation of a excel or csv file. Or the file encrypted in a password protected zip file.
Can create a Excel File
Can create a CSV File
Can create a password protected zip Excel File
Can create a password protected zip CSV File

## Installation
```xml
<dependency>
  <groupId>org.ernst223</groupId>
  <artifactId>spread-sheet-exporter</artifactId>
  <version>1.0-SNAPSHOT</version>
</dependency>
```

Re sync the pom.xml file

## Usage
```java
import com.ernst223.exporttospreadsheet.SpreadSheetExporter;
```

### Excel and CSV generating
```java
SpreadSheetExporter spreadSheetExporter = new SpreadSheetExporter(List<Object>, "Filename");
File fileCSV = spreadSheetExporter.getCSV();
File fileExcel = spreadSheetExporter.getExcel();
```

### Excel and CSV Protected Zip
```java
SpreadSheetExporter spreadSheetExporter = new SpreadSheetExporter(List<Object>, "Filename", "Password");
File fileCSV = spreadSheetExporter.getCSV();
File fileExcel = spreadSheetExporter.getExcel();
```
