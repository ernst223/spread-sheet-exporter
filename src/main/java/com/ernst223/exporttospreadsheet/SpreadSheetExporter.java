/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.ernst223.exporttospreadsheet;

import com.google.gson.Gson;
import net.lingala.zip4j.core.ZipFile;
import net.lingala.zip4j.exception.ZipException;
import net.lingala.zip4j.model.ZipParameters;
import net.lingala.zip4j.util.Zip4jConstants;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;

import java.io.*;
import java.util.*;
/**
 *
 * @author ernst223
 */
public class SpreadSheetExporter {
    private final Object T;
    private final Boolean isZip;
    private final String password;
    private final String fileName;
    private JSONArray jsonArray;

    // Constructor for a plain .csv or .xlsx
    public SpreadSheetExporter( Object T, String fileName){
        this.T = T;
        this.isZip = false;
        this.password = "";
        this.fileName = fileName;
        String temp = new Gson().toJson((ArrayList)T);
        try {
            JSONParser parser = new JSONParser();
            this.jsonArray = (JSONArray)parser.parse(temp);
        } catch (ParseException jsonException) {
            System.out.println(jsonException.getMessage());
            this.jsonArray = null;
        }
    }

    // Constructor for a password protected .zip
    public SpreadSheetExporter(Object T, String fileName, String password){
        this.T = T;
        this.isZip = true;
        this.password = password;
        this.fileName = fileName;
        String temp = new Gson().toJson((ArrayList)T);
        try {
            JSONParser parser = new JSONParser();
            this.jsonArray = (JSONArray)parser.parse(temp);
        } catch (ParseException jsonException) {
            System.out.println(jsonException.getMessage());
            this.jsonArray = null;
        }
    }

    public File toZip(File file) throws ZipException {
        ZipFile zipFile = new ZipFile(fileName + ".zip");
        // Setting parameters
        ZipParameters zipParameters = new ZipParameters();
        zipParameters.setCompressionMethod(Zip4jConstants.COMP_DEFLATE);
        zipParameters.setCompressionLevel(Zip4jConstants.DEFLATE_LEVEL_NORMAL);
        zipParameters.setEncryptFiles(true);
        zipParameters.setEncryptionMethod(Zip4jConstants.ENC_METHOD_STANDARD);
        // Setting password
        zipParameters.setPassword(password);
        zipFile.addFile(file, zipParameters);
        file.delete();
        return new File(fileName + ".zip");
    }

    public File getCSV() {
        File myFile = new File(fileName + ".csv");
        try {

            FileWriter out = new FileWriter(myFile);
            String[] myHeaders = getHeaders();
            CSVPrinter printer = new CSVPrinter(out,CSVFormat.DEFAULT.withHeader(myHeaders));

            List<Object> tempEntry;
            List<Map<String, String>> myBody = getBody();
            for (Map<String, String> entry: myBody) {
                tempEntry = new ArrayList<>();
                for (String header:myHeaders) {
                    String temp = entry.get(header);
                    tempEntry.add(temp);
                }
                printer.printRecord(tempEntry);
            }
            if(printer != null){
                printer.close();
                if(isZip)
                return toZip(myFile);
                return myFile;
            }else{
                return null;
            }

        } catch (IOException e) {
            System.out.println(e.getMessage());
            return null;
        } catch (ZipException e) {
            System.out.println(e.getMessage());
            return null;
        }
    }

//    public File getExcel() throws JSONException, IOException {
    public File getExcel() {
        Integer RowCounter = 2;
        Integer ColCounter = 0;
        File result = new File(fileName + ".xlsx");
        try{
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("sheet1");
            Row header = sheet.createRow(0);

            // Creating the headers
            CellStyle headerStyle = workbook.createCellStyle();
            headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            XSSFFont font = ((XSSFWorkbook) workbook).createFont();
            font.setFontName("Arial");
            font.setFontHeightInPoints((short)16);
            font.setBold(true);
            headerStyle.setFont(font);
            String[] myHeaders = getHeaders();
            Cell headerCell;
            for (String value: myHeaders) {
                headerCell = header.createCell(ColCounter);
                headerCell.setCellValue(value);
                headerCell.setCellStyle(headerStyle);
                ColCounter++;
            }
            ColCounter = 0;

            // Creating the body
            CellStyle style = workbook.createCellStyle();
            style.setWrapText(true);
            List<Map<String, String>> myBody = getBody();
            Row row;
            Cell cell;
            for (Map<String, String> entry: myBody) {
                row = sheet.createRow(RowCounter);
                for (String tempHeader: myHeaders) {
                    String value = entry.get(tempHeader);
                    cell = row.createCell(ColCounter);
                    cell.setCellValue(value);
                    cell.setCellStyle(style);
                    ColCounter++;
                }
                ColCounter = 0;
                RowCounter++;
            }

            // Writing File
            FileOutputStream outputStream = new FileOutputStream(result);
            workbook.write(outputStream);

            if(isZip)
                return toZip(result);;
            return result;
        } catch (IOException e) {
            System.out.println(e.getMessage());
            return null;
        } catch (ZipException e) {
            System.out.println(e.getMessage());
            return null;
        }
    }

    // Retrieving headers to be inserted
    private String[] getHeaders() {
        JSONObject jsonObject = (JSONObject) jsonArray.get(0);
        Iterator<String> keysTemp = jsonObject.keySet().iterator();
        Iterator<String> keys = jsonObject.keySet().iterator();
        Integer counter = 0;
        while (keysTemp.hasNext()) {
            keysTemp.next();
            counter++;
        }
        String[] result = new String[counter];
        counter =0;
        while (keys.hasNext()) {
            result[counter] = keys.next();
            counter++;
        }
        return result;
    }

    // Retrieving body to be inserted
    private List<Map<String, String>> getBody() {
        List<Map<String, String>> result = new ArrayList<>();
        Map<String, String> temp;
        for (int i = 0; i < jsonArray.size(); i++) {
            JSONObject jsonObject = (JSONObject) jsonArray.get(i);
            Iterator<String> keys = jsonObject.keySet().iterator();
            temp = new HashMap<>();
            while (keys.hasNext()) {
                String key = keys.next();
                Object value = jsonObject.get(key);
                temp.put(key, value.toString());
                result.add(temp);
            }
        }
        return result;
    }
}
