
package com.example.demo.service;

import com.example.demo.entity.DataRecord;
import org.apache.poi.ss.usermodel.*;
import java.io.File;
import java.io.FileInputStream;
import java.util.HashMap;
import java.util.Map;

public class ExcelReader {

    public DataRecord readExcelFile(String excelFilePath) throws Exception {
        FileInputStream fis = new FileInputStream(new File(excelFilePath));
        Workbook workbook = WorkbookFactory.create(fis);
        Sheet sheet = workbook.getSheetAt(0); // Première feuille

        Row headerRow = sheet.getRow(0); // La première ligne est l'en-tête
        DataRecord dataRecord = new DataRecord();

        for (int i = 1; i <= sheet.getLastRowNum(); i++) { // Commence à la 2ème ligne
            Row row = sheet.getRow(i);
            Map<String, String> rowData = new HashMap<>();

            for (Cell cell : row) {
                int colIndex = cell.getColumnIndex();
                String columnName = headerRow.getCell(colIndex).getStringCellValue(); // Nom de la colonne
                String cellValue = cell.toString(); // Valeur de la cellule

                rowData.put(columnName, cellValue); // Stockage dans la map
            }

            dataRecord.setData(rowData); // Associer les données à l'objet DataRecord
            // Sauvegarder l'objet dataRecord dans la base ici via le repository
        }

        workbook.close();
        fis.close();

        return dataRecord;
    }
}
