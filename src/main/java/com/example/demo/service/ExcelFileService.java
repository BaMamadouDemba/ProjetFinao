//AJOUT


package com.example.demo.service;

import com.example.demo.entity.DynamicColumn;
import org.apache.poi.ss.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Service
public class ExcelFileService {

    @Autowired
    private ExcelImportService excelImportService;

    public void readExcelFile(MultipartFile file) {
        try (Workbook workbook = WorkbookFactory.create(file.getInputStream())) {
            Sheet sheet = workbook.getSheetAt(0); // Lire la première feuille
            List<DynamicColumn> columns = new ArrayList<>();

            // Parcourir les lignes du fichier Excel
            for (Row row : sheet) {
                DynamicColumn column = new DynamicColumn();
                // Remplir les colonnes avec les données de chaque cellule
                for (Cell cell : row) {
                    switch (cell.getCellType()) {
                        case STRING:
                            column.setColumnValue(cell.getStringCellValue());
                            break;
                        case NUMERIC:
                            column.setColumnValue(String.valueOf(cell.getNumericCellValue()));
                            break;
                        case BOOLEAN:
                            column.setColumnValue(String.valueOf(cell.getBooleanCellValue()));
                            break;
                        // Ajouter d'autres types de cellules si nécessaire
                        default:
                            column.setColumnValue(""); // Valeur par défaut pour les cellules vides
                    }
                    // Définir le nom de la colonne (si vous avez une logique spécifique)
                    column.setColumnName("Column_" + cell.getColumnIndex());
                }
                columns.add(column); // Ajouter la colonne à la liste
            }

            // Appelez le service d'importation avec les colonnes extraites
            excelImportService.importData(columns);
        } catch (IOException e) {
            e.printStackTrace(); // Gérer l'exception (journaliser ou lancer une autre exception)
        }
    }
}
