/*package com.example.demo.service;

import com.example.demo.entity.ExcelData;
import com.example.demo.repository.ExcelDataRepository;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.InputStream;

@Service
public class ExcelToXmlService {

    @Autowired
    private ExcelDataRepository excelDataRepository;



    public String importExcelData(InputStream inputStream, String fileName) {
        try (Workbook workbook = new XSSFWorkbook(inputStream)) {


            // Ouvrir la première feuille
            Sheet sheet = workbook.getSheetAt(0);
            String[] columnNames = new String[sheet.getRow(0).getLastCellNum()];

            // Remplir le tableau avec les noms de colonnes
            Row headerRow = sheet.getRow(0);
            for (int j = 0; j < headerRow.getLastCellNum(); j++) {
                Cell headerCell = headerRow.getCell(j);
                columnNames[j] = getCellValue(headerCell);
            }

            // Afficher le nom du fichier une seule fois
            System.out.println("Nom du fichier : " + fileName);

            // Parcourir les lignes du fichier Excel à partir de la deuxième ligne
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row != null && !isRowEmpty(row)) {
                    // Créer une nouvelle entité ExcelData
                    ExcelData excelData = new ExcelData();
                    excelData.setFileName(fileName);
                    excelData.setSheetName(sheet.getSheetName());

                    // Parcourir les colonnes pour extraire les valeurs
                    for (int col = 0; col < columnNames.length; col++) {
                        Cell cell = row.getCell(col);
                        String cellValue = getCellValue(cell);
                        if (!cellValue.isEmpty()) {
                            // Ajoutez ici la logique pour stocker les valeurs spécifiques dans l'entité
                        }
                    }

                    // Enregistrer l'entité ExcelData dans la base de données
                    excelDataRepository.save(excelData);
                }
            }

            return "Données enregistrées dans la base de données.";
        } catch (Exception e) {
            e.printStackTrace();
            return "Erreur lors de l'enregistrement des données : " + e.getMessage();
        }
    }

    public ExcelData getExcelDataById(Long id) {
        return excelDataRepository.findById(id).orElse(null); // Retourne null si pas trouvé
    }

    // Méthode utilitaire pour récupérer la valeur d'une cellule Excel
    private String getCellValue(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString(); // Formatez la date si nécessaire
                } else {
                    return String.valueOf(cell.getNumericCellValue()).replaceAll("\\.0*$", ""); // Évite la notation scientifique
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            default:
                return "";
        }
    }


    // Vérifie si une ligne est vide
    private boolean isRowEmpty(Row row) {
        for (Cell cell : row) {
            if (cell != null && getCellValue(cell).length() > 0) {
                return false; // Si une cellule non vide est trouvée
            }
        }
        return true; // La ligne est vide
    }
}*/

/*
package com.example.demo.service;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import java.io.InputStream;

@Service
public class ExcelToXmlService {

    public String importExcelData(InputStream inputStream, String fileName) {
        StringBuilder result = new StringBuilder();

        try (Workbook workbook = new XSSFWorkbook(inputStream)) {
            // Ouvrir la première feuille
            Sheet sheet = workbook.getSheetAt(0);

            // Construire un tableau HTML (ou du texte brut)
            result.append("<table border='1'>");

            // Parcourir chaque ligne du fichier Excel
            for (int i = 0; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    result.append("<tr>");  // Pour HTML

                    // Parcourir chaque cellule de la ligne
                    for (int j = 0; j < row.getLastCellNum(); j++) {
                        Cell cell = row.getCell(j);
                        String cellValue = getCellValue(cell);
                        result.append("<td>").append(cellValue).append("</td>");  // Pour HTML
                    }

                    result.append("</tr>");  // Pour HTML
                }
            }
            result.append("</table>");  // Pour HTML

        } catch (Exception e) {
            e.printStackTrace();
            return "Erreur lors de la lecture du fichier Excel : " + e.getMessage();
        }

        return result.toString();  // Retourne le contenu sous forme de tableau HTML
    }

    // Méthode utilitaire pour récupérer la valeur d'une cellule Excel
    private String getCellValue(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return String.valueOf(cell.getNumericCellValue()).replaceAll("\\.0*$", "");
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            default:
                return "";
        }
    }
}
*/
/*package com.example.demo.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;

@Service
public class ExcelToXmlService {

    public void importExcelData(InputStream inputStream, String jdbcURL, String username, String password) {
        try (Connection connection = DriverManager.getConnection(jdbcURL, username, password);
             Workbook workbook = new XSSFWorkbook(inputStream)) {

            Sheet sheet = workbook.getSheetAt(0); // Lire la première feuille
            Row headerRow = sheet.getRow(0); // Lire la première ligne comme en-tête
            int columnCount = headerRow.getPhysicalNumberOfCells(); // Nombre de colonnes

            // Liste pour stocker les colonnes valides
            List<String> validColumns = new ArrayList<>();
            for (int i = 0; i < columnCount; i++) {
                String columnName = headerRow.getCell(i).getStringCellValue().trim();
                if (!columnName.isEmpty()) {
                    validColumns.add(columnName);
                } else {
                    System.out.println("Colonne vide détectée à l'index " + i);
                }
            }

            // Construire la requête SQL
            StringBuilder sqlBuilder = new StringBuilder("INSERT INTO excel_data (");
            for (int i = 0; i < validColumns.size(); i++) {
                sqlBuilder.append(validColumns.get(i));
                if (i < validColumns.size() - 1) {
                    sqlBuilder.append(", ");
                }
            }
            sqlBuilder.append(") VALUES (");
            for (int i = 0; i < validColumns.size(); i++) {
                sqlBuilder.append("?");
                if (i < validColumns.size() - 1) {
                    sqlBuilder.append(", ");
                }
            }
            sqlBuilder.append(")");

            PreparedStatement preparedStatement = connection.prepareStatement(sqlBuilder.toString());

            // Lire les données et insérer dans la base de données
            // Lire les données et insérer dans la base de données
            for (int i = 1; i <= sheet.getLastRowNum(); i++) { // Commencer à la deuxième ligne
                Row row = sheet.getRow(i);
                if (row != null) { // Vérifier si la ligne n'est pas nulle
                    for (int j = 0; j < validColumns.size(); j++) { // Utiliser validColumns.size()
                        Cell cell = row.getCell(j);
                        if (cell != null) {
                            // Gérer les différents types de cellules
                            switch (cell.getCellType()) {
                                case STRING:
                                    System.out.println("Valeur de la cellule (" + i + ", " + j + ") : " + cell.getStringCellValue());
                                    preparedStatement.setString(j + 1, cell.getStringCellValue());
                                    break;
                                case NUMERIC:
                                    if (DateUtil.isCellDateFormatted(cell)) {
                                        System.out.println("Date de la cellule (" + i + ", " + j + ") : " + cell.getDateCellValue());
                                        preparedStatement.setDate(j + 1, new java.sql.Date(cell.getDateCellValue().getTime()));
                                    } else {
                                        System.out.println("Numérique de la cellule (" + i + ", " + j + ") : " + cell.getNumericCellValue());
                                        preparedStatement.setDouble(j + 1, cell.getNumericCellValue());
                                    }
                                    break;
                                case BOOLEAN:
                                    System.out.println("Booléen de la cellule (" + i + ", " + j + ") : " + cell.getBooleanCellValue());
                                    preparedStatement.setBoolean(j + 1, cell.getBooleanCellValue());
                                    break;
                                case FORMULA:
                                    System.out.println("Formule de la cellule (" + i + ", " + j + ") : " + cell.getCellFormula());
                                    preparedStatement.setString(j + 1, cell.getCellFormula());
                                    break;
                                default:
                                    System.out.println("Cellule vide ou de type inconnu à l'index " + j);
                                    preparedStatement.setString(j + 1, null); // Gestion des cellules vides ou de types non pris en charge
                                    break;
                            }
                        } else {
                            System.out.println("Cellule vide à l'index " + j);
                            preparedStatement.setString(j + 1, null); // Gestion des cellules vides
                        }
                    }
                    preparedStatement.executeUpdate(); // Exécuter l'insertion après chaque ligne
                } else {
                    System.out.println("Ligne vide à l'index " + i); // Afficher si la ligne est vide
                }
            }
            System.out.println("Les données ont été insérées avec succès.");

        } catch (IOException | SQLException e) {
            e.printStackTrace();
        }
    }
}
*/
/*
package com.example.demo.service;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.InputStream;
import java.sql.*;
import java.util.ArrayList;
import java.util.List;

@Service
public class ExcelToXmlService {

    public void importExcelData(InputStream inputStream, String jdbcURL, String username, String password) {
        try (Connection connection = DriverManager.getConnection(jdbcURL, username, password);
             Workbook workbook = new XSSFWorkbook(inputStream)) {

            // Ouvrir la première feuille
            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            int columnCount = headerRow.getPhysicalNumberOfCells();

            // Obtenir les noms de colonnes
            List<String> columns = new ArrayList<>();
            for (int i = 0; i < columnCount; i++) {
                String columnName = headerRow.getCell(i).getStringCellValue().trim();
                columns.add(columnName);
            }

            // Créer dynamiquement une table si elle n'existe pas
            String tableName = "TableTest";
            createTableIfNotExists(connection, tableName, columns);

            // Insérer les données dans la table
            insertExcelDataIntoTable(connection, sheet, tableName, columns);

            System.out.println("Données insérées avec succès.");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // Méthode pour créer la table dynamiquement
    private void createTableIfNotExists(Connection connection, String tableName, List<String> columns) throws SQLException {
        StringBuilder createTableQuery = new StringBuilder("CREATE TABLE IF NOT EXISTS " + tableName + " (id INT AUTO_INCREMENT PRIMARY KEY, ");
        for (String column : columns) {
            createTableQuery.append(column).append(" VARCHAR(255), ");
        }
        // Retirer la dernière virgule et espace
        createTableQuery.setLength(createTableQuery.length() - 2);
        createTableQuery.append(")");

        try (Statement statement = connection.createStatement()) {
            statement.executeUpdate(createTableQuery.toString());
            System.out.println("Table créée ou déjà existante : " + tableName);
        }
    }

    // Méthode pour insérer les données dans la table
    private void insertExcelDataIntoTable(Connection connection, Sheet sheet, String tableName, List<String> columns) throws SQLException {
        StringBuilder insertQuery = new StringBuilder("INSERT INTO " + tableName + " (");
        for (String column : columns) {
            insertQuery.append(column).append(", ");
        }
        insertQuery.setLength(insertQuery.length() - 2); // Retirer la dernière virgule
        insertQuery.append(") VALUES (");
        for (int i = 0; i < columns.size(); i++) {
            insertQuery.append("?, ");
        }
        insertQuery.setLength(insertQuery.length() - 2); // Retirer la dernière virgule
        insertQuery.append(")");

        try (PreparedStatement preparedStatement = connection.prepareStatement(insertQuery.toString())) {
            // Lire les lignes du fichier Excel et insérer les données
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    for (int j = 0; j < columns.size(); j++) {
                        Cell cell = row.getCell(j);
                        preparedStatement.setString(j + 1, getCellValue(cell));
                    }
                    preparedStatement.executeUpdate();
                }
            }
        }
    }


    // Méthode pour obtenir la valeur de la cellule
    private String getCellValue(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return String.valueOf(cell.getNumericCellValue()).replaceAll("\\.0*$", "");
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            default:
                return "";
        }
    }
}*/
/*Code par defaut
package com.example.demo.service;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.InputStream;
import java.sql.*;
import java.util.ArrayList;
import java.util.List;

@Service
public class ExcelToXmlService {

    public void importExcelData(InputStream inputStream, String jdbcURL, String username, String password) {
        try (Connection connection = DriverManager.getConnection(jdbcURL, username, password);
             Workbook workbook = new XSSFWorkbook(inputStream)) {

            // Ouvrir la première feuille
            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            int columnCount = headerRow.getPhysicalNumberOfCells();

            // Obtenir les noms et types de colonnes
            List<String> columns = new ArrayList<>();
            List<String> columnTypes = new ArrayList<>();

            for (int i = 0; i < columnCount; i++) {
                String columnName = headerRow.getCell(i).getStringCellValue().trim();
                columns.add(columnName);

                // Détecter le type de données de la première ligne après l'en-tête pour déterminer le type de colonne
                CellType cellType = sheet.getRow(1).getCell(i).getCellType();
                columnTypes.add(getSQLType(cellType));
            }

            // Créer dynamiquement une table si elle n'existe pas
            String tableName = "TableTest";
            createTableIfNotExists(connection, tableName, columns, columnTypes);

            // Insérer les données dans la table
            insertExcelDataIntoTable(connection, sheet, tableName, columns);

            System.out.println("Données insérées avec succès.");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // Méthode pour déterminer le type SQL basé sur le type de cellule Excel
    private String getSQLType(CellType cellType) {
        switch (cellType) {
            case NUMERIC:
                return "DOUBLE";  // Utiliser INT si nécessaire
            case BOOLEAN:
                return "BOOLEAN";
            case STRING:
                return "TEXT";
            case FORMULA:
                return "TEXT";
            default:
                return "TEXT";  // Type par défaut
        }
    }


    // Méthode pour déterminer le type SQL basé sur le type de cellule Excel
    private String getSQLType(CellType cellType) {
        switch (cellType) {
            case NUMERIC:
                return "DOUBLE";  // Utiliser DOUBLE pour les valeurs numériques
            case BOOLEAN:
                return "BOOLEAN";
            case STRING:
            case FORMULA:
            default:
                return "VARCHAR(255)";  // Forcer VARCHAR pour tout autre type
        }
    }





    // Méthode pour créer la table dynamiquement en fonction des types de colonnes détectés
    private void createTableIfNotExists(Connection connection, String tableName, List<String> columns, List<String> columnTypes) throws SQLException {
        StringBuilder createTableQuery = new StringBuilder("CREATE TABLE IF NOT EXISTS " + tableName + " (id INT AUTO_INCREMENT PRIMARY KEY, ");
        for (int i = 0; i < columns.size(); i++) {
            createTableQuery.append(columns.get(i)).append(" ").append(columnTypes.get(i)).append(", ");
        }
        // Retirer la dernière virgule et espace
        createTableQuery.setLength(createTableQuery.length() - 2);
        createTableQuery.append(")");

        try (Statement statement = connection.createStatement()) {
            statement.executeUpdate(createTableQuery.toString());
            System.out.println("Table créée ou déjà existante : " + tableName);
        }
    }

    // Méthode pour insérer les données dans la table
    private void insertExcelDataIntoTable(Connection connection, Sheet sheet, String tableName, List<String> columns) throws SQLException {
        StringBuilder insertQuery = new StringBuilder("INSERT INTO " + tableName + " (");
        for (String column : columns) {
            insertQuery.append(column).append(", ");
        }
        insertQuery.setLength(insertQuery.length() - 2); // Retirer la dernière virgule
        insertQuery.append(") VALUES (");
        for (int i = 0; i < columns.size(); i++) {
            insertQuery.append("?, ");
        }
        insertQuery.setLength(insertQuery.length() - 2); // Retirer la dernière virgule
        insertQuery.append(")");

        try (PreparedStatement preparedStatement = connection.prepareStatement(insertQuery.toString())) {
            // Lire les lignes du fichier Excel et insérer les données
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    for (int j = 0; j < columns.size(); j++) {
                        Cell cell = row.getCell(j);
                        setPreparedStatementValue(preparedStatement, j + 1, cell);
                    }
                    preparedStatement.executeUpdate();
                }
            }
        }
    }

    // Méthode pour définir la valeur appropriée dans le PreparedStatement selon le type de cellule
    private void setPreparedStatementValue(PreparedStatement preparedStatement, int parameterIndex, Cell cell) throws SQLException {
        if (cell == null) {
            preparedStatement.setNull(parameterIndex, Types.VARCHAR); // Par défaut NULL
            return;
        }

        switch (cell.getCellType()) {
            case STRING:
                preparedStatement.setString(parameterIndex, cell.getStringCellValue());
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    preparedStatement.setDate(parameterIndex, new java.sql.Date(cell.getDateCellValue().getTime()));
                } else {
                    preparedStatement.setDouble(parameterIndex, cell.getNumericCellValue());
                }
                break;
            case BOOLEAN:
                preparedStatement.setBoolean(parameterIndex, cell.getBooleanCellValue());
                break;
            default:
                preparedStatement.setString(parameterIndex, "");
        }
    }
}
*/

/*package com.example.demo.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.InputStream;
import java.sql.*;
import java.util.ArrayList;
import java.util.List;

@Service
public class ExcelToXmlService {

    public void importExcelData(InputStream inputStream, String jdbcURL, String username, String password) {
        try (Connection connection = DriverManager.getConnection(jdbcURL, username, password);
             Workbook workbook = new XSSFWorkbook(inputStream)) {

            // Ouvrir la première feuille
            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            int columnCount = headerRow.getPhysicalNumberOfCells();

            // Obtenir les noms et types de colonnes
            List<String> columns = new ArrayList<>();
            List<String> columnTypes = new ArrayList<>();

            for (int i = 0; i < columnCount; i++) {
                String columnName = headerRow.getCell(i).getStringCellValue().trim();
                columns.add(columnName);

                // Détecter le type de données de la première ligne après l'en-tête pour déterminer le type de colonne
                CellType cellType = sheet.getRow(1).getCell(i).getCellType();
                columnTypes.add(getSQLType(cellType));
            }

            // Log des colonnes et types détectés
            System.out.println("Colonnes détectées : " + columns);
            System.out.println("Types de colonnes : " + columnTypes);

            // Créer dynamiquement une table si elle n'existe pas
            String tableName = "TableTest";
            createTableIfNotExists(connection, tableName, columns, columnTypes);

            // Insérer les données dans la table
            insertExcelDataIntoTable(connection, sheet, tableName, columns);

            System.out.println("Données insérées avec succès.");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // Méthode pour déterminer le type SQL basé sur le type de cellule Excel
    private String getSQLType(CellType cellType) {
        switch (cellType) {
            case NUMERIC:
                return "DOUBLE";  // Utiliser DOUBLE pour les valeurs numériques
            case BOOLEAN:
                return "BOOLEAN";
            case STRING:
            case FORMULA:
            default:
                return "VARCHAR(255)";  // Forcer VARCHAR pour tout autre type
        }
    }

    // Méthode pour créer la table dynamiquement en fonction des types de colonnes détectés
    private void createTableIfNotExists(Connection connection, String tableName, List<String> columns, List<String> columnTypes) throws SQLException {
        StringBuilder createTableQuery = new StringBuilder("CREATE TABLE IF NOT EXISTS " + tableName + " (id INT AUTO_INCREMENT PRIMARY KEY, ");
        for (int i = 0; i < columns.size(); i++) {
            createTableQuery.append(columns.get(i)).append(" ").append(columnTypes.get(i)).append(", ");
        }
        // Retirer la dernière virgule et espace
        createTableQuery.setLength(createTableQuery.length() - 2);
        createTableQuery.append(")");

        try (Statement statement = connection.createStatement()) {
            statement.executeUpdate(createTableQuery.toString());
            System.out.println("Table créée ou déjà existante : " + tableName);
        }
    }

    // Méthode pour insérer les données dans la table
    private void insertExcelDataIntoTable(Connection connection, Sheet sheet, String tableName, List<String> columns) throws SQLException {
        StringBuilder insertQuery = new StringBuilder("INSERT INTO " + tableName + " (");
        for (String column : columns) {
            insertQuery.append(column).append(", ");
        }
        insertQuery.setLength(insertQuery.length() - 2); // Retirer la dernière virgule
        insertQuery.append(") VALUES (");
        for (int i = 0; i < columns.size(); i++) {
            insertQuery.append("?, ");
        }
        insertQuery.setLength(insertQuery.length() - 2); // Retirer la dernière virgule
        insertQuery.append(")");

        // Log de la requête SQL
        System.out.println("Requête SQL générée : " + insertQuery);

        try (PreparedStatement preparedStatement = connection.prepareStatement(insertQuery.toString())) {
            // Lire les lignes du fichier Excel et insérer les données
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    for (int j = 0; j < columns.size(); j++) {
                        Cell cell = row.getCell(j);
                        setPreparedStatementValue(preparedStatement, j + 1, cell);
                    }
                    preparedStatement.addBatch();  // Ajouter la requête à un batch
                }
            }
            // Exécuter toutes les requêtes en une seule fois
            preparedStatement.executeBatch();
        }
    }

    // Méthode pour définir la valeur appropriée dans le PreparedStatement selon le type de cellule
    /*private void setPreparedStatementValue(PreparedStatement preparedStatement, int parameterIndex, Cell cell) throws SQLException {

        if (cell == null) {
            preparedStatement.setNull(parameterIndex, Types.VARCHAR); // Par défaut NULL
            return;
        }

        switch (cell.getCellType()) {
            case STRING:
                preparedStatement.setString(parameterIndex, cell.getStringCellValue());
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    preparedStatement.setDate(parameterIndex, new java.sql.Date(cell.getDateCellValue().getTime()));
                } else {
                    preparedStatement.setDouble(parameterIndex, cell.getNumericCellValue());
                }
                break;
            case BOOLEAN:
                preparedStatement.setBoolean(parameterIndex, cell.getBooleanCellValue());
                break;
            default:
                preparedStatement.setString(parameterIndex, "");  // Cas des cellules vides ou autres
        }
    }*/

    /*private void setPreparedStatementValue(PreparedStatement preparedStatement, int parameterIndex, Cell cell) throws SQLException {
        if (cell == null) {
            preparedStatement.setNull(parameterIndex, Types.VARCHAR); // Par défaut NULL
            System.out.println("Cellule vide à l'index " + parameterIndex);
            return;
        }

        switch (cell.getCellType()) {
            case STRING:
                String cellValue = cell.getStringCellValue();
                System.out.println("Valeur STRING à insérer : " + cellValue);
                preparedStatement.setString(parameterIndex, cellValue);
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    java.sql.Date dateValue = new java.sql.Date(cell.getDateCellValue().getTime());
                    System.out.println("Valeur DATE à insérer : " + dateValue);
                    preparedStatement.setDate(parameterIndex, dateValue);
                } else {
                    double numericValue = cell.getNumericCellValue();
                    System.out.println("Valeur NUMERIC à insérer : " + numericValue);
                    preparedStatement.setDouble(parameterIndex, numericValue);
                }
                break;
            case BOOLEAN:
                boolean booleanValue = cell.getBooleanCellValue();
                System.out.println("Valeur BOOLEAN à insérer : " + booleanValue);
                preparedStatement.setBoolean(parameterIndex, booleanValue);
                break;
            default:
                System.out.println("Valeur par défaut insérée : \"\"");
                preparedStatement.setString(parameterIndex, "");  // Cas des cellules vides ou autres
        }
    }
}
*/


/*Code bi
package com.example.demo.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.InputStream;
import java.sql.*;
import java.util.ArrayList;
import java.util.List;

@Service
public class ExcelToXmlService {

   /* public void importExcelData(InputStream inputStream, String jdbcURL, String username, String password) {
        try (Connection connection = DriverManager.getConnection(jdbcURL, username, password);
             Workbook workbook = new XSSFWorkbook(inputStream)) {

            // Ouvrir la première feuille
            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            int columnCount = headerRow.getPhysicalNumberOfCells();

            // Obtenir les noms et types de colonnes
            List<String> columns = new ArrayList<>();
            List<String> columnTypes = new ArrayList<>();

            for (int i = 0; i < columnCount; i++) {
                String columnName = headerRow.getCell(i).getStringCellValue().trim();
                columns.add(columnName);

                // Détecter le type de données de la première ligne après l'en-tête pour déterminer le type de colonne
                CellType cellType = sheet.getRow(1).getCell(i).getCellType();
                columnTypes.add(getSQLType(cellType));
            }

            // Log des colonnes et types détectés
            System.out.println("Colonnes détectées : " + columns);
            System.out.println("Types de colonnes : " + columnTypes);

            // Créer dynamiquement une table si elle n'existe pas
            String tableName = "TableTest";
            createTableIfNotExists(connection, tableName, columns, columnTypes);

            // Insérer les données dans la table
            insertExcelDataIntoTable(connection, sheet, tableName, columns);

            System.out.println("Données insérées avec succès.");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }*/
//Cool bi
   /* public void importExcelData(InputStream inputStream, String jdbcURL, String username, String password) {
        try (Connection connection = DriverManager.getConnection(jdbcURL, username, password);
             Workbook workbook = new XSSFWorkbook(inputStream)) {

            // Ouvrir la première feuille
            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            int columnCount = headerRow.getPhysicalNumberOfCells();

            // Log des informations sur la feuille
            System.out.println("Nombre de lignes : " + sheet.getPhysicalNumberOfRows());
            System.out.println("Nombre de colonnes : " + columnCount);

            // Obtenir les noms et types de colonnes
            List<String> columns = new ArrayList<>();
            List<String> columnTypes = new ArrayList<>();

            for (int i = 0; i < columnCount; i++) {
                System.out.println("Traitement de la colonne " + i);
                String columnName = headerRow.getCell(i).getStringCellValue().trim();
                columns.add(columnName);

                Row firstDataRow = sheet.getRow(1); // Utilisez une variable pour la ligne de données
                if (firstDataRow != null) {
                    Cell firstDataCell = firstDataRow.getCell(i);
                    if (firstDataCell != null) {
                        CellType cellType = firstDataCell.getCellType();
                        columnTypes.add(getSQLType(cellType));
                    } else {
                        System.out.println("Cellule vide à la ligne 1, colonne " + i);
                        columnTypes.add("VARCHAR(255)"); // Ou un autre type par défaut
                    }
                } else {
                    System.out.println("Ligne 1 est vide.");
                    columnTypes.add("VARCHAR(255)"); // Ou un autre type par défaut
                }
            }

            // Log des colonnes et types détectés
            System.out.println("Colonnes détectées : " + columns);
            System.out.println("Types de colonnes : " + columnTypes);

            // Créer dynamiquement une table si elle n'existe pas
            String tableName = "TableTest";
            createTableIfNotExists(connection, tableName, columns, columnTypes);

            // Insérer les données dans la table
            insertExcelDataIntoTable(connection, sheet, tableName, columns);

            System.out.println("Données insérées avec succès.");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }*/
/*Code bi
   public void importExcelData(InputStream inputStream, String jdbcURL, String username, String password) {
       try (Connection connection = DriverManager.getConnection(jdbcURL, username, password);
            Workbook workbook = new XSSFWorkbook(inputStream)) {

           // Ouvrir la première feuille
           Sheet sheet = workbook.getSheetAt(0);
           Row headerRow = sheet.getRow(0);
           int columnCount = headerRow.getPhysicalNumberOfCells();

           // Obtenir les noms des colonnes et types de colonnes
           List<String> columns = new ArrayList<>();
           List<String> columnTypes = new ArrayList<>();

           for (int i = 0; i < columnCount; i++) {
               String columnName = headerRow.getCell(i).getStringCellValue().trim();
               columns.add(columnName);

               Row firstDataRow = sheet.getRow(1);
               Cell firstDataCell = firstDataRow.getCell(i);
               CellType cellType = firstDataCell != null ? firstDataCell.getCellType() : CellType.STRING;
               columnTypes.add(getSQLType(cellType));
           }

           // Créer la table si elle n'existe pas
           String tableName = "TableTest"; // Nom de la table
           createTableIfNotExists(connection, tableName, columns, columnTypes);

           // Insérer les données
           insertExcelDataIntoTable(connection, sheet, tableName, columns);

       } catch (Exception e) {
           e.printStackTrace();
       }
   }

    // Méthode pour insérer les données dans la table
   /* private void insertExcelDataIntoTable(Connection connection, Sheet sheet, String tableName, List<String> columns) throws SQLException {
        StringBuilder insertQuery = new StringBuilder("INSERT INTO " + tableName + " (");
        for (String column : columns) {
            insertQuery.append(column).append(", ");
        }
        insertQuery.setLength(insertQuery.length() - 2); // Retirer la dernière virgule
        insertQuery.append(") VALUES (");
        for (int i = 0; i < columns.size(); i++) {
            insertQuery.append("?, ");
        }
        insertQuery.setLength(insertQuery.length() - 2); // Retirer la dernière virgule
        insertQuery.append(")");

        try (PreparedStatement preparedStatement = connection.prepareStatement(insertQuery.toString())) {
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    for (int j = 0; j < columns.size(); j++) {
                        Cell cell = row.getCell(j);
                        setPreparedStatementValue(preparedStatement, j + 1, cell);
                    }
                    preparedStatement.addBatch();  // Ajouter la requête à un batch
                }
            }
            preparedStatement.executeBatch(); // Exécuter toutes les requêtes en une seule fois
        }
    }*


    /*public void importExcelData(InputStream inputStream, String jdbcURL, String username, String password, String sourceTable) {
        try (Connection connection = DriverManager.getConnection(jdbcURL, username, password);
             Workbook workbook = new XSSFWorkbook(inputStream)) {

            // Ouvrir la première feuille
            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            int columnCount = headerRow.getPhysicalNumberOfCells();

            // Log des informations sur la feuille
            System.out.println("Nombre de lignes : " + sheet.getPhysicalNumberOfRows());
            System.out.println("Nombre de colonnes : " + columnCount);

            // Obtenir les noms et types de colonnes
            List<String> columns = new ArrayList<>();
            List<String> columnTypes = new ArrayList<>();

            for (int i = 0; i < columnCount; i++) {
                String columnName = headerRow.getCell(i).getStringCellValue().trim();
                columns.add(columnName);

                Row firstDataRow = sheet.getRow(1);
                Cell firstDataCell = firstDataRow.getCell(i);
                CellType cellType = firstDataCell != null ? firstDataCell.getCellType() : CellType.STRING;
                columnTypes.add(getSQLType(cellType));
            }

            // Log des colonnes et types détectés
            System.out.println("Colonnes détectées : " + columns);
            System.out.println("Types de colonnes : " + columnTypes);

            // Créer dynamiquement une table si elle n'existe pas
            String tableName = "TableTest";
            createTableIfNotExists(connection, tableName, columns, columnTypes);

            // Insérer les données depuis la table source dans la table cible
            insertDataFromSourceTable(connection, sourceTable, tableName);

            System.out.println("Données insérées avec succès.");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }*/

/*Code bi
    // Méthode pour insérer les données depuis une table source
    private void insertDataFromSourceTable(Connection connection, String sourceTable, String targetTable) throws SQLException {
        String selectQuery = "SELECT * FROM " + sourceTable;
        String insertQuery = "INSERT INTO " + targetTable + " (column1, column2, ...) VALUES (?, ?, ...)"; // Remplacez `column1`, `column2`, ... par vos noms de colonnes

        try (Statement selectStatement = connection.createStatement();
             ResultSet resultSet = selectStatement.executeQuery(selectQuery);
             PreparedStatement insertStatement = connection.prepareStatement(insertQuery)) {

            while (resultSet.next()) {
                // Assurez-vous que l'ordre des colonnes dans l'INSERT corresponde à celui de votre SELECT
                insertStatement.setString(1, resultSet.getString("column1")); // Remplacez par le bon index et le nom de colonne
                insertStatement.setString(2, resultSet.getString("column2")); // Remplacez par le bon index et le nom de colonne
                // Ajoutez d'autres colonnes selon vos besoins

                insertStatement.addBatch();
            }
            insertStatement.executeBatch(); // Exécutez tous les insertions en une seule fois
        }
    }





    // Méthode pour déterminer le type SQL basé sur le type de cellule Excel
    private String getSQLType(CellType cellType) {
        switch (cellType) {
            case NUMERIC:
                return "DOUBLE";  // Utiliser DOUBLE pour les valeurs numériques
            case BOOLEAN:
                return "BOOLEAN";
            case STRING:
            case FORMULA:
            default:
                return "VARCHAR(255)";  // Forcer VARCHAR pour tout autre type
        }
    }

    // Méthode pour créer la table dynamiquement en fonction des types de colonnes détectés
    /*private void createTableIfNotExists(Connection connection, String tableName, List<String> columns, List<String> columnTypes) throws SQLException {
        StringBuilder createTableQuery = new StringBuilder("CREATE TABLE IF NOT EXISTS " + tableName + " (id INT AUTO_INCREMENT PRIMARY KEY, ");
        for (int i = 0; i < columns.size(); i++) {
            createTableQuery.append(columns.get(i)).append(" ").append(columnTypes.get(i)).append(", ");
        }
        // Retirer la dernière virgule et espace
        createTableQuery.setLength(createTableQuery.length() - 2);
        createTableQuery.append(")");

        try (Statement statement = connection.createStatement()) {
            statement.executeUpdate(createTableQuery.toString());
            System.out.println("Table créée ou déjà existante : " + tableName);
        }
    }*/

/*Code bi
    private void createTableIfNotExists(Connection connection, String tableName, List<String> columns, List<String> columnTypes) throws SQLException {
        StringBuilder createTableQuery = new StringBuilder("CREATE TABLE IF NOT EXISTS " + tableName + " (id INT AUTO_INCREMENT PRIMARY KEY, ");

        for (int i = 0; i < columns.size(); i++) {
            // Vérifiez que le nom de la colonne n'est pas vide avant de l'ajouter à la requête
            if (columns.get(i) != null && !columns.get(i).trim().isEmpty()) {
                createTableQuery.append(columns.get(i)).append(" ").append(columnTypes.get(i)).append(", ");
            }
        }

        // Retirer la dernière virgule et espace si des colonnes ont été ajoutées
        if (createTableQuery.length() > 42) { // 42 est la longueur de "CREATE TABLE IF NOT EXISTS " + tableName + " (id INT AUTO_INCREMENT PRIMARY KEY, "
            createTableQuery.setLength(createTableQuery.length() - 2); // Supprimer la dernière virgule
        }

        createTableQuery.append(") ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_general_ci");

        // Afficher la requête SQL pour déboguer
        System.out.println("Requête SQL de création de table : " + createTableQuery.toString());

        try (Statement statement = connection.createStatement()) {
            statement.executeUpdate(createTableQuery.toString());
            System.out.println("Table créée ou déjà existante : " + tableName);
        }
    }




    // Méthode pour insérer les données dans la table
    private void insertExcelDataIntoTable(Connection connection, Sheet sheet, String tableName, List<String> columns) throws SQLException {
        StringBuilder insertQuery = new StringBuilder("INSERT INTO " + tableName + " (");
        for (String column : columns) {
            insertQuery.append(column).append(", ");
        }
        insertQuery.setLength(insertQuery.length() - 2); // Retirer la dernière virgule
        insertQuery.append(") VALUES (");
        for (int i = 0; i < columns.size(); i++) {
            insertQuery.append("?, ");
        }
        insertQuery.setLength(insertQuery.length() - 2); // Retirer la dernière virgule
        insertQuery.append(")");

        // Log de la requête SQL
        System.out.println("Requête SQL générée : " + insertQuery);

        try (PreparedStatement preparedStatement = connection.prepareStatement(insertQuery.toString())) {
            // Lire les lignes du fichier Excel et insérer les données
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    for (int j = 0; j < columns.size(); j++) {
                        Cell cell = row.getCell(j);
                        setPreparedStatementValue(preparedStatement, j + 1, cell);
                    }
                    preparedStatement.addBatch();  // Ajouter la requête à un batch
                }
            }
            // Exécuter toutes les requêtes en une seule fois
            preparedStatement.executeBatch();
        }
    }



    // Méthode pour définir la valeur appropriée dans le PreparedStatement selon le type de cellule
    private void setPreparedStatementValue(PreparedStatement preparedStatement, int parameterIndex, Cell cell) throws SQLException {
        if (cell == null) {
            preparedStatement.setNull(parameterIndex, Types.VARCHAR); // Par défaut NULL
            System.out.println("Cellule vide à l'index " + parameterIndex);
            return;
        }

        switch (cell.getCellType()) {
            case STRING:
                String cellValue = cell.getStringCellValue();
                // Vérifiez la longueur de la valeur et affichez-la
                System.out.println("Valeur STRING à insérer : " + cellValue);
                if (cellValue.length() > 255) {
                    System.out.println("Avertissement : La valeur de Wara dépasse 255 caractères : " + cellValue);
                }
                preparedStatement.setString(parameterIndex, cellValue);
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    java.sql.Date dateValue = new java.sql.Date(cell.getDateCellValue().getTime());
                    System.out.println("Valeur DATE à insérer : " + dateValue);
                    preparedStatement.setDate(parameterIndex, dateValue);
                } else {
                    double numericValue = cell.getNumericCellValue();
                    System.out.println("Valeur NUMERIC à insérer : " + numericValue);
                    preparedStatement.setDouble(parameterIndex, numericValue);
                }
                break;
            case BOOLEAN:
                boolean booleanValue = cell.getBooleanCellValue();
                System.out.println("Valeur BOOLEAN à insérer : " + booleanValue);
                preparedStatement.setBoolean(parameterIndex, booleanValue);
                break;
            default:
                System.out.println("Valeur par défaut insérée : \"\"");
                preparedStatement.setString(parameterIndex, "");  // Cas des cellules vides ou autres
        }
    }
}
*/


package com.example.demo.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.InputStream;
import java.sql.*;
import java.util.ArrayList;
import java.util.List;

@Service
public class ExcelToXmlService {
    public void importExcelData(InputStream inputStream, String jdbcURL, String username, String password) {
        try (Connection connection = DriverManager.getConnection(jdbcURL, username, password);
             Workbook workbook = new XSSFWorkbook(inputStream)) {

            // Ouvrir la première feuille
            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            int columnCount = headerRow.getPhysicalNumberOfCells();

            // Obtenir les noms des colonnes et types de colonnes
            List<String> columns = new ArrayList<>();
            List<String> columnTypes = new ArrayList<>();

            for (int i = 0; i < columnCount; i++) {
                String columnName = headerRow.getCell(i).getStringCellValue().trim();
                columns.add(columnName);

                Row firstDataRow = sheet.getRow(1);
                Cell firstDataCell = firstDataRow.getCell(i);
                CellType cellType = firstDataCell != null ? firstDataCell.getCellType() : CellType.STRING;
                columnTypes.add(getSQLType(cellType));
            }

            // Créer la table si elle n'existe pas
            String tableName = "TableTest"; // Nom de la table
            createTableIfNotExists(connection, tableName, columns, columnTypes);

            // Vérifiez et ajoutez les colonnes si nécessaire
            for (int i = 0; i < columns.size(); i++) {
                String columnName = columns.get(i);
                String columnType = columnTypes.get(i);
                addColumnIfNotExists(connection, tableName, columnName, columnType);
            }

            // Insérer les données
            insertExcelDataIntoTable(connection, sheet, tableName, columns);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // Méthode pour vérifier et ajouter une colonne si elle n'existe pas
    private void addColumnIfNotExists(Connection connection, String tableName, String columnName, String columnType) throws SQLException {
        String checkColumnQuery = "SHOW COLUMNS FROM " + tableName + " LIKE ?";
        try (PreparedStatement checkStatement = connection.prepareStatement(checkColumnQuery)) {
            checkStatement.setString(1, columnName);
            ResultSet resultSet = checkStatement.executeQuery();
            if (!resultSet.next()) { // Colonne n'existe pas
                String addColumnQuery = "ALTER TABLE " + tableName + " ADD COLUMN " + columnName + " " + columnType;
                try (Statement statement = connection.createStatement()) {
                    statement.executeUpdate(addColumnQuery);
                    System.out.println("Colonne ajoutée : " + columnName);
                }
            } else {
                System.out.println("La colonne existe déjà : " + columnName);
            }
        }
    }

    // Méthode pour insérer les données dans la table
    private void insertExcelDataIntoTable(Connection connection, Sheet sheet, String tableName, List<String> columns) throws SQLException {
        StringBuilder insertQuery = new StringBuilder("INSERT INTO " + tableName + " (");
        for (String column : columns) {
            insertQuery.append(column).append(", ");
        }
        insertQuery.setLength(insertQuery.length() - 2); // Retirer la dernière virgule
        insertQuery.append(") VALUES (");
        for (int i = 0; i < columns.size(); i++) {
            insertQuery.append("?, ");
        }
        insertQuery.setLength(insertQuery.length() - 2); // Retirer la dernière virgule
        insertQuery.append(")");

        // Log de la requête SQL
        System.out.println("Requête SQL générée : " + insertQuery);

        try (PreparedStatement preparedStatement = connection.prepareStatement(insertQuery.toString())) {
            // Lire les lignes du fichier Excel et insérer les données
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    for (int j = 0; j < columns.size(); j++) {
                        Cell cell = row.getCell(j);
                        setPreparedStatementValue(preparedStatement, j + 1, cell);
                    }
                    preparedStatement.addBatch();  // Ajouter la requête à un batch
                }
            }
            // Exécuter toutes les requêtes en une seule fois
            preparedStatement.executeBatch();
        }
    }

    // Méthode pour définir la valeur appropriée dans le PreparedStatement selon le type de cellule
    private void setPreparedStatementValue(PreparedStatement preparedStatement, int parameterIndex, Cell cell) throws SQLException {
        if (cell == null) {
            preparedStatement.setNull(parameterIndex, Types.VARCHAR); // Par défaut NULL
            System.out.println("Cellule vide à l'index " + parameterIndex);
            return;
        }

        switch (cell.getCellType()) {
            case STRING:
                String cellValue = cell.getStringCellValue();
                // Vérifiez la longueur de la valeur et affichez-la
                System.out.println("Valeur STRING à insérer : " + cellValue);
                if (cellValue.length() > 255) {
                    System.out.println("Avertissement : La valeur de Wara dépasse 255 caractères : " + cellValue);
                }
                preparedStatement.setString(parameterIndex, cellValue);
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    java.sql.Date dateValue = new java.sql.Date(cell.getDateCellValue().getTime());
                    System.out.println("Valeur DATE à insérer : " + dateValue);
                    preparedStatement.setDate(parameterIndex, dateValue);
                } else {
                    double numericValue = cell.getNumericCellValue();
                    System.out.println("Valeur NUMERIC à insérer : " + numericValue);
                    preparedStatement.setDouble(parameterIndex, numericValue);
                }
                break;
            case BOOLEAN:
                boolean booleanValue = cell.getBooleanCellValue();
                System.out.println("Valeur BOOLEAN à insérer : " + booleanValue);
                preparedStatement.setBoolean(parameterIndex, booleanValue);
                break;
            default:
                System.out.println("Valeur par défaut insérée : \"\"");
                preparedStatement.setString(parameterIndex, "");  // Cas des cellules vides ou autres
        }
    }

    // Méthode pour déterminer le type SQL basé sur le type de cellule Excel
    private String getSQLType(CellType cellType) {
        switch (cellType) {
            case NUMERIC:
                return "DOUBLE";  // Utiliser DOUBLE pour les valeurs numériques
            case BOOLEAN:
                return "BOOLEAN";
            case STRING:
            case FORMULA:
            default:
                return "VARCHAR(255)";  // Forcer VARCHAR pour tout autre type
        }
    }

    // Méthode pour créer la table dynamiquement en fonction des types de colonnes détectés
    private void createTableIfNotExists(Connection connection, String tableName, List<String> columns, List<String> columnTypes) throws SQLException {
        StringBuilder createTableQuery = new StringBuilder("CREATE TABLE IF NOT EXISTS " + tableName + " (id INT AUTO_INCREMENT PRIMARY KEY, ");
        for (int i = 0; i < columns.size(); i++) {
            if (columns.get(i) != null && !columns.get(i).trim().isEmpty()) {
                createTableQuery.append(columns.get(i)).append(" ").append(columnTypes.get(i)).append(", ");
            }
        }
        if (createTableQuery.length() > 42) { // Vérifiez que des colonnes ont été ajoutées
            createTableQuery.setLength(createTableQuery.length() - 2); // Supprimer la dernière virgule
        }
        createTableQuery.append(") ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_general_ci");

        // Afficher la requête SQL pour déboguer
        System.out.println("Requête SQL de création de table : " + createTableQuery.toString());

        try (Statement statement = connection.createStatement()) {
            statement.executeUpdate(createTableQuery.toString());
            System.out.println("Table créée ou déjà existante : " + tableName);
        }
    }
}
