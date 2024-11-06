//CODE
package com.example.demo.service;

import com.example.demo.entity.DataRecord;
import com.example.demo.entity.DynamicColumn;
import com.example.demo.repository.DataRecordRepository;
import org.apache.poi.ss.usermodel.*;
        import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.io.InputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.util.ArrayList;
import java.util.List;

@Service
public class ExcelToXmlService {

    private static final Logger logger = LoggerFactory.getLogger(ExcelToXmlService.class);

    @Autowired
    private DataRecordRepository dataRecordRepository;

    @Transactional
    public void importExcelData(InputStream inputStream, String jdbcURL, String username, String password) {
        try (Connection connection = DriverManager.getConnection(jdbcURL, username, password);
             Workbook workbook = new XSSFWorkbook(inputStream)) {

            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            int columnCount = headerRow.getPhysicalNumberOfCells();
            List<String> columns = new ArrayList<>();
            List<String> columnTypes = new ArrayList<>();

            // Récupération des noms de colonnes et types
            for (int i = 0; i < columnCount; i++) {
                String columnName = headerRow.getCell(i).getStringCellValue().trim();
                columns.add(columnName);
                Row firstDataRow = sheet.getRow(1);
                Cell firstDataCell = firstDataRow.getCell(i);
                CellType cellType = firstDataCell != null ? firstDataCell.getCellType() : CellType.STRING;
                columnTypes.add(getSQLType(cellType));
            }

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    DataRecord dataRecord = new DataRecord();
                    List<DynamicColumn> dynamicColumns = new ArrayList<>();

                    for (int j = 0; j < columns.size(); j++) {
                        Cell cell = row.getCell(j);
                        String cellValue = getCellValue(cell);

                        // Ne pas ajouter de colonnes vides
                        if (cellValue != null && !cellValue.isEmpty()) {
                            DynamicColumn dynamicColumn = new DynamicColumn();
                            dynamicColumn.setColumnName(columns.get(j));
                            dynamicColumn.setColumnValue(cellValue);
                            dynamicColumn.setDataRecord(dataRecord);
                            dynamicColumns.add(dynamicColumn);
                        }
                    }

                    // Vérifiez si des colonnes dynamiques valides ont été ajoutées
                    if (!dynamicColumns.isEmpty()) {
                        dataRecord.setColumns(dynamicColumns); // Assurez-vous d'utiliser le bon setter
                        dataRecordRepository.save(dataRecord);
                    } else {
                        logger.warn("Aucun champ valide trouvé dans la ligne {}", i);
                    }
                }
            }

        } catch (Exception e) {
            logger.error("Erreur lors de l'importation des données Excel : {}", e.getMessage());
            e.printStackTrace();
        }
    }

    private String getCellValue(Cell cell) {
        if (cell == null) return null;

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return new java.sql.Date(cell.getDateCellValue().getTime()).toString();
                } else {
                    return String.valueOf(cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            default:
                return null;
        }
    }

    public List<DataRecord> getDataRecords() {
        List<DataRecord> records = dataRecordRepository.findAll();
        logger.info("Nombre de données récupérées : {}", records.size());
        return records;
    }

    private String getSQLType(CellType cellType) {
        switch (cellType) {
            case NUMERIC:
                return "DOUBLE";
            case BOOLEAN:
                return "BOOLEAN";
            case STRING:
            case FORMULA:
            default:
                return "VARCHAR(255)";
        }
    }
}
