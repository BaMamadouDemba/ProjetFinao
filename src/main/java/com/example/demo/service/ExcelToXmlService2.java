/*package com.example.demo.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import javax.xml.stream.XMLOutputFactory;
import javax.xml.stream.XMLStreamWriter;
import java.io.FileOutputStream;
import java.io.InputStream;

@Service
public class ExcelToXmlService {

    public String convertExcelToXml(InputStream inputStream) {
        try (Workbook workbook = new XSSFWorkbook(inputStream)) {

            /*String userHome = System.getProperty("user.home");
            String desktopPath = userHome + "/Desktop/output.xml"; */// chemin complet vers le bureau

         /*   String desktopPath = "/Users/mamadou/Desktop/output.xml";
            System.out.println("Chemin du fichier XML : " + desktopPath);

            // Ouvrir la première feuill
          /*  Sheet sheet = workbook.getSheetAt(0);
            FileOutputStream fos = new FileOutputStream("output.xml");*/

            // Ouvrir la première feuille
       /*     Sheet sheet = workbook.getSheetAt(0);
            FileOutputStream fos = new FileOutputStream(desktopPath);

            // Initialiser le writer XML
            XMLOutputFactory xmlOutputFactory = XMLOutputFactory.newInstance();
            XMLStreamWriter xmlWriter = xmlOutputFactory.createXMLStreamWriter(fos, "UTF-8");

            // Commencer à écrire le fichier XML
            xmlWriter.writeStartDocument();
            xmlWriter.writeStartElement("Batch");

            // Parcourir les lignes du fichier Excel
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row != null && "No".equalsIgnoreCase(getCellValue(row.getCell(36)))) {
                    xmlWriter.writeStartElement("Contract");

                    xmlWriter.writeStartElement("ContractCode");
                    xmlWriter.writeCharacters(getCellValue(row.getCell(0)));  // Colonne 1
                    xmlWriter.writeEndElement();

                    xmlWriter.writeStartElement("ContractData");

                    xmlWriter.writeStartElement("Consented");
                    xmlWriter.writeCharacters(getCellValue(row.getCell(36)));  // Colonne 37
                    xmlWriter.writeEndElement();

                    xmlWriter.writeStartElement("Branch");
                    xmlWriter.writeCharacters(getCellValue(row.getCell(2)));  // Colonne 3
                    xmlWriter.writeEndElement();

                    xmlWriter.writeStartElement("PhaseOfContract");
                    xmlWriter.writeCharacters(getCellValue(row.getCell(3)));  // Colonne 4
                    xmlWriter.writeEndElement();

                    xmlWriter.writeStartElement("ContractStatus");
                    xmlWriter.writeCharacters(getCellValue(row.getCell(4)));  // Colonne 5
                    xmlWriter.writeEndElement();

                    // Fin de ContractData
                    xmlWriter.writeEndElement();
                    // Fin de Contract
                    xmlWriter.writeEndElement();
                }
            }

            // Fermer l'élément Batch
            xmlWriter.writeEndElement();
            xmlWriter.writeEndDocument();

            // Fermer le writer et le flux de sortie
            xmlWriter.close();
            fos.close();

            return "Exportation terminée. Le fichier XML a été généré.";

        } catch (Exception e) {
            e.printStackTrace();
            return "Erreur lors de l'exportation : " + e.getMessage();
        }
    }

    // Méthode utilitaire pour récupérer la valeur d'une cellule Excel
    private String getCellValue(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();  // Formatez la date si nécessaire
                } else {
                    return String.valueOf(cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            default:
                return "";
        }
    }
}
*/