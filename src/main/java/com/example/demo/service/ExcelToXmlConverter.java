/*package com.example.demo.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import javax.xml.stream.XMLOutputFactory;
import javax.xml.stream.XMLStreamWriter;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.stream.StreamResult;
import javax.xml.transform.stream.StreamSource;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.StringReader;
import java.io.StringWriter;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

@Service
public class ExcelToXmlConverter {

    public String convertExcelToXml(InputStream inputStream) {
        try (Workbook workbook = new XSSFWorkbook(inputStream)) {
            // Récupérer l'heure actuelle et la formater
            LocalDateTime now = LocalDateTime.now();
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss");
            String formattedDate = now.format(formatter);

            // Chemins des fichiers avec horodatage
            String desktopPath = System.getProperty("user.home") + "/Desktop/output_" + formattedDate + ".xml";
            String xsdPath = System.getProperty("user.home") + "/Desktop/output_" + formattedDate + ".xsd"; // Chemin pour le fichier XSD

            // Ouvrir la première feuille
            Sheet sheet = workbook.getSheetAt(0);
            String[] columnNames = new String[sheet.getRow(0).getLastCellNum()];

            // Remplir le tableau avec les noms de colonnes
            Row headerRow = sheet.getRow(0);
            for (int j = 0; j < headerRow.getLastCellNum(); j++) {
                Cell headerCell = headerRow.getCell(j);
                columnNames[j] = sanitizeXmlElementName(getCellValue(headerCell));
            }

            // Générer le fichier XSD
            generateXsd(xsdPath, columnNames);

            StringWriter stringWriter = new StringWriter();
            XMLOutputFactory xmlOutputFactory = XMLOutputFactory.newInstance();
            XMLStreamWriter xmlWriter = xmlOutputFactory.createXMLStreamWriter(stringWriter);

            // Écrire le début du document avec les namespaces
            xmlWriter.writeStartDocument("UTF-8", "1.0");
            xmlWriter.writeStartElement("Batch");
            xmlWriter.writeAttribute("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");
            xmlWriter.writeAttribute("xmlns:xsd", "http://www.w3.org/2001/XMLSchema");
            xmlWriter.writeAttribute("xmlns", "http://creditinfo.com/schemas/CB5/WestAfrica/contract");

            // Parcourir les lignes du fichier Excel à partir de la deuxième ligne
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                // Vérifiez si la ligne n'est pas nulle et contient au moins une cellule non vide
                if (row != null && !isRowEmpty(row)) {
                    xmlWriter.writeStartElement("Contract"); // Début de l'élément Contract

                    // Ajouter les éléments à partir des colonnes
                    for (int col = 0; col < columnNames.length; col++) {
                        Cell cell = row.getCell(col);
                        String cellValue = getCellValue(cell);
                        if (!cellValue.isEmpty()) {
                            String elementName = sanitizeXmlElementName(columnNames[col]); // Nettoyer le nom de l'élément
                            xmlWriter.writeStartElement(elementName);
                            xmlWriter.writeCharacters(cellValue);
                            xmlWriter.writeEndElement(); // Fin de l'élément
                        }
                    }

                    xmlWriter.writeEndElement(); // Fin de Contract
                }
            }

            xmlWriter.writeEndElement(); // Fin de Batch
            xmlWriter.writeEndDocument();
            xmlWriter.close();

            // Formater le XML avec une indentation
            TransformerFactory transformerFactory = TransformerFactory.newInstance();
            Transformer transformer = transformerFactory.newTransformer();
            transformer.setOutputProperty(OutputKeys.INDENT, "yes");
            transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "4");

            // Transformer le contenu écrit en XML formaté
            StreamSource xmlInput = new StreamSource(new StringReader(stringWriter.toString()));
            StringWriter formattedXmlWriter = new StringWriter();
            StreamResult xmlOutput = new StreamResult(formattedXmlWriter);
            transformer.transform(xmlInput, xmlOutput);

            // Sauvegarder le fichier formaté
            try (FileOutputStream fos = new FileOutputStream(desktopPath)) {
                fos.write(formattedXmlWriter.toString().getBytes());
            }

            return "Exportation terminée. Le fichier XML et le fichier XSD ont été générés.";

        } catch (Exception e) {
            e.printStackTrace();
            return "Erreur lors de l'exportation : " + e.getMessage();
        }
    }

    // Autres méthodes inchangées...

    private String getCellValue(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();  // Formatez la date si nécessaire
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

    // Générer le fichier XSD à partir des noms de colonnes
    public void generateXsd(String xsdPath, String[] columnNames) {
        try (FileOutputStream fos = new FileOutputStream(xsdPath)) {
            StringBuilder xsdContent = new StringBuilder();
            xsdContent.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n")
                    .append("<xs:schema xmlns:xs=\"http://www.w3.org/2001/XMLSchema\">\n")
                    .append("    <xs:element name=\"Batch\">\n")
                    .append("        <xs:complexType>\n")
                    .append("            <xs:sequence>\n")
                    .append("                <xs:element name=\"Contract\" maxOccurs=\"unbounded\">\n")
                    .append("                    <xs:complexType>\n")
                    .append("                        <xs:sequence>\n");

            for (String columnName : columnNames) {
                xsdContent.append("                            <xs:element name=\"")
                        .append(columnName)
                        .append("\" type=\"xs:string\"/>\n");
            }

            xsdContent.append("                        </xs:sequence>\n")
                    .append("                    </xs:complexType>\n")
                    .append("                </xs:element>\n")
                    .append("            </xs:sequence>\n")
                    .append("        </xs:complexType>\n")
                    .append("    </xs:element>\n")
                    .append("</xs:schema>\n");

            fos.write(xsdContent.toString().getBytes());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // Méthode utilitaire pour nettoyer les noms d'éléments XML
    private String sanitizeXmlElementName(String name) {
        return name.replaceAll("[^a-zA-Z0-9_]", "_"); // Remplacez les caractères non valides par un underscore
    }
}

*/
/*
package com.example.demo.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import javax.xml.stream.XMLOutputFactory;
import javax.xml.stream.XMLStreamWriter;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.stream.StreamResult;
import javax.xml.transform.stream.StreamSource;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.StringReader;
import java.io.StringWriter;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

@Service
public class ExcelToXmlConverter {

    public String convertExcelToXml(InputStream inputStream) {
        try (Workbook workbook = new XSSFWorkbook(inputStream)) {
            // Récupérer l'heure actuelle et la formater
            LocalDateTime now = LocalDateTime.now();
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss");
            String formattedDate = now.format(formatter);

            // Chemins des fichiers avec horodatage
            String desktopPath = System.getProperty("user.home") + "/Desktop/output_" + formattedDate + ".xml";
            String xsdPath = System.getProperty("user.home") + "/Desktop/output_" + formattedDate + ".xsd"; // Chemin pour le fichier XSD

            // Ouvrir la première feuille
            Sheet sheet = workbook.getSheetAt(0);
            String[] columnNames = new String[sheet.getRow(0).getLastCellNum()];

            // Remplir le tableau avec les noms de colonnes
            Row headerRow = sheet.getRow(0);
            for (int j = 0; j < headerRow.getLastCellNum(); j++) {
                Cell headerCell = headerRow.getCell(j);
                columnNames[j] = sanitizeXmlElementName(getCellValue(headerCell));
            }

            // Générer le fichier XSD
            generateXsd(xsdPath, columnNames);

            StringWriter stringWriter = new StringWriter();
            XMLOutputFactory xmlOutputFactory = XMLOutputFactory.newInstance();
            XMLStreamWriter xmlWriter = xmlOutputFactory.createXMLStreamWriter(stringWriter);

            // Écrire le début du document avec les namespaces
            xmlWriter.writeStartDocument("UTF-8", "1.0");
            xmlWriter.writeStartElement("Batch");
            xmlWriter.writeAttribute("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");
            xmlWriter.writeAttribute("xmlns:xsd", "http://www.w3.org/2001/XMLSchema");
            xmlWriter.writeAttribute("xmlns", "http://creditinfo.com/schemas/CB5/WestAfrica/contract");

            // Parcourir les lignes du fichier Excel à partir de la deuxième ligne
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                // Vérifiez si la ligne n'est pas nulle et contient au moins une cellule non vide
                if (row != null && !isRowEmpty(row)) {
                    // Ajouter les éléments à partir des colonnes directement sous l'élément racine
                    for (int col = 0; col < columnNames.length; col++) {
                        Cell cell = row.getCell(col);
                        String cellValue = getCellValue(cell);
                        if (!cellValue.isEmpty()) {
                            String elementName = sanitizeXmlElementName(columnNames[col]); // Nettoyer le nom de l'élément
                            xmlWriter.writeStartElement(elementName);
                            xmlWriter.writeCharacters(cellValue);
                            xmlWriter.writeEndElement(); // Fin de l'élément
                        }
                    }
                }
            }

            xmlWriter.writeEndElement(); // Fin de Batch
            xmlWriter.writeEndDocument();
            xmlWriter.close();

            // Formater le XML avec une indentation
            TransformerFactory transformerFactory = TransformerFactory.newInstance();
            Transformer transformer = transformerFactory.newTransformer();
            transformer.setOutputProperty(OutputKeys.INDENT, "yes");
            transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "4");

            // Transformer le contenu écrit en XML formaté
            StreamSource xmlInput = new StreamSource(new StringReader(stringWriter.toString()));
            StringWriter formattedXmlWriter = new StringWriter();
            StreamResult xmlOutput = new StreamResult(formattedXmlWriter);
            transformer.transform(xmlInput, xmlOutput);

            // Sauvegarder le fichier formaté
            try (FileOutputStream fos = new FileOutputStream(desktopPath)) {
                fos.write(formattedXmlWriter.toString().getBytes());
            }

            return "Exportation terminée. Le fichier XML et le fichier XSD ont été générés.";

        } catch (Exception e) {
            e.printStackTrace();
            return "Erreur lors de l'exportation : " + e.getMessage();
        }
    }

    // Autres méthodes inchangées...

    private String getCellValue(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();  // Formatez la date si nécessaire
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

    // Générer le fichier XSD à partir des noms de colonnes
    public void generateXsd(String xsdPath, String[] columnNames) {
        try (FileOutputStream fos = new FileOutputStream(xsdPath)) {
            StringBuilder xsdContent = new StringBuilder();
            xsdContent.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n")
                    .append("<xs:schema xmlns:xs=\"http://www.w3.org/2001/XMLSchema\">\n")
                    .append("    <xs:element name=\"Batch\">\n")
                    .append("        <xs:complexType>\n")
                    .append("            <xs:sequence>\n");

            for (String columnName : columnNames) {
                xsdContent.append("                <xs:element name=\"")
                        .append(columnName)
                        .append("\" type=\"xs:string\"/>\n");
            }

            xsdContent.append("            </xs:sequence>\n")
                    .append("        </xs:complexType>\n")
                    .append("    </xs:element>\n")
                    .append("</xs:schema>\n");

            fos.write(xsdContent.toString().getBytes());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // Méthode utilitaire pour nettoyer les noms d'éléments XML
    private String sanitizeXmlElementName(String name) {
        return name.replaceAll("[^a-zA-Z0-9_]", "_"); // Remplacez les caractères non valides par un underscore
    }
}
*/