//AJOUT
/*
package com.example.demo.service;

import com.example.demo.entity.DataRecord;
import com.example.demo.entity.DynamicColumn;
import com.example.demo.repository.DynamicDataRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;

    @Service
    public class ConvertionXmlService {

        @Autowired
        private DynamicDataRepository dynamicDataRepository;

        public String convertDataToXml() {
            List<DataRecord> dataRecords = dynamicDataRepository.findAll();
            StringBuilder xmlBuilder = new StringBuilder();

            // Début du document XML
            xmlBuilder.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
            xmlBuilder.append("<root>\n");

            // Vérifier s'il y a des enregistrements pour éviter les erreurs
            if (!dataRecords.isEmpty()) {
                DataRecord firstRecord = dataRecords.get(0);

                // Vérifier si `firstRecord` et ses colonnes ne sont pas null
                if (firstRecord != null && firstRecord.getColumns() != null) {
                    for (DynamicColumn column : firstRecord.getColumns()) {
                        if (column.getColumnName() != null) { // Vérifiez si le nom de la colonne n'est pas null
                            xmlBuilder.append("  <").append(column.getColumnName()).append(">\n");

                            for (DataRecord record : dataRecords) {
                                // Assurez-vous que `record` et ses colonnes ne sont pas null
                                if (record != null && record.getColumns() != null) {
                                    String value = record.getColumns().stream()
                                            .filter(c -> column.getColumnName().equals(c.getColumnName()))
                                            .map(DynamicColumn::getColumnValue)
                                            .findFirst()
                                            .orElse(""); // Valeur par défaut vide si non trouvée

                                    xmlBuilder.append("    <value>").append(value != null ? value : "").append("</value>\n");
                                }
                            }
                            xmlBuilder.append("  </").append(column.getColumnName()).append(">\n");
                        }
                    }
                }
            }

            xmlBuilder.append("</root>"); // Fin du document XML
            return xmlBuilder.toString(); // Retourne la chaîne XML formatée
        }
    }

*/

package com.example.demo.service;
import com.example.demo.entity.DataRecord;
import com.example.demo.entity.DynamicColumn;
import com.example.demo.repository.DynamicDataRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import javax.xml.stream.XMLOutputFactory;
import javax.xml.stream.XMLStreamWriter;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.stream.StreamResult;
import javax.xml.transform.stream.StreamSource;

import java.io.FileOutputStream;
import java.io.StringReader;
import java.io.StringWriter;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;

@Service
public class ConvertionXmlService {

    @Autowired
    private DynamicDataRepository dynamicDataRepository;


    public String convertDataToXml() {
        try {
            // Récupération des données dynamiques depuis la base de données
            List<DataRecord> dataRecords = dynamicDataRepository.findAll();

            // Récupérer l'heure actuelle pour générer un nom de fichier unique
            LocalDateTime now = LocalDateTime.now();
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss");
            String formattedDate = now.format(formatter);
            String desktopPath = System.getProperty("user.home") + "/Desktop/output_" + formattedDate + ".xml";

            // Préparation pour écrire le contenu XML
            StringWriter stringWriter = new StringWriter();
            XMLOutputFactory xmlOutputFactory = XMLOutputFactory.newInstance();
            XMLStreamWriter xmlWriter = xmlOutputFactory.createXMLStreamWriter(stringWriter);

            // Début du document XML
            xmlWriter.writeStartDocument("UTF-8", "1.0");
            xmlWriter.writeStartElement("root");
            xmlWriter.writeNamespace("fin", "http://mabanque.com/schemas/finances/WestAfrica " + " " + formattedDate);


            // Parcourir chaque `DataRecord` et écrire ses données dans le fichier XML
            for (DataRecord record : dataRecords) {
                xmlWriter.writeStartElement("Record"); // Nouvel enregistrement

                // Ajouter chaque colonne dynamique de l'enregistrement comme élément XML
                for (DynamicColumn column : record.getColumns()) {
                    if (column.getColumnName() != null && column.getColumnValue() != null) {
                        xmlWriter.writeStartElement(sanitizeXmlElementName(column.getColumnName())); // Nom de la colonne
                        xmlWriter.writeCharacters(column.getColumnValue()); // Valeur de la colonne
                        xmlWriter.writeEndElement(); // Fin de la balise de colonne
                    }
                }

                xmlWriter.writeEndElement(); // Fin de <Record>
            }

            xmlWriter.writeEndElement(); // Fin de <root>
            xmlWriter.writeEndDocument();
            xmlWriter.close();

            // Transformer pour indenter le XML
            TransformerFactory transformerFactory = TransformerFactory.newInstance();
            Transformer transformer = transformerFactory.newTransformer();
            transformer.setOutputProperty(OutputKeys.INDENT, "yes");
            transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "4");

            // Lire le contenu XML non formaté et le transformer en XML formaté
            StreamSource xmlInput = new StreamSource(new StringReader(stringWriter.toString()));
            StringWriter formattedWriter = new StringWriter();
            transformer.transform(xmlInput, new StreamResult(formattedWriter));

            // Sauvegarder le fichier XML formaté
            try (FileOutputStream fos = new FileOutputStream(desktopPath)) {
                fos.write(formattedWriter.toString().getBytes());
            }

            return "Le fichier XML a été généré avec succès sur le Bureau : " + desktopPath;

        } catch (Exception e) {
            e.printStackTrace();
            return "Erreur lors de la conversion en XML : " + e.getMessage();
        }
    }

    // Méthode utilitaire pour nettoyer les noms d'éléments XML
    private String sanitizeXmlElementName(String name) {
        return name.replaceAll("[^a-zA-Z0-9_]", "_"); // Remplace les caractères non valides par un underscore
    }
}
