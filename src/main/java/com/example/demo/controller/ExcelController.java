//CODE
package com.example.demo.controller;

import com.example.demo.entity.DataRecord;
import com.example.demo.service.ExcelToXmlService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;

@CrossOrigin(origins = "http://localhost:8080")
@RestController
@RequestMapping("/api/excel")
public class ExcelController {

    private static final Logger logger = LoggerFactory.getLogger(ExcelController.class);
    private final ExcelToXmlService excelToXmlService;

    public ExcelController(ExcelToXmlService excelToXmlService) {
        this.excelToXmlService = excelToXmlService;
    }

    @PostMapping("/import")
    public ResponseEntity<String> importExcelFile(@RequestParam("file") MultipartFile file) {
        if (file.isEmpty()) {
            return new ResponseEntity<>("Le fichier est vide", HttpStatus.BAD_REQUEST);
        }

        try (InputStream inputStream = file.getInputStream()) {
            // Remplacez par votre propre configuration de base de données
            String jdbcURL = "jdbc:mysql://localhost:3306/BaseTest";
            String username = "root";
            String password = "root";

            // Appel du service pour importer les données
            excelToXmlService.importExcelData(inputStream, jdbcURL, username, password);

            return new ResponseEntity<>("Fichier Excel importé avec succès.", HttpStatus.OK);
        } catch (IOException e) {
            return new ResponseEntity<>("Erreur lors de la lecture du fichier.", HttpStatus.INTERNAL_SERVER_ERROR);
        } catch (Exception e) {
            logger.error("Erreur lors de l'importation des données : {}", e.getMessage());
            return new ResponseEntity<>("Erreur lors de l'importation des données.", HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }

    // Endpoint pour récupérer les données
    @GetMapping("/data")
    public ResponseEntity<List<DataRecord>> getDataRecords() {
        try {
            List<DataRecord> records = excelToXmlService.getDataRecords();
            if (records.isEmpty()) {
                logger.warn("Aucune donnée trouvée dans la base de données.");
                return new ResponseEntity<>(HttpStatus.NO_CONTENT);
            }
            return new ResponseEntity<>(records, HttpStatus.OK);
        } catch (Exception e) {
            logger.error("Erreur lors de la récupération des données : {}", e.getMessage(), e);
            return new ResponseEntity<>(HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }
}
