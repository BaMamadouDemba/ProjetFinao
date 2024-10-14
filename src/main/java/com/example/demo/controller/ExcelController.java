/*
package com.example.demo.controller;

import com.example.demo.entity.ExcelData;
import com.example.demo.service.ExcelToXmlService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.InputStream;

@RestController
@RequestMapping("/api/excel")
@CrossOrigin(origins = "http://localhost:8080")
public class ExcelController {

    @Autowired
    private ExcelToXmlService excelToXmlService;

    @PostMapping("/import")
    public String importExcelData(@RequestParam("file") MultipartFile file) {
        try (InputStream inputStream = file.getInputStream()) {
            String fileName = file.getOriginalFilename();
            return excelToXmlService.importExcelData(inputStream, fileName);
        } catch (Exception e) {
            e.printStackTrace();
            return "Erreur lors du traitement du fichier : " + e.getMessage();
        }
    }
//adpoint ajouter
   /* @GetMapping("/data/{id}")
    public ResponseEntity<ExcelData> getExcelData(@PathVariable Long id) {
        ExcelData excelData = excelToXmlService.getExcelDataById(id);

        if (excelData == null) {
            return ResponseEntity.notFound().build(); // Si l'entrée n'existe pas
        }

        return ResponseEntity.ok(excelData); // Renvoie les détails de l'entrée
    }
}

*/

/*package com.example.demo.controller;

import com.example.demo.service.ExcelToXmlService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;

@RestController
public class ExcelController {

    @Autowired
    private ExcelToXmlService excelToXmlService;

    @PostMapping("/import")
    public String importExcel(@RequestParam("file") MultipartFile file) {
        String jdbcURL = "jdbc:mysql://localhost:3306/excel_data_db"; // Spécifiez votre URL de base de données
        String username = "root"; // Votre nom d'utilisateur MySQL
        String password = "root"; // Votre mot de passe MySQL

        try (InputStream inputStream = file.getInputStream()) {
            excelToXmlService.importExcelData(inputStream, jdbcURL, username, password);
            return "Données importées avec succès";
        } catch (IOException e) {
            e.printStackTrace();
            return "Erreur lors de l'importation des données";
        }
    }
}
*/

package com.example.demo.controller;

import com.example.demo.service.ExcelToXmlService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;

@RestController
@RequestMapping("/api/excel")
public class ExcelController {

    @Autowired
    private ExcelToXmlService excelToXmlService;

    // Endpoint pour importer un fichier Excel
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

            // Appel du service pour importer les données du fichier Excel
            excelToXmlService.importExcelData(inputStream, jdbcURL, username, password);

            return new ResponseEntity<>("Fichier Excel importé avec succès.", HttpStatus.OK);
        } catch (IOException e) {
            return new ResponseEntity<>("Erreur lors de la lecture du fichier.", HttpStatus.INTERNAL_SERVER_ERROR);
        } catch (Exception e) {
            return new ResponseEntity<>("Erreur lors de l'importation des données : " + e.getMessage(), HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }
}


