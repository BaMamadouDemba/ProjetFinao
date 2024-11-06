package com.example.demo.controller;

import com.example.demo.service.ExcelFileService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

@RestController
public class ExcelFileController {

    @Autowired
    private ExcelFileService excelFileService;

    @PostMapping("/import-excel")
    public ResponseEntity<String> importExcel(@RequestParam("file") MultipartFile file) {
        excelFileService.readExcelFile(file);
        return ResponseEntity.ok("Fichier Excel importé avec succès !");
    }
}
