
/*package com.example.demo.controller;

import com.example.demo.entity.ExcelData;
import com.example.demo.service.ExcelDataService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.util.List;

@RestController
@RequestMapping("/api/excel/data")
public class ExcelDataController {

    @Autowired
    private ExcelDataService excelDataService;

    @GetMapping
    public List<ExcelData> getAllData() {
        return excelDataService.getAllData();
    }

    @GetMapping("/{id}")
    public ResponseEntity<ExcelData> getDataById(@PathVariable Long id) {
        ExcelData data = excelDataService.getDataById(id);
        if (data != null) {
            return ResponseEntity.ok(data);
        }
        return ResponseEntity.notFound().build();
    }

    @PostMapping
    public ExcelData addData(@RequestBody ExcelData data) {
        return excelDataService.addData(data);
    }

    @PutMapping("/{id}")
    public ResponseEntity<ExcelData> updateData(@PathVariable Long id, @RequestBody ExcelData data) {
        ExcelData updatedData = excelDataService.updateData(id, data);
        if (updatedData != null) {
            return ResponseEntity.ok(updatedData);
        }
        return ResponseEntity.notFound().build();
    }

    @DeleteMapping("/{id}")
    public ResponseEntity<Void> deleteData(@PathVariable Long id) {
        excelDataService.deleteData(id);
        return ResponseEntity.noContent().build();
    }
}
*/