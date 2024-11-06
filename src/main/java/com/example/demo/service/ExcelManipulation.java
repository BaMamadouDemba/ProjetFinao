
/*package com.example.demo.service;

import com.example.demo.entity.ExcelData;
import com.example.demo.repository.ExcelDataRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;

@Service
public class ExcelManipulation {

    @Autowired
    private ExcelDataRepository excelDataRepository;

    // Récupérer toutes les données
    public List<ExcelData> getAllData() {
        return excelDataRepository.findAll();
    }

    // Récupérer une donnée par ID
    public ExcelData getDataById(Long id) {
        return excelDataRepository.findById(id).orElse(null);
    }

    // Ajouter une nouvelle donnée
    public ExcelData addData(ExcelData data) {
        return excelDataRepository.save(data);
    }

    // Mettre à jour une donnée
    public ExcelData updateData(Long id, ExcelData data) {
        if (excelDataRepository.existsById(id)) {
            data.setId(id); // Assurez-vous d'avoir un setter pour ID
            return excelDataRepository.save(data);
        }
        return null;
    }

    // Supprimer une donnée
    public void deleteData(Long id) {
        excelDataRepository.deleteById(id);
    }
}
*/