//AJOUT

package com.example.demo.service;

import com.example.demo.entity.DataRecord;
import com.example.demo.entity.DynamicColumn;
import com.example.demo.repository.DynamicDataRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;

@Service
public class ExcelImportService {

    @Autowired
    private DynamicDataRepository dynamicDataRepository;

    public void importData(List<DynamicColumn> columns) {
        DataRecord dataRecord = new DataRecord(); // Créez un nouvel enregistrement

        // Vérifiez si les colonnes ne sont pas nulles ou vides avant de les ajouter
        if (columns != null && !columns.isEmpty()) {
            dataRecord.setColumns(columns); // Assignez les colonnes à l'enregistrement
            dynamicDataRepository.save(dataRecord); // Enregistrez l'enregistrement dans la base de données
        }
    }
}
