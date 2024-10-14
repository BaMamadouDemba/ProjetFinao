package com.example.demo.entity;


/*import jakarta.persistence.*;

import java.sql.Timestamp;

@Entity
@Table(name = "excel_data")
public class ExcelData {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    @Column(name = "file_name")
    private String fileName;

    @Column(name = "sheet_name")
    private String sheetName;

    @Column(name = "column_name")
    private String columnName;

    @Column(name = "value")
    private String value;

    @Column(name = "created_at")
    private Timestamp createdAt;

    // Getters et Setters
    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public String getFileName() {
        return fileName;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public String getColumnName() {
        return columnName;
    }

    public void setColumnName(String columnName) {
        this.columnName = columnName;
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    public Timestamp getCreatedAt() {
        return createdAt;
    }

    public void setCreatedAt(Timestamp createdAt) {
        this.createdAt = createdAt;
    }
}

*/

import jakarta.persistence.*;
import java.util.HashMap;
import java.util.Map;

@Entity
@Table(name = "data_records")  // Table dans la base de données
public class DataRecord {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    @ElementCollection
    @MapKeyColumn(name="column_name")  // Le nom de la colonne Excel
    @Column(name="column_value")       // La valeur de la cellule Excel
    @CollectionTable(name="data_record_values", joinColumns=@JoinColumn(name="data_record_id"))
    private Map<String, String> data = new HashMap<>();

    public DataRecord() {
    }

    public Map<String, String> getData() {
        return data;
    }

    public void setData(Map<String, String> data) {
        this.data = data;
    }
}

