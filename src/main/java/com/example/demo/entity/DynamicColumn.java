//CODE
package com.example.demo.entity;

import com.fasterxml.jackson.annotation.JsonBackReference;
import jakarta.persistence.*;

@Entity
@Table(name = "dynamic_column")
public class DynamicColumn {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    private String columnName;
    private String columnValue;

    @ManyToOne
    @JoinColumn(name = "data_record_id")
    @JsonBackReference
    private DataRecord dataRecord;

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public String getColumnName() {
        return columnName;
    }

    public void setColumnName(String columnName) {
        this.columnName = columnName;
    }

    public String getColumnValue() {
        return columnValue;
    }

    public void setColumnValue(String columnValue) {
        this.columnValue = columnValue;
    }

    public DataRecord getDataRecord() {
        return dataRecord;
    }

    public void setDataRecord(DataRecord dataRecord) {
        this.dataRecord = dataRecord;
    }
}
