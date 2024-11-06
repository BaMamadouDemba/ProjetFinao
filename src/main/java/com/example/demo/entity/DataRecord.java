//CODE
package com.example.demo.entity;

import com.fasterxml.jackson.annotation.JsonManagedReference;
import jakarta.persistence.*;
import java.util.ArrayList;
import java.util.List;

@Entity
@Table(name = "TableTest")
public class DataRecord {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    @OneToMany(mappedBy = "dataRecord", cascade = CascadeType.ALL, fetch = FetchType.LAZY)
    @JsonManagedReference
    private List<DynamicColumn> columns = new ArrayList<>();

    // Getters et setters
    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public List<DynamicColumn> getColumns() {
        return columns;
    }

    public void setColumns(List<DynamicColumn> columns) {
        this.columns = columns;
    }
}
