package com.example.demo.repository;

import com.example.demo.entity.DataRecord;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface DataRecordRepository extends JpaRepository<DataRecord, Long> {
}
