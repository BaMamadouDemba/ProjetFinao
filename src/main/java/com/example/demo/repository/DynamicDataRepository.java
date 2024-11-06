//AJOUT
package com.example.demo.repository;
import com.example.demo.entity.DataRecord;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.stereotype.Repository;
import java.util.List;
import java.util.Map;

@Repository
public interface DynamicDataRepository extends JpaRepository<DataRecord, Long> {

    @Query(value = "SELECT * FROM TableTest", nativeQuery = true)
    List<Map<String, Object>> findAllDynamicData();
}
