package com.example.managementSystem.repository;



import com.example.managementSystem.model.MunicipalRecord;
import org.springframework.data.mongodb.repository.MongoRepository;

public interface MunicipalRecordRepository extends MongoRepository<MunicipalRecord, String> {
}
