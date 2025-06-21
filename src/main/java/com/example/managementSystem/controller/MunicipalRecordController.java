package com.example.managementSystem.controller;



import com.example.managementSystem.model.MunicipalRecord;
import com.example.managementSystem.repository.MunicipalRecordRepository;

import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.http.ResponseEntity;
import org.springframework.http.HttpStatus;

@RestController
@RequestMapping("/api/municipal-records")
// @CrossOrigin(origins = "*")
@CrossOrigin(origins = "http://localhost:8081") 
public class MunicipalRecordController {

    @Autowired
    private MunicipalRecordRepository municipalRecordRepository;

    @PostMapping("/create")
    public ResponseEntity<MunicipalRecord> createRecord(@RequestBody MunicipalRecord record) {
        System.out.println("➡️ Incoming data: " + record);  // DEBUG PRINT
        MunicipalRecord saved = municipalRecordRepository.save(record);
        return ResponseEntity.ok(saved);
    }

    @GetMapping("/list")
    public ResponseEntity<List<MunicipalRecord>> listRecords() {
        return ResponseEntity.ok(municipalRecordRepository.findAll());
    }
}
