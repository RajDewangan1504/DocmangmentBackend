package com.example.managementSystem.model;

import org.springframework.data.annotation.Id;
import org.springframework.data.mongodb.core.mapping.Document;
import org.springframework.format.annotation.DateTimeFormat;

import lombok.Data;
import lombok.NoArgsConstructor;

@Document(collection = "records")
@Data

@NoArgsConstructor
public class MunicipalRecord {

    @Id
    private String id;

    private String wardNo;

    private String financialYear;

    private String workName;

    private String sanctionedNumber;

  
    private String sanctionDate;

    private String sanctionAmount;

    private String allotedAmount;

    private String tender1Number;

    private String tender1Date;

    private String tender2Number;

    private String tender2Date;

    private String tender3Number;

    private String tender3Date;

    private String tenderRate;

    private String postTenderApprovedAmount;

    private String contractorName;

    private String contractorMobile;

    private String workOrderNumber;

    private String workOrderDate;

    private String agreementNumber;

    private String agreementDate;

    private String actualStartDate;

    private String expectedCompletionDate;

    private String actualCompletionDate;

    private String firstPayment;

    private String secondPayment;

    private String finalPayment;

    private String totalPayment;

    private String remainingAmount;

    private String physicalStatus;

    private String estimatedPhysicalAchievement;

    private String actualPhysicalAchievement;

    private String remarks;

    private String madName;

    private String vidhansabhaName;

    // Add getters & setters (or use Lombok to reduce code)
    

}
