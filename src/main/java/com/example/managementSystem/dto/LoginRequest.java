// LoginRequest.java
package com.example.managementSystem.dto;

import lombok.Data;
import lombok.NoArgsConstructor;

@Data

@NoArgsConstructor
public class LoginRequest {
    private String username;
    private String password;

    // Getters and Setters
}
