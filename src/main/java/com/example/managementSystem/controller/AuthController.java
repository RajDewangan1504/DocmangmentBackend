// AuthController.java
package com.example.managementSystem.controller;

import com.example.managementSystem.dto.LoginRequest;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/auth")
@CrossOrigin(origins = "http://localhost:8081")
public class AuthController {

    @PostMapping("/login")
    public ResponseEntity<?> login(@RequestBody LoginRequest request) {
        // Replace this with real user fetching logic (e.g., from DB)
        if ("admin".equals(request.getUsername()) && "admin".equals(request.getPassword())) {
            return ResponseEntity.ok().body(
                    new UserResponse("प्रशासक", "Administrator", "नगर पालिक निगम"));
        }
        return ResponseEntity.status(401).body("Invalid username or password");
    }

    static class UserResponse {
        public String name;
        public String role;
        public String department;

        public UserResponse(String name, String role, String department) {
            this.name = name;
            this.role = role;
            this.department = department;
        }
    }

    @PostMapping("/logout")
    public ResponseEntity<?> logout() {
        // In a real setup, you'd invalidate session or token here.
        return ResponseEntity.ok("Logged out successfully");
    }
}
