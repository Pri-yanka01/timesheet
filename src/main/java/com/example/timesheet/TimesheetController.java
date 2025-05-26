package com.example.timesheet;

import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
public class TimesheetController {

    @PostMapping("/submit")
    public ResponseEntity<String> handleSubmit(@RequestBody TimesheetRequest request) {
        try {
            ExcelWriter.write(request);
            return ResponseEntity.ok("Saved to timesheet.xlsx");
        } catch (Exception e) {
            e.printStackTrace();
            return ResponseEntity.status(500).body("Error writing to Excel: " + e.getMessage());
        }
    }
}
