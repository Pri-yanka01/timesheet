package com.example.timesheet;

import org.springframework.core.io.FileSystemResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.io.File;

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

    @GetMapping("/download")
    public ResponseEntity<Resource> downloadExcel() {
        try {
            File file = new File("timesheet.xlsx");
            if (!file.exists()) {
                return ResponseEntity.notFound().build();
            }

            Resource resource = new FileSystemResource(file);
            
            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"timesheet.xlsx\"")
                    .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                    .body(resource);
        } catch (Exception e) {
            e.printStackTrace();
            return ResponseEntity.status(500).build();
        }
    }
}
