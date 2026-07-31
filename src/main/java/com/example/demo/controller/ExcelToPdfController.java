package com.example.demo.controller;

import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.example.demo.service.IExcelToPdfService;

import jakarta.servlet.http.HttpServletRequest;
import lombok.RequiredArgsConstructor;

@RestController
@RequestMapping("excel")
@RequiredArgsConstructor
public class ExcelToPdfController {

	private final IExcelToPdfService exceltopdfserv;
	
	@PostMapping(value = "/convert", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
	public ResponseEntity<byte[]> convertExcelToPdf(HttpServletRequest request,
	        @RequestParam("file") MultipartFile file) {

		System.err.println("Content type is "+request.getContentType());
	    try {

	        if (file.isEmpty()) {
	            return ResponseEntity.badRequest().build();
	        }

	        byte[] pdfBytes = exceltopdfserv.convertExcelToPdf(file.getInputStream());

	        HttpHeaders headers = new HttpHeaders();
	        headers.setContentType(MediaType.APPLICATION_PDF);
	        headers.setContentDispositionFormData("attachment", "converted_report.pdf");

	        return new ResponseEntity<>(pdfBytes, headers, HttpStatus.OK);

	    } catch (Exception e) {
	        e.printStackTrace();
	        return ResponseEntity.internalServerError().build();
	    }
	}
}
