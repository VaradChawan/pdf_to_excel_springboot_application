package com.pdfToExcel.controller;

import java.io.InputStream;
import java.util.List;
import java.util.Map;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.pdfToExcel.services.ConvertPdfToExcelService;

@RestController
public class DocumentConversionController {
    
    @Autowired
    private ConvertPdfToExcelService convertPdfToExcelService;

    @PostMapping("/hello")
    public  String hello(){
        return "Hello World";
    }

    @PostMapping("/pdf-to-excel")
    public ResponseEntity<Object> pdfToExcelConvert(@RequestParam("file") MultipartFile file) throws Exception {
        
        System.out.println(file.getOriginalFilename());
        System.out.println(file.getName());
        System.out.println(file.getSize());
        System.out.println(file.isEmpty());        
        
        if(file.isEmpty()  || !file.getOriginalFilename().toLowerCase().endsWith(".pdf")){
            return new ResponseEntity<>("Invalid File Type. Please upload a PDF file", HttpStatus.BAD_REQUEST);
        }
        List<Map<String,String>> transactions=convertPdfToExcelService.convertPdfToExcel(file);
       
        return new ResponseEntity<>(transactions,HttpStatus.OK);
    }


}
