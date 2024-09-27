package com.pdfToExcel.services;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

@Service
public class ConvertPdfToExcelService {

    public List<Map<String,String>> convertPdfToExcel(MultipartFile file) throws Exception {

        System.out.println(file.getOriginalFilename());
        System.out.println(file.getName());
        System.out.println(file.getSize());
        System.out.println(file.isEmpty());

        InputStream is = file.getInputStream();
        String getTextFromPDF = extractTextFromString(is);
        
        System.err.println(">>>>>>>>>>>  :::::::: "+ getTextFromPDF);
        
        ByteArrayOutputStream excelOutputStream =  convertTextToExcelFile(getTextFromPDF);

        List<Map<String,String>> transcations= sendResponseFromExcelFile(excelOutputStream);
        //String [] lines= getTextFromPDF.split("\n");
        
        
        return transcations;
    }

    private List<Map<String, String>> sendResponseFromExcelFile(ByteArrayOutputStream excelInputStream) throws Exception{
        
        List<Map<String,String>> transactions = new ArrayList<>();
        Workbook workbook = new HSSFWorkbook(new ByteArrayInputStream(excelInputStream.toByteArray()));
        Sheet parsedSheet = workbook.getSheetAt(0);
        
        for(int rowIndex=1; rowIndex<= parsedSheet.getLastRowNum();rowIndex++){
            
            Row row =parsedSheet.getRow(rowIndex);

            if(row !=null){
                Map<String,String> transaction = new HashMap<>();
                transaction.put("Date", getCellValueAsString(row.getCell(0)));
                transaction.put("Description", getCellValueAsString(row.getCell(1)));
                transaction.put("Transaction Id", getCellValueAsString(row.getCell(2)));
                transaction.put("Transaction Amount", getCellValueAsString(row.getCell(3)));
                transaction.put("Balance", getCellValueAsString(row.getCell(4)));
                transactions.add(transaction);
            }
        }
        
        
        // TODO Auto-generated method stub
        //throw new UnsupportedOperationException("Unimplemented method 'sendResponseFromExcelFile'");
        workbook.close();

        return transactions;
    }

    private String getCellValueAsString(Cell cell){
            if(cell==null){
                return "";
            }
            
            switch (cell.getCellType()) {
                case STRING:
                    return cell.getStringCellValue();
                case NUMERIC:
                    if(DateUtil.isCellDateFormatted(cell)){
                        return cell.getDateCellValue().toString();
                    }else{
                        return Double.toString(cell.getNumericCellValue());
                    }
                case BOOLEAN:
                    return Boolean.toString(cell.getBooleanCellValue());
                default:
                    return "";
            }
    }

    private String extractTextFromString(InputStream inputStream) throws Exception {
        PDDocument document = PDDocument.load(inputStream);

        PDFTextStripper pdfTextStripper = new PDFTextStripper();
        pdfTextStripper.setStartPage(1);
        pdfTextStripper.setEndPage(document.getNumberOfPages());
        String extractedText = pdfTextStripper.getText(document);
        // System.err.println("DOcument total pages==> "+ extractedText);

        document.close();
        String tableData= extractTable(extractedText);

        return tableData;
    }

    private String extractTable(String fullText){
        String tableStartMarker="Date";
        String tableEndMarker="Total";


        int startIndex=fullText.indexOf(tableStartMarker);
        int endIndex=fullText.indexOf(tableEndMarker,startIndex);

        if(startIndex != -1 && endIndex != -1){
            return fullText.substring(startIndex, endIndex).trim();
        }

        return "";
    }


    private ByteArrayOutputStream convertTextToExcelFile(String text)throws Exception {
        Workbook workbook= new HSSFWorkbook();
        Sheet sheet = workbook.createSheet("PDF_Data");

        String []lines=text.split("\n");

        int rowNum=0;
        Row headerRow=sheet.createRow(rowNum++);
        headerRow.createCell(0).setCellValue("Date");
        headerRow.createCell(1).setCellValue("Description");
        headerRow.createCell(2).setCellValue("Transaction Id");
        headerRow.createCell(3).setCellValue("Transaction Amount");
        headerRow.createCell(4).setCellValue("Balance");


        for(String line : lines){
            String []transactionDetails= line.split("\\s+");

            if(transactionDetails.length>=5){
                Row row= sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(transactionDetails[0]);
                row.createCell(1).setCellValue(transactionDetails[1]);
                row.createCell(2).setCellValue(transactionDetails[2]);
                row.createCell(3).setCellValue(transactionDetails[3]);
                row.createCell(4).setCellValue(transactionDetails[4]);
                
            }
        }

        ByteArrayOutputStream byteArrayOutputStream=new ByteArrayOutputStream();
        workbook.write(byteArrayOutputStream);
        workbook.close();
        return byteArrayOutputStream;
    }

}
