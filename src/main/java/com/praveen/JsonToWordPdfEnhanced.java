package com.praveen;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.core.type.TypeReference;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.util.Map;

import com.itextpdf.text.Document;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.tool.xml.XMLWorkerHelper;

public class JsonToWordPdfEnhanced {

    public static void main(String[] args) {
        try {
            // Input JSON string
            String jsonString = "{\"name\":\"John Doe\", \"age\":\"30\", \"position\":\"Software Engineer\", \"department\":\"IT\", \"additional_info\":\"Key performer in Q4\"}";

            // Parse JSON into a map
            ObjectMapper objectMapper = new ObjectMapper();
            Map<String, String> data = objectMapper.readValue(jsonString, new TypeReference<Map<String, String>>() {});

            // Input Word template file
            String templatePath = "/Users/bommali/Downloads/JsonToWordPdf/JsonToWordPdf/complex_template.docx";
            String outputWordPath = "output1.docx";
            String outputPdfPath = "output1.pdf";

            // Replace placeholders in Word document
            replacePlaceholdersInWord(templatePath, outputWordPath, data);

            // Convert Word to PDF
            convertWordToPdf(outputWordPath, outputPdfPath);

            System.out.println("Word and PDF files generated successfully.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void replacePlaceholdersInWord(String templatePath, String outputPath, Map<String, String> data) throws IOException {
        try (FileInputStream fis = new FileInputStream(templatePath);
             XWPFDocument document = new XWPFDocument(fis)) {

            for (XWPFParagraph paragraph : document.getParagraphs()) {
                for (XWPFRun run : paragraph.getRuns()) {
                    String text = run.getText(0);
                    if (text != null) {
                        for (Map.Entry<String, String> entry : data.entrySet()) {
                            String placeholder = "{{" + entry.getKey() + "}}";
                            if (text.contains(placeholder)) {
                                text = text.replace(placeholder, entry.getValue());
                                run.setText(text, 0);
                            }
                        }
                    }
                }
            }

            for (XWPFTable table : document.getTables()) {
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph paragraph : cell.getParagraphs()) {
                            for (XWPFRun run : paragraph.getRuns()) {
                                String text = run.getText(0);
                                if (text != null) {
                                    for (Map.Entry<String, String> entry : data.entrySet()) {
                                        String placeholder = "{{" + entry.getKey() + "}}";
                                        if (text.contains(placeholder)) {
                                            text = text.replace(placeholder, entry.getValue());
                                            run.setText(text, 0);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                document.write(fos);
            }
        }
    }

    private static void convertWordToPdf(String wordPath, String pdfPath) throws Exception {
        try (FileInputStream fis = new FileInputStream(wordPath);
             XWPFDocument document = new XWPFDocument(fis);
             FileOutputStream fos = new FileOutputStream(pdfPath)) {

            // Write Word content to HTML
            StringWriter stringWriter = new StringWriter();
            try (PrintWriter writer = new PrintWriter(stringWriter)) {
                for (XWPFParagraph paragraph : document.getParagraphs()) {
                    writer.println("<p>" + paragraph.getText() + "</p>");
                }

                for (XWPFTable table : document.getTables()) {
                    writer.println("<table border='1'>");
                    for (XWPFTableRow row : table.getRows()) {
                        writer.println("<tr>");
                        for (XWPFTableCell cell : row.getTableCells()) {
                            writer.println("<td>" + cell.getText() + "</td>");
                        }
                        writer.println("</tr>");
                    }
                    writer.println("</table>");
                }
            }

            // Convert HTML to PDF
            Document pdfDocument = new Document();
            PdfWriter pdfWriter = PdfWriter.getInstance(pdfDocument, fos);
            pdfDocument.open();
            XMLWorkerHelper.getInstance().parseXHtml(pdfWriter, pdfDocument, new StringReader(stringWriter.toString()));
            pdfDocument.close();
        }
    }
}

