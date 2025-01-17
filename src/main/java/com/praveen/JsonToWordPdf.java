package com.praveen;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.core.type.TypeReference;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.*;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import com.itextpdf.text.Document;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.tool.xml.XMLWorkerHelper;

public class JsonToWordPdf {

    public static void main(String[] args) {
        try {
            // Input JSON string
            String jsonString = "{\"name\":\"John Doe\", \"age\":\"30\", \"position\":\"Software Engineer\"}";

            // Parse JSON into a map
            ObjectMapper objectMapper = new ObjectMapper();
            Map<String, String> data = objectMapper.readValue(jsonString, new TypeReference<Map<String, String>>() {});

            // Input Word template file
            String templatePath = "/Users/bommali/Downloads/JsonToWordPdf/JsonToWordPdf/template.docx";
            String outputWordPath = "output.docx";
            String outputPdfPath = "output.pdf";

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
