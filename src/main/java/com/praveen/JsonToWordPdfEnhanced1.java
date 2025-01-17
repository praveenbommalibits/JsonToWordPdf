package com.praveen;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.core.type.TypeReference;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.util.Map;

import com.itextpdf.text.Document;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.tool.xml.XMLWorkerHelper;

/**
 * https://mvnrepository.com/artifact/org.apache.poi/poi-ooxml/5.3.0
 * https://mvnrepository.com/artifact/org.apache.poi/poi/5.3.0
 * https://mvnrepository.com/artifact/org.apache.poi/poi-ooxml-schemas/4.1.2
 * https://mvnrepository.com/artifact/org.apache.poi/poi-scratchpad/5.3.0
 * https://mvnrepository.com/artifact/com.itextpdf.tool/xmlworker/5.5.6
 */

public class JsonToWordPdfEnhanced1 {

    public static void main(String[] args) {
        try {
            // Input JSON string
            String jsonString = "{\"name\":\"John Doe\", \"age\":\"30\", \"position\":\"Software Engineer\", \"department\":\"IT\", \"additional_info\":\"Key performer in Q4\"}";

            // Parse JSON into a map
            ObjectMapper objectMapper = new ObjectMapper();
            Map<String, String> data = objectMapper.readValue(jsonString, new TypeReference<Map<String, String>>() {});

            // Input Word template file
            String templatePath = "/Users/bommali/Downloads/JsonToWordPdf/JsonToWordPdf/complex_template.docx";
            String outputPdfPath = "output3.pdf";

            // Replace placeholders and generate PDF directly
            replacePlaceholdersAndGeneratePdf(templatePath, outputPdfPath, data);

            System.out.println("PDF file generated successfully.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void replacePlaceholdersAndGeneratePdf(String templatePath, String pdfPath, Map<String, String> data) throws Exception {
        try (FileInputStream fis = new FileInputStream(templatePath);
             XWPFDocument document = new XWPFDocument(fis);
             FileOutputStream fos = new FileOutputStream(pdfPath)) {

            // Replace placeholders in the document
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

