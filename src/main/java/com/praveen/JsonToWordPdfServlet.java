package com.praveen;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.core.type.TypeReference;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.util.Map;

import com.itextpdf.text.Document;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.tool.xml.XMLWorkerHelper;

import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

@WebServlet("/GeneratePdf")
public class JsonToWordPdfServlet extends HttpServlet {

    @Override
    protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
        response.setContentType("application/pdf");
        try {
            // Parse JSON data from request
            String jsonString = request.getParameter("data");
            ObjectMapper objectMapper = new ObjectMapper();
            Map<String, String> data = objectMapper.readValue(jsonString, new TypeReference<Map<String, String>>() {});

            // Read the template file from the request
            InputStream templateInputStream = request.getPart("template").getInputStream();

            // Generate PDF
            ByteArrayOutputStream pdfOutputStream = new ByteArrayOutputStream();
            replacePlaceholdersAndGeneratePdf(templateInputStream, pdfOutputStream, data);

            // Write PDF to response
            response.setHeader("Content-Disposition", "attachment; filename=output.pdf");
            response.getOutputStream().write(pdfOutputStream.toByteArray());
        } catch (Exception e) {
            response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
            response.getWriter().write("Error generating PDF: " + e.getMessage());
        }
    }

    private void replacePlaceholdersAndGeneratePdf(InputStream templateInputStream, OutputStream pdfOutputStream, Map<String, String> data) throws Exception {
        try (XWPFDocument document = new XWPFDocument(templateInputStream)) {

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
            PdfWriter pdfWriter = PdfWriter.getInstance(pdfDocument, pdfOutputStream);
            pdfDocument.open();
            XMLWorkerHelper.getInstance().parseXHtml(pdfWriter, pdfDocument, new StringReader(stringWriter.toString()));
            pdfDocument.close();
        }
    }
}
