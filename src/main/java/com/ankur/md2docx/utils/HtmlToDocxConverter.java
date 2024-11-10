package com.ankur.md2docx.utils;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;

import java.io.File;

public class HtmlToDocxConverter {

    private static ObjectFactory factory = new ObjectFactory(); // Factory for creating DOCX elements

    // Create a method to convert HTML content into DOCX paragraph
    private void createHtmlParagraph(String htmlContent, MainDocumentPart mainDocPart) {
        // Example: This method simply converts a simple HTML content (like text inside a <p> tag) to a DOCX paragraph.

        // Parse the HTML content (we'll assume it contains simple text, not actual HTML tags for now)
        P paragraph = factory.createP(); // Create a paragraph
        R run = factory.createR(); // Create a run (this is like a span in HTML)
        Text text = factory.createText(); // Create the text element

        // Set the value of the text element (you could add parsing logic for HTML here)
        text.setValue(htmlContent); // Set the text from the HTML content (this would need parsing if it's more complex)

        // Add the text to the run
        run.getContent().add(text);

        // Add the run to the paragraph
        paragraph.getContent().add(run);

        // Add the paragraph to the main document part
        mainDocPart.getContent().add(paragraph);
    }

    // Method to create a DOCX file
    public void createDocxFile(String htmlContent, String filePath) throws Docx4JException {
        // Create a WordprocessingML package (DOCX file)
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();

        // Get the main document part
        MainDocumentPart mainDocPart = wordMLPackage.getMainDocumentPart();

        // Convert HTML content to DOCX paragraph
        createHtmlParagraph(htmlContent, mainDocPart);

        // Save the DOCX file
        File docxFile = new File(filePath);
        wordMLPackage.save(docxFile);
    }

    public static void main(String[] args) throws Docx4JException {
        // Example HTML content to convert to DOCX (simplified, no real HTML parsing)
        String htmlContent = """
                <p>
                    <h1>Welcome Ankur Gupta</h1>
                    <h2>Good Evening</h2>
                </p>
                """;

        // Specify the output DOCX file path
        String filePath = "output1.docx";

        // Create the DOCX file from HTML content
        HtmlToDocxConverter converter = new HtmlToDocxConverter();
        converter.createDocxFile(htmlContent, filePath);

        System.out.println("DOCX file created: " + filePath);
    }
}

