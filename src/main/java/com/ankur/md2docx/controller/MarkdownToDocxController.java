package com.ankur.md2docx.controller;

import org.commonmark.parser.Parser;
import org.commonmark.renderer.html.HtmlRenderer;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;
import org.docx4j.convert.in.word2003xml.*;

import org.docx4j.wml.ObjectFactory;
import org.springframework.web.bind.annotation.*;
import org.springframework.http.ResponseEntity;
import org.springframework.http.MediaType;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;

@RestController
@RequestMapping("/api/convert")
public class MarkdownToDocxController {

    private static final ObjectFactory factory = new ObjectFactory(); // Factory for creating DOCX elements

    @PostMapping("/markdown-to-docx")
    public ResponseEntity<byte[]> convertMarkdownToDocx(@RequestParam("file") MultipartFile file) throws IOException, Docx4JException {
        // Step 1: Parse the uploaded markdown file
        String markdownContent = new String(file.getBytes());

        // Step 2: Parse markdown to HTML using CommonMark
        Parser parser = Parser.builder().build();
        HtmlRenderer renderer = HtmlRenderer.builder().build();
        org.commonmark.node.Node document = parser.parse(markdownContent);
        String htmlContent = renderer.render(document);

        // Step 3: Convert HTML content to DOCX
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
        MainDocumentPart mainDocPart = wordMLPackage.getMainDocumentPart();

        // Convert the HTML content into DOCX format
        mainDocPart.addParagraphOfText("Markdown Converted to DOCX:");
        createHtmlParagraph(htmlContent, mainDocPart);

        // Step 4: Write the DOCX to a byte array
        ByteArrayOutputStream docxStream = new ByteArrayOutputStream();
        wordMLPackage.save(docxStream);

        byte[] docxFile = docxStream.toByteArray();

        // Step 5: Return the DOCX file as a download
        return ResponseEntity.ok()
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .header("Content-Disposition", "attachment; filename=\"Converted.docx\"")
                .body(docxFile);
    }

//    private void createHtmlParagraph(String htmlContent, MainDocumentPart mainDocPart) {
//        // Using docx4j's HTML converter (optional for advanced formatting)
//
//        // Parse the HTML and convert to a WordParagraph
//        // You can customize this method as per your needs.
//
//        // For simplicity, we can add basic HTML-to-DOCX mappings manually or use an existing tool like the HTMLImporter.
//
//        // Example: Add a simple paragraph to the document
//        P paragraph = factory.createP();
//        R run = factory.createR();
//        Text text = factory.createText();
//        text.setValue(htmlContent);
//        run.getContent().add(text);
//        paragraph.getContent().add(run);
//        mainDocPart.getContent().add(paragraph);
//    }

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


//    public void convertHtmlToDocx(String htmlContent, MainDocumentPart mainDocPart) {
//        HTMLImporter htmlImporter = new HTMLImporter(mainDocPart);
//        htmlImporter.convert(htmlContent, mainDocPart);
//    }
//
//    public static File convertHtmlToDocx( Reader htmlReader ) throws IOException, JAXBException, Docx4JException {
//        File outFile = File.createTempFile("edition", "docx");
//        outFile.deleteOnExit();
//
//        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
//        NumberingDefinitionsPart ndp = new NumberingDefinitionsPart();
//        wordMLPackage.getMainDocumentPart().addTargetPart(ndp);
//        ndp.unmarshalDefaultNumbering();
//
//        // Convert the XHTML, and add it into the empty docx we made
//        wordMLPackage.getMainDocumentPart().getContent().addAll(
//                XHTMLImporter.convert(htmlReader, null, wordMLPackage) );
//
//        wordMLPackage.save( outFile );
//
//        return outFile;
//    }
}