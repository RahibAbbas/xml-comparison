package com.xml.comparison;

import java.io.ByteArrayInputStream;
import java.io.FileOutputStream;
import java.io.StringWriter;
import java.nio.file.Files;
import java.nio.file.Paths;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

@SpringBootApplication
public class ComparisonApplication implements CommandLineRunner {

    public static void main(String[] args) {
        SpringApplication.run(ComparisonApplication.class, args);
    }

    @Override
    public void run(String... args) throws Exception {

        try {
//          String xmlBeforeFilePath = "C:\\Users\\Rahib\\Desktop\\beforeXml.xml"; // Provide the actual file path
//          String xmlAfterFilePath = "C:\\Users\\Rahib\\Desktop\\afterXml.xml";   // Provide the actual file path
            String xmlAfterFilePath = "afterXml.xml";   // Provide in project file path
            String xmlBeforeFilePath = "beforeXml.xml";   // Provide in project file path

            // Read XML content from files
            String xmlBefore = new String(Files.readAllBytes(Paths.get(xmlBeforeFilePath)));
            String xmlAfter = new String(Files.readAllBytes(Paths.get(xmlAfterFilePath)));

            // Parse XML strings
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            DocumentBuilder builder = factory.newDocumentBuilder();
            Document docBefore = builder.parse(new ByteArrayInputStream(xmlBefore.getBytes()));
            Document docAfter = builder.parse(new ByteArrayInputStream(xmlAfter.getBytes()));

            // Get ApplicationID
            String applicationId = getApplicationID(docBefore);

            // Get PolicyMessage elements
            NodeList policyMessagesBefore = docBefore.getElementsByTagName("com.auto.ccs.caf.bom.PolicyMessage");
            NodeList policyMessagesAfter = docAfter.getElementsByTagName("com.auto.ccs.caf.bom.PolicyMessage");

            // Create Excel workbook and sheet
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("XML Comparison");

            // Create headers in Excel
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("ApplicationID");
            headerRow.createCell(1).setCellValue("Before");
            headerRow.createCell(2).setCellValue("Action");
            headerRow.createCell(3).setCellValue("After");
            headerRow.createCell(4).setCellValue("Action");

            // Compare and write changes to Excel
            int rowNumber = 1;
            for (int i = 0; i < Math.max(policyMessagesBefore.getLength(), policyMessagesAfter.getLength()); i++) {
                Element messageBefore = (Element) (i < policyMessagesBefore.getLength() ? policyMessagesBefore.item(i) : null);
                Element messageAfter = (Element) (i < policyMessagesAfter.getLength() ? policyMessagesAfter.item(i) : null);

                String beforeXml = (messageBefore != null) ? elementToString(messageBefore) : "";
                String afterXml = (messageAfter != null) ? elementToString(messageAfter) : "";

                if (messageBefore == null) {
                    // Added case
                    Row dataRow = sheet.createRow(rowNumber++);
                    dataRow.createCell(0).setCellValue(applicationId);
                    dataRow.createCell(1).setCellValue("");
                    dataRow.createCell(2).setCellValue("");
                    dataRow.createCell(3).setCellValue(afterXml);
                    dataRow.createCell(4).setCellValue("Added to AfterXml");
                } else if (messageAfter == null) {
                    // Deleted case
                    Row dataRow = sheet.createRow(rowNumber++);
                    dataRow.createCell(0).setCellValue(applicationId);
                    dataRow.createCell(1).setCellValue(beforeXml);
                    dataRow.createCell(2).setCellValue("Deleted from AfterXml");
                    dataRow.createCell(3).setCellValue("");
                    dataRow.createCell(4).setCellValue("");
                } else if (!beforeXml.equals(afterXml)) {
                    // Updated case
                    Row dataRow = sheet.createRow(rowNumber++);
                    dataRow.createCell(0).setCellValue(applicationId);
                    dataRow.createCell(1).setCellValue(beforeXml);
                    dataRow.createCell(2).setCellValue("Changed in BeforeXml");
                    dataRow.createCell(3).setCellValue(afterXml);
                    dataRow.createCell(4).setCellValue("Added to AfterXml");
                } else {
                    // Unchanged case
                    Row dataRow = sheet.createRow(rowNumber++);
                    dataRow.createCell(0).setCellValue(applicationId);
                    dataRow.createCell(1).setCellValue(beforeXml);
                    dataRow.createCell(2).setCellValue("");
                    dataRow.createCell(3).setCellValue(afterXml);
                    dataRow.createCell(4).setCellValue("");
                }
            }

            // Save Excel file
            try (FileOutputStream fileOut = new FileOutputStream("XMLComparison.xlsx")) {
                workbook.write(fileOut);
            }

            // Close workbook
            workbook.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // Convert an XML element to a string
    private static String elementToString(Element element) {
        try {
            StringWriter sw = new StringWriter();
            TransformerFactory tf = TransformerFactory.newInstance();
            Transformer transformer = tf.newTransformer();
            transformer.setOutputProperty(OutputKeys.OMIT_XML_DECLARATION, "yes");
            transformer.transform(new DOMSource(element), new StreamResult(sw));
            return sw.toString();
        } catch (Exception e) {
            e.printStackTrace();
            return "";
        }
    }

    // Get ApplicationID from <ApplicationID> tag
    private static String getApplicationID(Document document) {
        NodeList nodeList = document.getElementsByTagName("ApplicationID");
        if (nodeList.getLength() > 0) {
            return nodeList.item(0).getTextContent();
        }
        return "";
    }
}

