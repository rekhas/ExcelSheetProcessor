package main.com.rekhas.handler;

import main.com.rekhas.processor.SheetProcessor;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRst;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class XLSXHandler extends DefaultHandler {
    private static final char ASCII_OF_A = 'A';
    private static final int ALPHABET_BASE = 26;
    private Map<String, SheetProcessor> sheetProcessors;
    private Map<String, String> sheetsAndIds = new HashMap<String, String>();
    private String activeSheetName;
    private boolean isCellValueAString;
    private boolean shouldProcessCell;
    private long colIndex;
    private long rowIndex;
    private String contents;
    private SharedStringsTable sst;

    public void handle(File file, Map<String, SheetProcessor> sheetProcessors) throws IOException, OpenXML4JException, SAXException {
        this.sheetProcessors = sheetProcessors;
        XSSFReader reader = new XSSFReader(OPCPackage.open(new FileInputStream(file)));
        sst = reader.getSharedStringsTable();
        XMLReader parser = XMLReaderFactory.createXMLReader(
                "com.sun.org.apache.xerces.internal.parsers.SAXParser"
        );
        parser.setContentHandler(this);
        InputStream workbookData = reader.getWorkbookData();
        parser.parse(new InputSource(workbookData));
        for (String sheetId : sheetsAndIds.keySet()) {
            InputStream stream = null;
            try {
                stream = reader.getSheet(sheetId);
                activeSheetName = sheetsAndIds.get(sheetId);
                if (sheetProcessors.containsKey(activeSheetName)) {
                    InputSource sheetSource = new InputSource(stream);
                    parser.parse(sheetSource);
                }
            } finally {
                if (stream != null) stream.close();
            }
        }
    }

    @Override
    public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {
        if (qName.equals("c")) handleStartOfCellNode(attributes);
        else if (qName.equals("v")) handleStartOfValueNode();
        else if (qName.equals("sheet")) handleSheetNode(attributes);
    }

    private void handleSheetNode(Attributes attributes) {
        sheetsAndIds.put(attributes.getValue("r:id"), attributes.getValue("name"));
    }

    @Override
    public void characters(char[] ch, int start, int length) throws SAXException {
        contents = new String(ch, start, length);
    }

    @Override
    public void endElement(String uri, String localName, String qName) throws SAXException {
        if (shouldProcessCell && qName.equals("c")) handleEndOfCellNode();
    }

    private void handleEndOfCellNode() {
        if (isCellValueAString) {
            CTRst entryAt = sst.getEntryAt(Integer.parseInt(contents));
            contents = new XSSFRichTextString(entryAt).toString();
        }
        sheetProcessors.get(activeSheetName).process(rowIndex, colIndex, contents);
        shouldProcessCell = false;
    }

    private void handleStartOfValueNode() {
        shouldProcessCell = true;
    }

    private void handleStartOfCellNode(Attributes attributes) {
        String cellPosn = attributes.getValue("r");
        isCellValueAString = "s".equals(attributes.getValue("t"));
        setRowAndColIndex(cellPosn);
    }

    private void setRowAndColIndex(String cellPosn) {
        Matcher matcher = Pattern.compile("([A-Z]*)([0-9]*)").matcher(cellPosn);
        matcher.find();
        colIndex = makeIndexZeroBased(convertExcelColumnIndexToNumericIndex(matcher.group(1)));
        rowIndex = makeIndexZeroBased(Long.parseLong(matcher.group(2)));
    }

    private long makeIndexZeroBased(long nonZeroBasedIndex) {
        return nonZeroBasedIndex - 1;
    }

    private int convertExcelColumnIndexToNumericIndex(String group) {
        int length = group.length() - 1;
        int index = length;
        int baseAscii = ASCII_OF_A - 1;
        int result = 0;
        while (index >= 0) {
            int asciiValue = group.charAt(index);
            result += (Math.pow(ALPHABET_BASE, length - index)) * (asciiValue % baseAscii);
            index--;
        }
        return result;
    }
}
