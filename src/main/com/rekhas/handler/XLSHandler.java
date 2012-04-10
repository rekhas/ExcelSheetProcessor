package main.com.rekhas.handler;

import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.hssf.record.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import main.com.rekhas.processor.SheetProcessor;

public class XLSHandler implements HSSFListener {
    private Map<String, SheetProcessor> sheetProcessors;
    private int sheetIndex;
    private String activeSheetName;
    private List<BoundSheetRecord> sheets = new ArrayList<BoundSheetRecord>();
    private SSTRecord sstRecord;

    public void handle(File file, Map<String, SheetProcessor> sheetProcessors) throws IOException {
        FileInputStream inputStream = null;
        try {
            this.sheetProcessors = sheetProcessors;
            inputStream = new FileInputStream(file);
            HSSFRequest hssfRequest = new HSSFRequest();
            hssfRequest.addListenerForAllRecords(this);

            HSSFEventFactory factory = new HSSFEventFactory();
            POIFSFileSystem poifsFileSystem = new POIFSFileSystem(inputStream);
            factory.processWorkbookEvents(hssfRequest, poifsFileSystem);
        } finally {
            if (inputStream != null) inputStream.close();
        }
    }

    @Override
    public void processRecord(Record record) {
        short recordType = record.getSid();
        switch (recordType) {
            case BOFRecord.sid:
                handleBOFRecord((BOFRecord) record);
                break;

            case BoundSheetRecord.sid:
                handleBoundSheetRecord((BoundSheetRecord) record);
                break;

            case SSTRecord.sid:
                handleSSTRecord((SSTRecord) record);
                break;

            case LabelSSTRecord.sid:
                if (!sheetProcessors.containsKey(activeSheetName)) return;
                handleLabelSSTRecord((LabelSSTRecord) record);
                break;

            case NumberRecord.sid:
                if (!sheetProcessors.containsKey(activeSheetName)) return;
                handleNumberRecord((NumberRecord) record);
                break;
        }

    }

    private void handleNumberRecord(NumberRecord record) {
        processValue(record.getRow(), record.getColumn(), String.valueOf(record.getValue()));
    }

    private void handleLabelSSTRecord(LabelSSTRecord record) {
        processValue(record.getRow(), record.getColumn(), sstRecord.getString(record.getSSTIndex()).toString());
    }

    private void handleSSTRecord(SSTRecord record) {
        sstRecord = record;
    }

    private void handleBoundSheetRecord(BoundSheetRecord record) {
        sheets.add(record);
    }

    private void handleBOFRecord(BOFRecord record) {
        if (record.getType() == BOFRecord.TYPE_WORKSHEET) {
            BoundSheetRecord[] sheetRecords = BoundSheetRecord.orderByBofPosition(sheets);
            activeSheetName = sheetRecords[sheetIndex++].getSheetname();
        }
    }

    private void processValue(int rowIndex, short colIndex, String cellValue) {
        SheetProcessor processor = sheetProcessors.get(activeSheetName);
        processor.process(rowIndex, colIndex, cellValue);
    }
}
