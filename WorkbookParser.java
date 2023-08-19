package com.pointcarbon.parsers.FLOW;

import com.pointcarbon.esb.commons.beans.EsbMessage;
import com.pointcarbon.parserframework.service.IParser;
import com.pointcarbon.parserframework.service.OutputServiceFactory;
import com.pointcarbon.parserframework.service.output.STFDataOutputter;
import org.apache.poi.ss.usermodel.*;
import org.joda.time.DateTime;
import org.joda.time.format.DateTimeFormat;
import org.joda.time.format.DateTimeFormatter;

import java.io.IOException;
import java.io.InputStream;
import java.util.LinkedHashMap;
import java.util.Map;

import static com.pointcarbon.parserframework.service.output.STFDataOutputter.COLUMN.*;

public class WorkbookParser implements IParser {
    private static final DateTimeFormatter DTF = DateTimeFormat.forPattern("dd-MMM-yyyy");
    private static final String[] headerCells = {"reference", "update", "vessel", "imo", "status", "depcountry", "depport", "depcode",
            "descountry", "desport", "descode", "opecountry", "opetype", "eta", "opestart", "opeend",
            "comgroup", "comdesc", "qty", "origin1", "origin2", "origin3", "export reference"
    };
    private static final String DTF1 = "MM/dd/yyyy/HH:mmZ";
    private static final String DTF2 = "MM/dd/yyyy/HH:mm";
    private static final String EXPORT_TYPE = "Export";
    private static final String IMPORT_TYPE = "Import";
    private static final String SAILED_STATUS = "Sailed";
    private static final String LOAD_DISCH_STATUS = "Loading/Discharging";
    private static final String ACTUAL_DATA_TYPE = "Actual";
    private static final String FORECAST_DATA_TYPE = "Forecast";
    private static final String UNKNOWN_OPERATION = "Unknown Operation";
    private static final int ID = 0;
    private static final int SHEET_PAGE = 0;
    private static final int HEADER_ROW = 0;
    private static final int SUBSTRING_START = 0;
    private static final int SUBSTRING_END_FRE_COMMENTS = 1;
    private static final int START_ROW = 1;
    private static final int UPDATE = 1;
    private static final int VESSEL_NAME = 2;
    private static final int VESSEL_IMO = 3;
    private static final int STATUS = 4;
    private static final int DEP_COUNTRY = 5;
    private static final int DEPORT = 6;
    private static final int DESCOUNTRY = 8;
    private static final int DESPORT = 9;
    private static final int TYPE = 12;
    private static final int ETA = 13;
    private static final int LOAD_DATA_FROW = 14;
    private static final int LOAD_DATA_TO = 15;
    private static final int COM_DESC = 17;
    private static final int VOLUME = 18;
    private static final int ORIGIN_1 = 19;
    private static final int ORIGIN_2 = 20;
    private static final int ORIGIN_3 = 21;
    private static final int CELL_LIMIT = 22;
    private static final int EXPORT_REFERENCE = 22;


    @Override
    public void parse(InputStream is, EsbMessage message, OutputServiceFactory factory) throws Exception {
        STFDataOutputter outputter = factory.createSTFDataOutputter();
        Workbook workbook = WorkbookFactory.create(is);
        final Sheet mainSheet = workbook.getSheetAt(SHEET_PAGE);
        parseXLS(mainSheet, outputter);
    }

    private void parseXLS(Sheet sheet, STFDataOutputter outputter) throws IOException {
        for (int rowIdx = START_ROW; rowIdx < sheet.getLastRowNum(); rowIdx++) {
            Row rowData = sheet.getRow(rowIdx);
            verifyHeader(sheet, headerCells);
            if (rowData.getLastCellNum() >= CELL_LIMIT) {
                String id = getValue(rowData, ID);
                String createDate = getFreCreateDate(rowData);
                String type = getValue(rowData, TYPE);
                String loadDateFrom = getLoadDate(rowData, type, LOAD_DATA_FROW);
                String loadDateTo = getLoadDate(rowData, type, LOAD_DATA_TO);
                String loadPortName = getLoadPortName(rowData, type);
                String loadDateFcstActl = getLoadDateFcstActl(rowData, type);
                String dischLocation = getDischLocation(rowData, type);
                String arrDateFrom = getFreArrDateFrom(rowData, type, LOAD_DATA_FROW);
                String arrDateTo = getFreArrDateFrom(rowData, type, LOAD_DATA_TO);
                String dischPortName = getDischPortName(rowData, type);
                String arrDateFcstActl = getFreArrDateFcstActl(rowData, type);
                String loadLocation = getLoadLocation(rowData, type);
                String status = getValue(rowData, STATUS);
                String vesName = getValue(rowData, VESSEL_NAME);
                String vesImo = getValue(rowData, VESSEL_IMO);
                String cmdName = getValue(rowData, COM_DESC);
                String volume = getValue(rowData, VOLUME);
                String comments = concatFreComments(rowData);
                String commentsPrivate = getFreCommentsPrivate(type);
                String loadGunName = getLoadGunName(rowData, type);
                String dischGunName = getDishGunName(rowData, type);

                try {
                    outputter.commit(id, FLOW_ID);
                    outputter.commit(createDate, FLOW_CREATE_DATE);
                    outputter.commit(loadDateFrom, FRE_LOAD_DATE_FROM);
                    outputter.commit(loadDateTo, FRE_LOAD_DATE_TO);
                    outputter.commit(loadPortName, FRE_LOAD_PORT_NAME);
                    outputter.commit(loadDateFcstActl, FRE_LOAD_DATE_FCST_ACTL);
                    outputter.commit(dischLocation, FLOW_DISCH_LOCATION);
                    outputter.commit(arrDateFrom, FRE_ARR_DATE_FROM);
                    outputter.commit(arrDateTo, FRE_ARR_DATE_TO);
                    outputter.commit(dischPortName, FRE_DISCH_PORT_NAME);
                    outputter.commit(arrDateFcstActl, FRE_ARR_DATE_FCST_ACTL);
                    outputter.commit(loadLocation, FLOW_LOAD_LOCATION);
                    outputter.commit(status, FRE_STATUS);
                    outputter.commit(vesName, FRE_VES_NAME);
                    outputter.commit(vesImo, FRE_VES_IMO);
                    outputter.commit(cmdName, FRE_CMD_NAME);
                    outputter.commit(volume, FRE_VOLUME);
                    outputter.commit(comments, FRE_COMMENTS);
                    outputter.commit(commentsPrivate, FRE_COMMENTS_PRIVATE);
                    outputter.commit(loadGunName, FRE_LOAD_GUN_NAME);
                    outputter.commit(dischGunName, FRE_DISCH_GUN_NAME);
                } catch (Exception e) {
                    outputter.rollabck();
                }
                outputter.push();
            }
        }
    }

    public void verifyHeader(Sheet sheet, String[] header) {
        Row headerRow = sheet.getRow(HEADER_ROW);
        for (int i = HEADER_ROW; i < headerRow.getLastCellNum(); i++) {
            Cell headerCell = headerRow.getCell(i);
            if (headerCell != null) {
                String headerValue = headerCell.toString();
                if (!header[i].equalsIgnoreCase(headerValue)) {
                    throw new RuntimeException("Header has been changed! Expected: " + header[i] + "; Actual: " + headerValue);
                }
            }
        }
    }

    private String concatFreComments(Row row) {
        LinkedHashMap<String, String> commentsMap = new LinkedHashMap<>();
        commentsMap.put("ETA:", getParseDateTime(row, ETA));
        commentsMap.put("Operation:", getValue(row, TYPE));
        commentsMap.put("Origin1:", getValue(row, ORIGIN_1));
        commentsMap.put("Origin2:", getValue(row, ORIGIN_2));
        commentsMap.put("Origin3:", getValue(row, ORIGIN_3));
        commentsMap.put("ExportReference:", getValue(row, EXPORT_REFERENCE));
        return concatHelper(commentsMap, true, SUBSTRING_END_FRE_COMMENTS);
    }

    private String concatHelper(LinkedHashMap<String, String> portNames, boolean freComments, int subEndPosition) {
        String concat = "";
        for (Map.Entry<String, String> entry : portNames.entrySet()) {
            if (!entry.getValue().isEmpty()) {
                if (freComments) {
                    concat = concat + entry.getKey() + entry.getValue() + ";";
                } else {
                    concat = concat + entry.getValue() + ", ";
                }
            }
        }
        if (!concat.isEmpty()) {
            return concat.substring(SUBSTRING_START, concat.length() - subEndPosition);
        }
        return "";
    }

    private String getFreCommentsPrivate(String type) {
        if (!type.equalsIgnoreCase(EXPORT_TYPE) && !type.equalsIgnoreCase(IMPORT_TYPE)) {
            return UNKNOWN_OPERATION;
        }
        return "";
    }

    private String getFreArrDateFcstActl(Row row, String type) {
        if (!type.equalsIgnoreCase(EXPORT_TYPE)) {
            return helperLoadDateFcstActl(row);
        }
        return "";
    }

    private String getLoadDateFcstActl(Row row, String type) {
        if (type.equalsIgnoreCase(EXPORT_TYPE)) {
            return helperLoadDateFcstActl(row);
        }
        return "";
    }

    private String helperLoadDateFcstActl(Row row) {
        String status = getValue(row, STATUS);
        if (status.equalsIgnoreCase(SAILED_STATUS) || status.equalsIgnoreCase(LOAD_DISCH_STATUS)) {
            return ACTUAL_DATA_TYPE;
        }
        return FORECAST_DATA_TYPE;
    }

    private String getLoadLocation(Row row, String type) {
        if (!type.equalsIgnoreCase(EXPORT_TYPE)) {
            return getValue(row, DEPORT);
        }
        return "";
    }

    private String getDischPortName(Row row, String type) {
        if (!type.equalsIgnoreCase(EXPORT_TYPE)) {
            return getValue(row, DESPORT);
        }
        return "";
    }

    private String getLoadPortName(Row row, String type) {
        if (type.equalsIgnoreCase(EXPORT_TYPE)) {
            return getValue(row, DEPORT);
        }
        return "";
    }

    private String getDischLocation(Row row, String type) {
        if (type.equalsIgnoreCase(EXPORT_TYPE)) {
            return getValue(row, DESPORT);
        }
        return "";
    }

    private String getFreArrDateFrom(Row row, String type, int position) {
        if (!type.equalsIgnoreCase(EXPORT_TYPE)) {
            return getParseDateTime(row, position);
        }
        return "";
    }

    private String getLoadDate(Row row, String type, int position) {
        if (type.equalsIgnoreCase(EXPORT_TYPE)) {
            return getParseDateTime(row, position);
        }
        return "";
    }

    private String getLoadGunName(Row row, String type) {
        if (type.equalsIgnoreCase(EXPORT_TYPE)) {
            return getValue(row, DEP_COUNTRY);
        }
        return "";
    }

    private String getDishGunName(Row row, String type) {
        if (!type.equalsIgnoreCase(EXPORT_TYPE)) {
            return getValue(row, DESCOUNTRY);
        }
        return "";
    }

    private String getParseDateTime(Row row, int position) {
        String date = getValue(row, position);
        if (!date.isEmpty()) {
            return DTF.parseDateTime(getValue(row, position)).toString(DTF2);
        }
        return "";
    }

    private String getFreCreateDate(Row row) {
        String createDate = getValue(row, UPDATE);
        if (!createDate.isEmpty()) {
            return DTF.parseDateTime(createDate).toString(DTF1);
        } else {
            return DateTime.now().toString(DTF1);
        }
    }

    private String getValue(Row row, int position) {
        Cell cell = row.getCell(position);
        if (cell != null) {
            return cell.toString();
        }
        return "";
    }

}