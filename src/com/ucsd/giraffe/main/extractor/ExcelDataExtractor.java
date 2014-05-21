package com.ucsd.giraffe.main.extractor;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelDataExtractor {
    private File excelFile;
    private String[] headers = null;
    String title;

    public ExcelDataExtractor(File excelFile) {
        this.excelFile = excelFile;
    }

    public List<ExtractedRow> extract() {
        try {
            InputStream is = new FileInputStream(excelFile);
            Workbook wb = WorkbookFactory.create(is);
            Sheet s = wb.getSheetAt(1);
            List<ExtractedRow> ret = new ArrayList<>(s.getLastRowNum());
            for (int i = 0; i < s.getLastRowNum(); i++) {
                Row r = s.getRow(i);
                if (r == null) {
                    continue;
                } else if (r.getLastCellNum() == 1) { // skipping title/blank
                                                      // cells/rows
                    title = r.getCell(0).getStringCellValue();
                    continue;
                }
                if (headers == null) {
                    // generate labels
                    headers = new String[r.getLastCellNum()];
                    for (int j = 0; j < r.getLastCellNum(); j++) {
                        headers[j] = r.getCell(j).getStringCellValue();
                    }
                    i++; // skip instruction thing
                } else {
                    ret.add(new ExtractedRow());
                    if (!validateRow(r, headers.length)) {
                        continue;
                    }
                    // populate header map
                    for (int j = 0; j < r.getLastCellNum(); j++) {
                        Cell c = r.getCell(j);
                        String value = null;
                        if (c.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                            value = String.valueOf(c.getNumericCellValue());
                        } else {
                            value = c.getStringCellValue();
                        }
                        if (!value.isEmpty()) {
                            ret.get(ret.size() - 1).addField(headers[j], value);
                        }
                    }
                }
            }
            return ret;
        } catch (Exception e) { // should never happen
            e.printStackTrace();
            return null;
        }
    }

    private boolean validateRow(Row r, int length) {
        for (int i = 0; i < length; i++) {
            if (r.getCell(i) == null) {
                r.createCell(i).setCellValue("");
            }
        }
        return true;
    }

    public String[] getHeaders() {
        return headers;
    }

    public String getTitle() {
        return title;
    }
}
