package org.example.features.builder;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.example.features.builder.stages.SheetsStage;

public class ExcelBuilder {
    protected static String fileName;
    protected static Workbook workbook;
    protected static int currentSheetIndex;

    protected ExcelBuilder() {
    }

    public static SheetsStage init(String name) {
        workbook = new HSSFWorkbook();
        fileName = name;
        return new SheetsStage();
    }
}
