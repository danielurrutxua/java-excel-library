package org.example.features.builder;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class OpenXLSX {
    private String fileName;
    private final Workbook workbook;
    private int currentSheetIndex;

    public OpenXLSX(){
        workbook = new XSSFWorkbook();
    }

    public  SheetsStage initWorkbook(String name) {
        fileName = name;
        return new SheetsStage(this);
    }

    protected String getFileName() {
        return fileName;
    }

    protected Workbook getWorkbook() {
        return workbook;
    }

    protected int getCurrentSheetIndex() {
        return currentSheetIndex;
    }

    protected void setCurrentSheetIndex(int currentSheetIndex) {
        this.currentSheetIndex = currentSheetIndex;
    }
}
