package org.example.features.builder;

import org.apache.poi.ss.usermodel.Sheet;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class SheetsStage {

    private final OpenXLSX openXLSX;

    public SheetsStage(OpenXLSX openXLSX) {
        this.openXLSX = openXLSX;
    }

    public DataModelStage createSheet(String sheetName) {
        Sheet sheet = openXLSX.getWorkbook().createSheet(sheetName);
        openXLSX.setCurrentSheetIndex(openXLSX.getWorkbook().getSheetIndex(sheet));
        return new DataModelStage(openXLSX);
    }

    public DataModelStage createSheet() {
        Sheet sheet = openXLSX.getWorkbook().createSheet();
        openXLSX.setCurrentSheetIndex(openXLSX.getWorkbook().getSheetIndex(sheet));
        return new DataModelStage(openXLSX);
    }

    public SheetsStage deleteSheet(String sheetName) {
        openXLSX.getWorkbook().removeSheetAt(openXLSX.getWorkbook().getSheetIndex(openXLSX.getWorkbook().getSheet(sheetName)));
        return this;
    }

    public SheetsStage deleteSheet(int index) {
        openXLSX.getWorkbook().removeSheetAt(index);
        return this;
    }

    public File generateOutputFileFromSheet() {
        FileOutputStream outputStream;
        String tempFileName = openXLSX.getFileName() + "_" + System.currentTimeMillis();
        File file;
        try {
            file = File.createTempFile(tempFileName, ".xlsx");
            outputStream = new FileOutputStream(file);
            openXLSX.getWorkbook().write(outputStream);
            openXLSX.getWorkbook().close();
        } catch (IOException e) {
            throw new ExcelBuilderException(ErrorMessages.OUTPUT_FILE_ERROR + e.getMessage());
        }
        return file;
    }


}