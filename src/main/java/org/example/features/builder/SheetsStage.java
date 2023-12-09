package org.example.features.builder;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

 class SheetsStage extends ExcelBuilder {


    public static DataModelStage createSheet(String sheetName) {
        Sheet sheet = workbook.createSheet(sheetName);
        currentSheetIndex = workbook.getSheetIndex(sheet);
        return new DataModelStage();
    }

    public SheetsStage deleteSheet(String sheetName) {
        workbook.removeSheetAt(workbook.getSheetIndex(workbook.getSheet(sheetName)));
        return this;
    }

    public SheetsStage deleteSheet(int index) {
        workbook.removeSheetAt(index);
        return this;
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    public File generateOutputFileFromSheet() {
        // Utiliza un nombre de archivo basado en un prefijo y un sufijo
        String tempFileName = fileName + "_" + System.currentTimeMillis();
        File file;

        try {
            file = File.createTempFile(tempFileName, ".xlsx");

            // Uso de try-with-resources para asegurarse de que outputStream se cierra correctamente
            try (FileOutputStream outputStream = new FileOutputStream(file)) {
                workbook.write(outputStream);
            }
        } catch (IOException e) {
            throw new ExcelBuilderException(ErrorMessages.OUTPUT_FILE_ERROR + e.getMessage());
        } finally {
            // Cierra el workbook si está abierto
            if (workbook != null) {
                try {
                    workbook.close();
                } catch (IOException e) {
                    // Considera registrar este error o lanzar una excepción, dependiendo de tu estrategia de manejo de errores
                }
            }
        }

        return file;
    }



}