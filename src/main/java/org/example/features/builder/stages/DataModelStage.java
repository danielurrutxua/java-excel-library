package org.example.features.builder.stages;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.example.features.builder.ExcelBuilder;
import org.example.features.builder.stages.exceptions.ErrorMessages;
import org.example.features.builder.stages.exceptions.ExcelBuilderException;

import java.lang.reflect.Field;
import java.lang.reflect.Member;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

public class DataModelStage extends ExcelBuilder {

    public <T> SheetsStage setData(List<T> dataList) {
        if (dataList == null || dataList.isEmpty()) {
            throw new ExcelBuilderException(ErrorMessages.LIST_NULL_OR_EMPTY);
        }

        Class<?> dataModelClass = dataList.get(0).getClass();

        if (!Object.class.isAssignableFrom(dataModelClass)) {
            throw new ExcelBuilderException(ErrorMessages.INVALID_DATA_TYPE);
        }

        setHeader(dataModelClass.getDeclaredFields());

        for (T dataItem : dataList) {
            addRowContent(dataItem);
        }

        return new SheetsStage();
    }

    private <T> void addRowContent(T dataItem) {

        Sheet sheet = workbook.getSheetAt(currentSheetIndex);
        Row row = sheet.createRow(sheet.getLastRowNum() + 1);
        Field[] fields = dataItem.getClass().getDeclaredFields();
        try {
            for (int i = 0; i < fields.length; i++) {
                fields[i].setAccessible(true); // Necesario para acceder a campos privados
                Object value = fields[i].get(dataItem);
                fields[i].setAccessible(false);
                Cell cell = row.createCell(i);
                setCellValue(cell, value);
            }
        } catch (IllegalAccessException e) {
            throw new ExcelBuilderException(ErrorMessages.FIELD_ACCESS_ERROR + " " + e.getMessage());
        }
    }

    private void setCellValue(Cell cell, Object value) {
        if (value instanceof Integer integer) {
            cell.setCellValue(integer);
        } else if (value instanceof Float floatVal) {
            cell.setCellValue(floatVal);
        } else if (value instanceof Double doubleVal) {
            cell.setCellValue(doubleVal);
        } else if (value instanceof Long longVal) {
            cell.setCellValue(longVal);
        } else if (value instanceof String stringVal) {
            cell.setCellValue(stringVal);
        } else if (value instanceof Boolean booleanVal) {
            cell.setCellValue(Boolean.TRUE.equals(booleanVal) ? 1 : 0);
        } else if (value instanceof Date dateVal) {
            cell.setCellValue(new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(dateVal));
        } else if (value instanceof BigDecimal bigDecimalVal) {
            cell.setCellValue(bigDecimalVal.doubleValue());
        } else if (value instanceof Enum<?> enumVal) {
            cell.setCellValue(enumVal.name());
        } else {
            cell.setCellValue(value.toString());
        }
    }

    private void setHeader(Field[] declaredFields) {
        List<String> header = Arrays.stream(declaredFields).map(Member::getName).toList();

        // Crear una fila en la hoja para los encabezados
        Row headerRow = workbook.getSheetAt(currentSheetIndex).createRow(0); // La fila de encabezado suele ser la primera

        // Rellenar la fila con los nombres de los campos
        for (int i = 0; i < header.size(); i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(header.get(i));
        }

    }
}
