package org.example.features.builder;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.example.features.builder.annotations.OpenXLSXColumn;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.stream.IntStream;

public class DataModelStage {

    private final OpenXLSX openXLSX;

    public DataModelStage(OpenXLSX openXLSX) {
        this.openXLSX = openXLSX;
    }

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

        return new SheetsStage(openXLSX);
    }

    private <T> void addRowContent(T dataItem) {

        Sheet sheet = openXLSX.getWorkbook().getSheetAt(openXLSX.getCurrentSheetIndex());
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
        if (value instanceof Integer) {
            cell.setCellValue((Integer) value);
        } else if (value instanceof Float) {
            cell.setCellValue((Float) value);
        } else if (value instanceof Double) {
            cell.setCellValue((Double) value);
        } else if (value instanceof Long) {
            cell.setCellValue((Long) value);
        } else if (value instanceof String) {
            cell.setCellValue((String) value);
        } else if (value instanceof Boolean) {
            cell.setCellValue(Boolean.TRUE.equals(value) ? 1 : 0);
        } else if (value instanceof Date) {
            cell.setCellValue(new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format((Date) value));
        } else if (value instanceof BigDecimal) {
            cell.setCellValue(((BigDecimal) value).doubleValue());
        } else if (value instanceof Enum) {
            cell.setCellValue(((Enum<?>) value).name());
        } else {
            cell.setCellValue(value.toString());
        }
    }

    private void setHeader(Field[] declaredFields) {
        Row headerRow = openXLSX.getWorkbook().getSheetAt(openXLSX.getCurrentSheetIndex()).createRow(0);

        IntStream.range(0, declaredFields.length).forEach(i -> {
            Cell cell = headerRow.createCell(i);
            OpenXLSXColumn column = declaredFields[i].getAnnotation(OpenXLSXColumn.class);
            String headerName = column != null ? column.name() : declaredFields[i].getName();
            cell.setCellValue(headerName);
        });
    }

}
