package model;

import annotation.ReportKey;
import enums.DataType;
import enums.KeyType;
import exception.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.ParameterizedType;
import java.text.SimpleDateFormat;
import java.util.*;

public class Report {

    private static final String COMPLEX_KEY = "complex_";
    private static final String DATABASE_KEY = "REPORT_KEYS";
    private static final String COUNTER_KEY = "key_counter";
    public static final double PIXEL_TO_ROW_HEIGHT = 15.0;

    private final Object reportData;
    private final Map<String, KeyData> keysMap;
    private final Set<String> complexKeys;
    private final Set<Class<?>> classes;

    public Report(Object reportData) throws ReportKeyException {
        this.reportData = reportData;
        this.keysMap = new HashMap<>();
        this.complexKeys = new HashSet<>();
        this.classes = new HashSet<>();
        fillKeysMap(reportData.getClass());
    }

    public void createTemplate(String path) throws IOException {
        XSSFWorkbook wb = new XSSFWorkbook();
        wb.createSheet("Sheet1");
        keysMap.forEach((key, value) -> {
            if (value.getReportKey().keyType() == KeyType.COMPLEX) {
                wb.createSheet(key);
            }
        });
        XSSFSheet sheet = wb.createSheet(DATABASE_KEY);
        fillDataBaseSheet(sheet);

        FileOutputStream output = new FileOutputStream(path);
        wb.write(output);
        output.close();
        wb.close();
    }

    public void createReport(InputStream template, File file) throws IOException, IncorrectTemplateException, ReportKeyException {
        XSSFWorkbook wb = new XSSFWorkbook(template);
        checkCorrectnessFillingTemplate(wb);

        List<XSSFCell> cellList = new ArrayList<>();
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            XSSFSheet sheet = wb.getSheetAt(i);
            if (!sheet.getSheetName().equals(DATABASE_KEY) && !keysMap.containsKey(sheet.getSheetName())) {
                cellList.addAll(getCellListWithKey(sheet, 0, sheet.getLastRowNum()));
            }
        }

        for (XSSFCell cell : cellList) {
            fillCellByKey(cell, reportData, 0);
        }

        deleteKeySheets(wb);

        FileOutputStream fileOutput = new FileOutputStream(file);
        wb.write(fileOutput);
        wb.close();
        fileOutput.close();
    }

    private void fillKeysMap(Class<?> c) throws ReportKeyException {
        classes.add(c);
        for (Field field : c.getDeclaredFields()) {
            if (field.isAnnotationPresent(ReportKey.class)) {
                ReportKey reportKey = field.getAnnotation(ReportKey.class);
                String key = reportKey.name();
                if (key.equals("default")) {
                    String prefix = (reportKey.keyType() == KeyType.COMPLEX) ? COMPLEX_KEY : "key_";
                    key = prefix + c.getSimpleName() + "_" + field.getName();
                }
                if (keysMap.containsKey(key)) {
                    throw new IdenticalReportKeyException("Annotated fields have the identical names of ReportKey");
                }
                String description = reportKey.description();
                if (description.equals("default")) {
                    description = "object: " + c.getSimpleName() + ", " + "field: " + field.getName();
                }
                field.setAccessible(true);
                KeyData keyData = new KeyData(key, description, reportKey, field);
                keysMap.put(key, keyData);

                if (reportKey.keyType() == KeyType.COMPLEX) {
                    Class<?> n = field.getType();
                    if (List.class.isAssignableFrom(n)) { //todo add Collections
                        complexKeys.add(key);
                        n = (Class<?>) ((ParameterizedType) field.getGenericType()).getActualTypeArguments()[0];
                        if (!classes.contains(n)) {
                            fillKeysMap(n);
                        }
                    } else {
                        throw new IncorrectReportKeyException("The annotation ReportKey does not match the field");
                    }
                }
            }
        }
    }

    private void fillDataBaseSheet(XSSFSheet sheet) {
        XSSFRow headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Имя ключа");
        headerRow.createCell(1).setCellValue("Тип ключа");
        headerRow.createCell(2).setCellValue("Тип данных");
        headerRow.createCell(3).setCellValue("Удаляемый ключ");
        headerRow.createCell(4).setCellValue("Формат даты");
        headerRow.createCell(5).setCellValue("Формат времени");
        headerRow.createCell(6).setCellValue("Описание ключа");

        addKeysToTemplate(sheet);

        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);
        sheet.autoSizeColumn(2);
        sheet.autoSizeColumn(3);
        sheet.autoSizeColumn(4);
        sheet.autoSizeColumn(5);
        sheet.autoSizeColumn(6);
    }

    private void addKeysToTemplate(XSSFSheet sheet) {
        List<String> keysList = new ArrayList<>(keysMap.keySet());
        for (int i = 0; i < keysList.size(); i++) {
            XSSFRow row = sheet.createRow(i + 1);
            String key = keysList.get(i);
            ReportKey reportKey = keysMap.get(keysList.get(i)).getReportKey();
            row.createCell(0).setCellValue(key);
            row.createCell(1).setCellValue(reportKey.keyType().name());
            row.createCell(2).setCellValue(reportKey.type().name());
            row.createCell(3).setCellValue(String.valueOf(reportKey.temporary()));
            if (reportKey.type() == DataType.DATE) {
                row.createCell(4).setCellValue(reportKey.dateFormatPattern());
            }
            if (reportKey.type() == DataType.TIME) {
                row.createCell(5).setCellValue(reportKey.timeFormatPattern());
            }
            row.createCell(6).setCellValue(keysMap.get(keysList.get(i)).getDescription());
        }
        XSSFRow row = sheet.createRow(keysList.size() + 1);
        row.createCell(0).setCellValue(COUNTER_KEY);
        row.createCell(6).setCellValue("This is a universal key that can be used in complex sheets to count the element number.");
    }

    private void checkCorrectnessFillingTemplate(XSSFWorkbook wb) throws IncorrectTemplateException {
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            XSSFSheet sheet = wb.getSheetAt(i);
            if (keysMap.containsKey(sheet.getSheetName())) {
                checkComplexKeyWithSheetName(sheet);
            } else if (!sheet.getSheetName().equals(DATABASE_KEY)) {
                checkCounterKeyOutsideComplex(sheet);
            }
        }
        for (String key : complexKeys) {
            checkLoopedKeys(wb.getSheet(key), new HashSet<>());
        }
    }

    //Проверяет наличие комплексного ключа, ссылающегося на этот же лист
    private void checkComplexKeyWithSheetName(XSSFSheet sheet) throws IncorrectTemplateException {
        List<XSSFCell> cellList = getCellListWithKey(sheet, 0, sheet.getLastRowNum());
        for (XSSFCell cell : cellList) {
            if (cell.getStringCellValue().equals(sheet.getSheetName())) {
                throw new IncorrectTemplateException("The sheet contains a complex key that references the same sheet.");
            }
        }
    }

    //Проверяет наличие ключа счетчика, который находится вне комплексного листа
    private void checkCounterKeyOutsideComplex(XSSFSheet sheet) throws IncorrectTemplateException {
        List<XSSFCell> cellList = getCellListWithKey(sheet, 0, sheet.getLastRowNum());
        for (XSSFCell cell : cellList) {
            if (cell.getStringCellValue().equals(COUNTER_KEY)) {
                throw new IncorrectTemplateException("The counter key is located outside the complex sheets.");
            }
        }
    }

    private void checkLoopedKeys(XSSFSheet sheet, Set<String> keys) throws IncorrectTemplateException {
        if (sheet != null) {
            if (keys.contains(sheet.getSheetName())) {
                throw new IncorrectTemplateException("Complex keys have looped references.");
            }
            keys.add(sheet.getSheetName());
            List<XSSFCell> cellList = getCellListWithKey(sheet, 0, sheet.getLastRowNum());
            for (XSSFCell cell : cellList) {
                String cellValue = cell.getStringCellValue();
                if (complexKeys.contains(cellValue)) {
                    checkLoopedKeys(sheet.getWorkbook().getSheet(cellValue), new HashSet<>(keys));
                }
            }
        }
    }

    private List<XSSFCell> getCellListWithKey(XSSFSheet sheet, int startRow, int endRow) {
        List<XSSFCell> ans = new ArrayList<>();
        for (int i = startRow; i <= endRow; i++) {
            XSSFRow row = sheet.getRow(i);
            if (row != null) {
                for (int j = 0; j < row.getLastCellNum(); j++) {
                    XSSFCell cell = row.getCell(j);
                    if (cell != null) {
                        if (cell.getCellType().equals(CellType.STRING)) {
                            if (keysMap.containsKey(cell.getStringCellValue()) || cell.getStringCellValue().equals(COUNTER_KEY)) {
                                ans.add(cell);
                            }
                        }
                    }
                }
            }
        }
        return ans;
    }

    private void fillCellByKey(XSSFCell cell, Object reportData, int count) throws IOException, IncorrectTemplateException, ReportKeyException {
        String key = cell.getStringCellValue();
        if (key.equals(COUNTER_KEY)) {
            cell.setCellValue(count);
        } else {
            ReportKey reportKey = keysMap.get(key).getReportKey();
            KeyType keyType = reportKey.keyType();
            if (keyType == KeyType.SINGLE) {
                String data = null;
                try {
                    Object fieldData = keysMap.get(key).getField().get(reportData);
                    if (fieldData != null) data = fieldData.toString();
                } catch (IllegalArgumentException e) {
                    throw new IncorrectTemplateException("Incorrect use of the single key. There is no access to the data object in this sheet.");
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                }

                if (reportKey.temporary() && data == null) {
                    deleteTemporaryKey(cell);
                } else {
                    setCellValue(cell, reportKey, data);
                }
            }
            if (keyType == KeyType.COMPLEX) {
                fillComplexKey(cell, reportData);
            }
        }
    }

    private void fillComplexKey(XSSFCell cell, Object reportData) throws IOException, IncorrectTemplateException, ReportKeyException {
        String key = cell.getStringCellValue();
        XSSFWorkbook wb = cell.getSheet().getWorkbook();
        XSSFSheet sourceSheet = wb.getSheet(key);
        XSSFSheet destinationSheet = cell.getSheet();

//        int rowIndex = cell.getRowIndex();
//        if (rowIndex == 0) destinationSheet.shiftRows(rowIndex, destinationSheet.getLastRowNum(), 1);

        if (sourceSheet != null) {
            List<?> list = null;
            if (keysMap.get(key).getField().getType() == List.class) {
                try {
                    list = (List<?>) keysMap.get(key).getField().get(reportData);
                } catch (IllegalArgumentException e) {
                    throw new IncorrectTemplateException("Incorrect use of the complex key. There is no access to the data object.");
                } catch (IllegalAccessException e) {
                    cell.setCellValue("");
                }
            }

            if (list != null && !list.isEmpty()) {
                List<XSSFRow> rowList = getRowList(sourceSheet);
                for (int i = 0; i < list.size(); i++) {
                    destinationSheet.shiftRows(cell.getRowIndex(), destinationSheet.getLastRowNum(), rowList.size());
                    destinationSheet.copyRows(rowList, cell.getRowIndex() - rowList.size(), new CellCopyPolicy());
                    List<XSSFCell> cellList = getCellListWithKey(destinationSheet,
                            cell.getRowIndex() - rowList.size(),
                            cell.getRowIndex() - 1);
                    for (XSSFCell xssfCell : cellList) {
                        fillCellByKey(xssfCell, list.get(i), i + 1);
                    }
                }
                deleteRow(cell.getRow());
            } else {
                if (keysMap.get(key).getReportKey().temporary()) {
                    deleteTemporaryKey(cell);
                } else {
                    cell.setCellValue("");
                }
            }
        }

    }

    private void setCellValue(XSSFCell cell, ReportKey reportKey, String data) throws IOException, FormatReportKeyException {
        try {
            if (data == null) {
                cell.setCellValue("");
            } else {
                DataType dataType = reportKey.type();
                switch (dataType) {
                    case DATE:
                        if (isStringCellFormat(cell)) {
                            SimpleDateFormat dateFormat = new SimpleDateFormat(reportKey.dateFormatPattern());
                            cell.setCellValue(dateFormat.format(new Date(Long.parseLong(data))));
                        } else {
                            cell.setCellValue(new Date(Long.parseLong(data)));
                        }
                        break;
                    case TIME:
                        if (isStringCellFormat(cell)) {
                            SimpleDateFormat timeFormat = new SimpleDateFormat(reportKey.timeFormatPattern());
                            cell.setCellValue(timeFormat.format(new Date(Long.parseLong(data))));
                        } else {
                            cell.setCellValue(new Date(Long.parseLong(data)));
                        }
                        break;
                    case NUMERIC:
                        cell.setCellValue(Double.parseDouble(data));
                        break;
                    case TEXT:
                        cell.setCellValue(data);
                        break;
                    case IMAGE:
                        cell.setCellValue("");
                        insertImage(cell, data);
                        break;
                }
            }
        } catch (NumberFormatException e) {
            throw new FormatReportKeyException("Field datatype does not match the ReportKey datatype.");
        }
    }

    private List<XSSFRow> getRowList(XSSFSheet sheet) {
        List<XSSFRow> ans = new ArrayList<>();
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            if (sheet.getRow(i) == null) {
                XSSFRow row = sheet.createRow(i);
                ans.add(row);
            } else {
                ans.add(sheet.getRow(i));
            }
        }
        return ans;
    }

    private void deleteKeySheets(XSSFWorkbook wb) {
        for (int i = wb.getNumberOfSheets() - 1; i >= 0; i--) {
            String sheetName = wb.getSheetAt(i).getSheetName();
            if (sheetName.equals(DATABASE_KEY) || keysMap.containsKey(sheetName)) {
                wb.removeSheetAt(i);
            }
        }
    }

    private void deleteTemporaryKey(XSSFCell cell) {
        if (isAvailableToDeleteKey(cell)) deleteRow(cell.getRow());
        else cell.setCellValue("");
    }

    private boolean isAvailableToDeleteKey(XSSFCell cell) {
        XSSFRow deletableRow = cell.getRow();
        for (CellRangeAddress mergedRegion : deletableRow.getSheet().getMergedRegions()) {
            if (!mergedRegion.isInRange(cell) && mergedRegion.containsRow(deletableRow.getRowNum())) {
                return false;
            }
        }
        Iterator<Cell> iterator = deletableRow.cellIterator();
        while (iterator.hasNext()) {
            XSSFCell xssfCell = (XSSFCell) iterator.next();
            if (xssfCell != cell && xssfCell.getCellType().equals(CellType.STRING)) {
                if (keysMap.containsKey(xssfCell.getStringCellValue())) {
                    return false;
                }
            }
        }
        return true;
    }

    private void deleteRow(XSSFRow row) {
        XSSFSheet sheet = row.getSheet();
        int index = row.getRowNum();
        if (sheet.getMergedRegions().size() != 0) {
            for (int i = 0; i < row.getLastCellNum(); i++) {
                if (row.getCell(i) != null) {
                    for (int j = sheet.getNumMergedRegions() - 1; j >= 0; j--) {
                        if (sheet.getMergedRegion(j).isInRange(row.getCell(i))) {
                            sheet.removeMergedRegion(j);
                        }
                    }
                }
            }
        }
        if (index == sheet.getLastRowNum()) {
            sheet.removeRow(row);
        } else {
            sheet.shiftRows(index + 1, sheet.getLastRowNum(), -1);
        }
    }

    private void insertImage(XSSFCell cell, String path) throws IOException {
        XSSFSheet sheet = cell.getSheet();
        XSSFWorkbook wb = sheet.getWorkbook();

        InputStream inputStream = new FileInputStream(path);
        byte[] bytes = inputStream.readAllBytes();
        int pictureIdx = wb.addPicture(bytes, XSSFWorkbook.PICTURE_TYPE_PNG);
        inputStream.close();

        XSSFCreationHelper helper = wb.getCreationHelper();
        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = helper.createClientAnchor();
        anchor.setCol1(cell.getColumnIndex());
        anchor.setRow1(cell.getRowIndex());

        XSSFPicture picture = drawing.createPicture(anchor, pictureIdx);
        autoSizeImage(picture, cell);
    }

    private void autoSizeImage(XSSFPicture picture, XSSFCell cell) {
        XSSFSheet sheet = cell.getRow().getSheet();
        CellRangeAddress range = null;
        for (CellRangeAddress mergedRegion : sheet.getMergedRegions()) {
            if (mergedRegion.isInRange(cell)) {
                range = mergedRegion;
            }
        }

        if (range == null) {
            double widthPxl = sheet.getColumnWidthInPixels(cell.getColumnIndex());
            double heightPxl = cell.getRow().getHeight() / PIXEL_TO_ROW_HEIGHT;
            autoSizeImageByRange(picture, widthPxl, heightPxl);
        } else {
            double widthPxl = 0;
            for (int i = range.getFirstColumn(); i <= range.getLastColumn(); i++) {
                widthPxl += sheet.getColumnWidthInPixels(i);
            }

            if (range.getFirstRow() == range.getLastRow()) {
                autoSizeImageByWidth(picture, widthPxl, cell);
            } else {
                double heightPxl = 0;
                for (int i = range.getFirstRow(); i <= range.getLastRow(); i++) {
                    heightPxl += sheet.getRow(i).getHeight() / PIXEL_TO_ROW_HEIGHT;
                }
                autoSizeImageByRange(picture, widthPxl, heightPxl);
            }
        }
    }

    private void autoSizeImageByRange(XSSFPicture picture, double width, double height) {
        picture.resize();
        double ratioWidth = width / picture.getImageDimension().width;
        double ratioHeight = height / picture.getImageDimension().height;
        if (ratioWidth < 1 || ratioHeight < 1) {
            picture.resize(Math.min(ratioHeight, ratioWidth));
        }
    }

    private void autoSizeImageByWidth(XSSFPicture picture, double width, XSSFCell cell) {
        double height;
        if (picture.getImageDimension().width <= width) {
            height = picture.getImageDimension().height * PIXEL_TO_ROW_HEIGHT;
            cell.getRow().setHeight((short) height);
            picture.resize();
        } else {
            double ratio = width / picture.getImageDimension().width;
            height = picture.getImageDimension().height * ratio * PIXEL_TO_ROW_HEIGHT;
            cell.getRow().setHeight((short) height);
            picture.resize();
            picture.resize(ratio);
        }
    }

    private boolean isStringCellFormat(XSSFCell cell) {
        String dataFormat = cell.getCellStyle().getDataFormatString();
        return dataFormat.equals("@") || dataFormat.equals("General");
    }

}


