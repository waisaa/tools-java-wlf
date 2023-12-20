package cn.wlftool.excel;

import cn.hutool.core.collection.ListUtil;
import cn.hutool.core.util.StrUtil;
import cn.wlftool.excel.constant.Const;
import cn.wlftool.excel.entity.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.util.*;

public class ExcelUtil {

    /**
     * 校验excel文件指定sheet
     *
     * @param filepath         文件路径
     * @param sheetNameOrIndex sheet名称或索引
     * @param tClass           接收数据的对象
     * @return 返回数据对象集合
     */
    public static boolean validTheSheet(String filepath, String validErrorPromptFilepath, String sheetNameOrIndex, Class<?> tClass) {
        return validTheSheet(filepath, validErrorPromptFilepath, sheetNameOrIndex, tClass, null);
    }

    /**
     * 校验excel文件指定sheet
     *
     * @param filepath         文件路径
     * @param sheetNameOrIndex sheet名称或索引
     * @param tClass           接收数据的对象
     * @return 返回数据对象集合
     */
    public static boolean validTheSheet(String filepath, String validErrorPromptFilepath, String sheetNameOrIndex, Class<?> tClass, Integer startRowIndex) {
        boolean valid;
        try (
                FileInputStream in = new FileInputStream(filepath);
                XSSFWorkbook workbook = new XSSFWorkbook(in);
                FileOutputStream fos = new FileOutputStream(validErrorPromptFilepath);
        ) {
            valid = validSheetValues(workbook, sheetNameOrIndex, tClass, startRowIndex);
            if (!valid) {
                workbook.write(fos);
            }
        } catch (Exception e) {
            e.printStackTrace();
            throw new RuntimeException("validAllSheet error: " + e.getMessage());
        }
        return valid;
    }

    /**
     * 校验excel文件中的所有sheet
     *
     * @return 返回每个sheet的校验结果，校验通过的结果为true，否则为false，且生成包含错误提示信息的文件
     * 错误提示信息：单元格被设置成了红色背景，并且鼠标选中提示错误信息
     */
    public static boolean validAllSheet(ExcelReadParam readParam) {
        boolean valid = true;
        String filepath = readParam.getFilepath();
        Map<String, Class<?>> sheetObject = readParam.getSheetObject();
        Map<String, Integer> startRowIndexes = readParam.getStartRowIndexes();
        try (
                FileInputStream in = new FileInputStream(filepath);
                XSSFWorkbook workbook = new XSSFWorkbook(in);
                FileOutputStream fos = new FileOutputStream(readParam.getValidErrorPromptFilepath());
        ) {
            int notValidCount = 0;
            for (String sheetNameOrIndex : sheetObject.keySet()) {
                Integer startRowIndex = null;
                if (null != startRowIndexes) {
                    startRowIndex = startRowIndexes.get(sheetNameOrIndex);
                }
                Class<?> tClass = sheetObject.get(sheetNameOrIndex);
                boolean validSheetValues = validSheetValues(workbook, sheetNameOrIndex, tClass, startRowIndex);
                if (!validSheetValues) {
                    notValidCount++;
                }
            }
            if (notValidCount > 0) {
                workbook.write(fos);
            }
        } catch (Exception e) {
            e.printStackTrace();
            throw new RuntimeException("validAllSheet error: " + e.getMessage());
        }
        return valid;
    }

    /**
     * 从excel文件指定sheet读取数据（有上下标信息字符串自动处理成富文本）
     *
     * @param filepath         文件路径
     * @param sheetNameOrIndex sheet名称或索引
     * @param tClass           接收数据的对象
     * @return 返回数据对象集合
     */
    public static <T> List<T> readTheSheet(String filepath, String sheetNameOrIndex, Class<T> tClass) {
        return readTheSheet(filepath, sheetNameOrIndex, tClass, null);
    }

    /**
     * 从excel文件指定sheet读取数据（有上下标信息字符串自动处理成富文本）
     *
     * @param filepath         文件路径
     * @param sheetNameOrIndex sheet名称或索引
     * @param startRowIndex    读取数据的起始行索引
     * @param tClass           接收数据的对象
     * @return 返回数据对象集合
     */
    public static <T> List<T> readTheSheet(String filepath, String sheetNameOrIndex, Class<T> tClass, Integer startRowIndex) {
        List<T> res = new ArrayList<>();
        try (
                FileInputStream in = new FileInputStream(filepath);
                XSSFWorkbook workbook = new XSSFWorkbook(in);
        ) {
            List<T> sheetValues = readSheetValues(workbook, sheetNameOrIndex, tClass, startRowIndex);
            res.addAll(sheetValues);
        } catch (Exception e) {
            e.printStackTrace();
            throw new RuntimeException("readTheSheet error: " + e.getMessage());
        }
        return res;
    }

    /**
     * 从excel文件读取所有sheet数据（有上下标信息字符串自动处理成富文本）
     * <p>
     * 注意：行转列
     *
     * @return 返回所有sheet数据对象集合
     */
    public static Map<String, List<?>> readAllSheet(ExcelReadParam readParam) {
        Map<String, List<?>> res = new HashMap<>();
        String filepath = readParam.getFilepath();
        Map<String, Class<?>> sheetObject = readParam.getSheetObject();
        Map<String, Integer> startRowIndexes = readParam.getStartRowIndexes();
        try (
                FileInputStream in = new FileInputStream(filepath);
                XSSFWorkbook workbook = new XSSFWorkbook(in);
        ) {
            for (String sheetNameOrIndex : sheetObject.keySet()) {
                Integer startRowIndex = null;
                if (null != startRowIndexes) {
                    startRowIndex = startRowIndexes.get(sheetNameOrIndex);
                }
                Class<?> tClass = sheetObject.get(sheetNameOrIndex);
                List<?> sheetValues = readSheetValues(workbook, sheetNameOrIndex, tClass, startRowIndex);
                res.put(sheetNameOrIndex, sheetValues);
            }
        } catch (Exception e) {
            e.printStackTrace();
            throw new RuntimeException("readAllSheet error: " + e.getMessage());
        }
        return res;
    }

    /**
     * 把数据写入到excel文件中，指定sheet、起始行索引和起始列索引（有上下标信息字符串自动处理成富文本）
     * <p>
     * 注意：没有表头的话不用配置表头信息
     *
     * @param writeParam 请求参数对象
     */
    public static void writeAllSheetWithHeader(ExcelWriteParam writeParam) {
        Map<String, List<?>> sheetValues = writeParam.getSheetValues();
        Map<String, Integer> startRowIndexes = writeParam.getStartRowIndexes();
        Map<String, List<SheetHeader>> headers = writeParam.getHeaders();
        try (
                FileInputStream fis = new FileInputStream(writeParam.getTplFilepath());
                XSSFWorkbook workbook = new XSSFWorkbook(fis);
                FileOutputStream fos = new FileOutputStream(writeParam.getFilepath());
        ) {
            for (String sheetNameOrIndex : sheetValues.keySet()) {
                Integer startRowIndex = null;
                if (null != startRowIndexes) {
                    startRowIndex = startRowIndexes.get(sheetNameOrIndex);
                }
                List<?> values = sheetValues.get(sheetNameOrIndex);
                if (null != headers) {
                    writeSheetHeaders(workbook, sheetNameOrIndex, headers.get(sheetNameOrIndex));
                }
                writeSheetValues(workbook, sheetNameOrIndex, values, startRowIndex);
            }
            workbook.write(fos);
        } catch (Exception e) {
            e.printStackTrace();
            throw new RuntimeException("write2Sheet error: " + e.getMessage());
        }
    }

    /**
     * 写入一个sheet的数据
     */
    private static void writeSheetValues(XSSFWorkbook workbook, String sheetNameOrIndex, List<?> values, Integer startRowIndex) {
        if (values.size() > 0) {
            XSSFSheet sheet = getSheet(workbook, sheetNameOrIndex);
            Object valueObj = values.get(0);
            Class<?> tClass = valueObj.getClass();
            if (null != startRowIndex) {
                XSSFCellStyle cellStyle = getDefaultCellStyle(workbook, null, false);
                Map<Integer, Field> fieldMap = getWriteFieldMap(tClass);
                for (Object t : values) {
                    XSSFRow row = sheet.getRow(startRowIndex);
                    if (null == row) {
                        row = sheet.createRow(startRowIndex);
                    }
                    for (Integer colIndex : fieldMap.keySet()) {
                        if (colIndex < 0) {
                            continue;
                        }
                        Field field = fieldMap.get(colIndex);
                        boolean numeric = field.getAnnotation(ExcelCol.class).numeric();
                        String defaultValue = field.getAnnotation(ExcelCol.class).defaultValue();
                        boolean scientificNotation = field.getAnnotation(ExcelCol.class).scientificNotation();
                        SheetCell sheetCell = SheetCell.builder()
                                .colIndex(colIndex)
                                .defaultValue(defaultValue)
                                .numeric(numeric)
                                .scientificNotation(scientificNotation)
                                .cellStyle(cellStyle)
                                .build();
                        String typeName = field.getType().getTypeName();
                        if (typeName.endsWith(JavaType.LIST.getType())) {
                            List<Object> headerValues = new ArrayList<>();
                            try {
                                headerValues = (List<Object>) field.get(t);
                            } catch (Exception ignored) {
                            }
                            for (int i = 0; i < headerValues.size(); i++) {
                                int ci = colIndex + i;
                                Object val = headerValues.get(i);
                                String content = null == val ? Const.PTN_EMPTY : String.valueOf(val);
                                sheetCell.setColIndex(ci);
                                sheetCell.setContent(content);
                                writeCell(workbook, row, sheetCell);
                            }
                        } else {
                            String content;
                            try {
                                content = field.get(t).toString();
                            } catch (Exception e) {
                                content = Const.PTN_EMPTY;
                            }
                            sheetCell.setContent(content);
                            writeCell(workbook, row, sheetCell);
                        }
                    }
                    startRowIndex++;
                }
            } else {
                Map<Field, SheetRegion> regionFieldMap = getRegionFieldMap(tClass);
                for (Field field : regionFieldMap.keySet()) {
                    boolean bold = field.getAnnotation(ExcelRegion.class).bold();
                    boolean numeric = field.getAnnotation(ExcelRegion.class).numeric();
                    String defaultValue = field.getAnnotation(ExcelRegion.class).defaultValue();
                    boolean scientificNotation = field.getAnnotation(ExcelRegion.class).scientificNotation();
                    XSSFCellStyle cellStyle = getDefaultCellStyle(workbook, null, bold);
                    SheetCell sheetCell = SheetCell.builder()
                            .defaultValue(defaultValue)
                            .numeric(numeric)
                            .scientificNotation(scientificNotation)
                            .cellStyle(cellStyle)
                            .build();
                    SheetRegion sheetRegion = regionFieldMap.get(field);
                    int sri = sheetRegion.getStartRowIndex();
                    int sci = sheetRegion.getStartColIndex();
                    List<List<String>> rowValues = new ArrayList<>();
                    try {
                        rowValues = (List<List<String>>) field.get(valueObj);
                    } catch (Exception ignored) {
                    }
                    if (rowValues.size() > 0) {
                        int eri = rowValues.size() - 1 + sri;
                        for (int rowIndex = sri; rowIndex <= eri; rowIndex++) {
                            List<String> rowValue = rowValues.get(rowIndex - sri);
                            XSSFRow row = sheet.getRow(rowIndex);
                            if (null == row) {
                                row = sheet.createRow(rowIndex);
                            }
                            int eci = rowValue.size() - 1 + sci;
                            for (int colIndex = sci; colIndex <= eci; colIndex++) {
                                Object val = rowValue.get(colIndex - sci);
                                String content = null == val ? Const.PTN_EMPTY : String.valueOf(val);
                                sheetCell.setColIndex(colIndex);
                                sheetCell.setContent(content);
                                writeCell(workbook, row, sheetCell);
                            }
                        }
                    }
                }
            }
        }
    }

    /**
     * 校验一个sheet的数据
     */
    private static boolean validSheetValues(XSSFWorkbook workbook, String sheetNameOrIndex, Class<?> tClass, Integer startRowIndex) {
        boolean res = true;
        XSSFSheet sheet = getSheet(workbook, sheetNameOrIndex);
        if (null != startRowIndex) {
            Map<Integer, Field> fieldMap = getReadFieldMap(tClass);
            if (fieldMap.size() > 0) {
                for (int rowIndex = startRowIndex; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                    Row row = sheet.getRow(rowIndex);
                    if (null != row) {
                        for (Integer colIndex : fieldMap.keySet()) {
                            Field field = fieldMap.get(colIndex);
                            String[] validValueRange = field.getAnnotation(ExcelCol.class).validValueRange();
                            String validErrorPrompt = field.getAnnotation(ExcelCol.class).validErrorPrompt();
                            if (colIndex > 0) {
                                Cell cell = row.getCell(colIndex);
                                if (cell != null) {
                                    String typeName = field.getType().getName();
                                    String cellValue = getCellValue(cell, workbook, typeName, validValueRange);
                                    if (null == cellValue) {
                                        res = false;
                                        XSSFCellStyle cellStyle = getDefaultCellStyle(workbook, IndexedColors.RED, false);
                                        setCellPrompt(sheet, validErrorPrompt, rowIndex, colIndex);
                                        cell.setCellStyle(cellStyle);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        } else {
            Map<Field, SheetRegion> regionFieldMap = getRegionFieldMap(tClass);
            List<Integer> sheetMaxRowAndColIndex = getSheetMaxRowAndColIndex(sheet);
            int maxRowIndex = sheetMaxRowAndColIndex.get(0), maxColIndex = sheetMaxRowAndColIndex.get(1);
            for (Field field : regionFieldMap.keySet()) {
                SheetRegion sheetRegion = regionFieldMap.get(field);
                int sri = sheetRegion.getStartRowIndex();
                int eri = sheetRegion.getEndRowIndex();
                eri = eri == -1 ? maxRowIndex : eri;
                int sci = sheetRegion.getStartColIndex();
                int eci = sheetRegion.getEndColIndex();
                eci = eci == -1 ? maxColIndex : eci;
                String[] validValueRange = field.getAnnotation(ExcelRegion.class).validValueRange();
                String validErrorPrompt = field.getAnnotation(ExcelRegion.class).validErrorPrompt();
                for (int rowIndex = sri; rowIndex <= eri; rowIndex++) {
                    Row row = sheet.getRow(rowIndex);
                    if (null != row) {
                        for (int colIndex = sci; colIndex <= eci; colIndex++) {
                            Cell cell = row.getCell(colIndex);
                            if (null != cell) {
                                String cellValue = getCellValue(cell, workbook, JavaType.STRING.getType(), validValueRange);
                                if (null == cellValue) {
                                    res = false;
                                    XSSFCellStyle cellStyle = getDefaultCellStyle(workbook, IndexedColors.RED, false);
                                    setCellPrompt(sheet, validErrorPrompt, rowIndex, colIndex);
                                    cell.setCellStyle(cellStyle);
                                }
                            }
                        }
                    }
                }
            }
        }
        return res;
    }

    /**
     * 获取一个sheet的数据
     */
    private static <T> List<T> readSheetValues(XSSFWorkbook workbook, String sheetNameOrIndex, Class<T> tClass, Integer startRowIndex) throws Exception {
        List<T> res = new ArrayList<>();
        XSSFSheet sheet = getSheet(workbook, sheetNameOrIndex);
        if (null != startRowIndex) {
            Map<Integer, Field> fieldMap = getReadFieldMap(tClass);
            if (fieldMap.size() > 0) {
                List<Integer> sheetMaxRowAndColIndex = getSheetMaxRowAndColIndex(sheet);
                int maxRowIndex = sheetMaxRowAndColIndex.get(0), maxColIndex = sheetMaxRowAndColIndex.get(1);
                for (int rowIndex = startRowIndex; rowIndex <= maxRowIndex; rowIndex++) {
                    Row row = sheet.getRow(rowIndex);
                    T rowValues = getRowValues(row, fieldMap, tClass, workbook);
                    if (null != rowValues) {
                        res.add(rowValues);
                    }
                }
            }
        } else {
            Map<Field, SheetRegion> regionFieldMap = getRegionFieldMap(tClass);
            T t = tClass.newInstance();
            List<Integer> sheetMaxRowAndColIndex = getSheetMaxRowAndColIndex(sheet);
            int maxRowIndex = sheetMaxRowAndColIndex.get(0), maxColIndex = sheetMaxRowAndColIndex.get(1);
            for (Field field : regionFieldMap.keySet()) {
                SheetRegion sheetRegion = regionFieldMap.get(field);
                int sri = sheetRegion.getStartRowIndex();
                int eri = sheetRegion.getEndRowIndex();
                eri = eri == -1 ? maxRowIndex : eri;
                int sci = sheetRegion.getStartColIndex();
                int eci = sheetRegion.getEndColIndex();
                eci = eci == -1 ? maxColIndex : eci;
                boolean numeric = field.getAnnotation(ExcelRegion.class).numeric();
                String[] validValueRange = field.getAnnotation(ExcelRegion.class).validValueRange();
                String typeName = numeric ? JavaType.DOUBLE.getType() : JavaType.STRING.getType();
                List<List<String>> rowValues = new ArrayList<>();
                for (int rowIndex = sri; rowIndex <= eri; rowIndex++) {
                    List<String> rowValue = new ArrayList<>();
                    Row row = sheet.getRow(rowIndex);
                    if (null != row) {
                        for (int colIndex = sci; colIndex <= eci; colIndex++) {
                            Cell cell = row.getCell(colIndex);
                            if (null != cell) {
                                String cellValue = getCellValue(cell, workbook, typeName, validValueRange);
                                rowValue.add(cellValue);
                            } else {
                                rowValue.add(Const.PTN_EMPTY);
                            }
                        }
                    } else {
                        for (int colIndex = sci; colIndex <= eci; colIndex++) {
                            rowValue.add(Const.PTN_EMPTY);
                        }
                    }
                    rowValues.add(rowValue);
                }
                field.set(t, rowValues);
            }
            res.add(t);
        }
        return res;
    }

    /**
     * 获取sheet的最大行、列索引
     */
    private static List<Integer> getSheetMaxRowAndColIndex(XSSFSheet sheet) {
        // 返回最后一行的索引，即 比行总数小1
        int maxRowIndex = sheet.getLastRowNum(), maxColIndex = 0;
        for (int i = 0; i < maxRowIndex; i++) {
            XSSFRow row = sheet.getRow(i);
            if (null != row) {
                // 返回的是最后一列的列数，即 等于总列数
                short lastCellNum = row.getLastCellNum();
                if (lastCellNum > maxColIndex) {
                    maxColIndex = lastCellNum - 1;
                }
            }
        }
        return ListUtil.toList(maxRowIndex, maxColIndex);
    }

    /**
     * 写入表头数据
     */
    private static void writeSheetHeaders(XSSFWorkbook workbook, String sheetNameOrIndex, List<SheetHeader> headers) {
        XSSFSheet sheet = getSheet(workbook, sheetNameOrIndex);
        XSSFCellStyle cellStyle = getDefaultCellStyle(workbook, null, true);
        for (SheetHeader sheetHeader : headers) {
            int rowIndex = sheetHeader.getRowIndex();
            int startColIndex = sheetHeader.getStartColIndex();
            List<String> headerList = sheetHeader.getHeaders();
            XSSFRow row = sheet.getRow(rowIndex);
            if (null == row) {
                row = sheet.createRow(rowIndex);
            }
            for (int i = 0; i < headerList.size(); i++) {
                String header = headerList.get(i);
                int ci = startColIndex + i;
                XSSFCell cell = row.getCell(ci);
                if (null == cell) {
                    cell = row.createCell(ci);
                }
                cell.setCellValue(header);
                cell.setCellStyle(cellStyle);
            }
        }
    }

    /**
     * 向一个单元格写入数据
     */
    private static void writeCell(XSSFWorkbook workbook, XSSFRow row, SheetCell sheetCell) {
        int colIndex = sheetCell.getColIndex();
        String content = sheetCell.getContent();
        String defaultValue = sheetCell.getDefaultValue();
        boolean numeric = sheetCell.isNumeric();
        boolean scientificNotation = sheetCell.isScientificNotation();
        XSSFCellStyle cellStyle = sheetCell.getCellStyle();
        List<List<int[]>> tagIndexArr = new ArrayList<>();
        if (containTag(content)) {
            content = getIndexes(content, tagIndexArr);
        }
        XSSFCell cell = row.getCell(colIndex);
        if (null == cell) {
            cell = row.createCell(colIndex);
        }
        if (!Const.S_NULL.equals(defaultValue)) {
            cell.setCellValue(defaultValue);
        } else if (tagIndexArr.size() > 0) {
            cell.setCellValue(richTextString(workbook, content, tagIndexArr));
        } else {
            if (numeric) {
                if (scientificNotation) {
                    XSSFDataFormat dataFormat = workbook.createDataFormat();
                    cellStyle.setDataFormat(dataFormat.getFormat(Const.FMT_SCIENTIFIC_NOTATION));
                }
                try {
                    cell.setCellValue(Double.parseDouble(content));
                } catch (Exception ignored) {
                    cell.setCellValue(content);
                }
            } else {
                cell.setCellValue(content);
            }
        }
        cell.setCellStyle(cellStyle);
    }

    /**
     * 获取默认格式的单元格样式
     */
    private static XSSFCellStyle getDefaultCellStyle(XSSFWorkbook workbook, IndexedColors color, boolean bold) {
        XSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        XSSFFont font = workbook.createFont();
        font.setFontName(Const.FONT_TNR);
        font.setBold(bold);
        cellStyle.setFont(font);
        XSSFDataFormat dataFormat = workbook.createDataFormat();
        cellStyle.setDataFormat(dataFormat.getFormat(Const.FMT_TEXT));
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        if (null != color) {
            cellStyle.setFillForegroundColor(color.getIndex());
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
        return cellStyle;
    }

    /**
     * 设置单元格上提示
     */
    public static void setCellPrompt(XSSFSheet sheet, String promptContent, int rowIndex, int colIndex) {
        String promptTitle = Const.S_TITLE;
        XSSFDataValidationHelper helper = new XSSFDataValidationHelper(sheet);
        DataValidationConstraint constraint = helper.createCustomConstraint(Const.S_A0);
        CellRangeAddressList region = new CellRangeAddressList(rowIndex, rowIndex, colIndex, colIndex);
        DataValidation validation = helper.createValidation(constraint, region);
        validation.createPromptBox(promptTitle, promptContent);
        validation.setShowPromptBox(true);
        sheet.addValidationData(validation);
    }

    /**
     * 获取sheet
     */
    private static XSSFSheet getSheet(XSSFWorkbook workbook, String sheetNameOrIndex) {
        XSSFSheet sheet;
        try {
            int sheetIndex = Integer.parseInt(sheetNameOrIndex);
            sheet = workbook.getSheetAt(sheetIndex);
        } catch (Exception e) {
            sheet = workbook.getSheet(sheetNameOrIndex);
        }
        return sheet;
    }

    /**
     * 获取一行的数据
     */
    private static <T> T getRowValues(Row row, Map<Integer, Field> fieldMap, Class<T> tClass, XSSFWorkbook workbook) throws Exception {
        T res = null;
        if (null != row) {
            T t = tClass.newInstance();
            for (Integer colIndex : fieldMap.keySet()) {
                Field field = fieldMap.get(colIndex);
                if (colIndex < 0) {
                    if (colIndex == -1) {
                        String defaultValue = field.getAnnotation(ExcelCol.class).defaultValue();
                        setFieldValue(field, defaultValue, t);
                    }
                } else {
                    Cell cell = row.getCell(colIndex);
                    if (cell != null) {
                        String[] validValueRange = field.getAnnotation(ExcelCol.class).validValueRange();
                        String typeName = field.getType().getName();
                        String cellValue = getCellValue(cell, workbook, typeName, validValueRange);
                        setFieldValue(field, cellValue, t);
                    }
                }
            }
            if (!isAllFieldNull(t)) {
                res = t;
            }
        }
        return res;
    }

    /**
     * @param val 数值
     * @return 返回科学计数法字符串
     */
    public static String scientificNotationString(Double val) {
        if (val == null || Double.isNaN(val)) {
            return null;
        }
        String res = new DecimalFormat(Const.FMT_DOUBLE).format(val);
        if (val >= 1 || val == 0) {
            int length = res.length();
            String prefix = res.substring(0, length - 2);
            String suffix = res.substring(length - 2, length);
            res = String.format(Const.TPL_E1, prefix, suffix);
        }
        return res;
    }

    /**
     * 科学计数法转普通数值字符串
     */
    public static String fromScientificNotation(Double val) {
        if (val == null || Double.isNaN(val)) {
            return null;
        }
        String scientificNotationString = scientificNotationString(val);
        String str = new BigDecimal(scientificNotationString).toPlainString();
        int len = str.length();
        return Const.S_0.equals(String.valueOf(str.charAt(len - 1))) ? str.substring(0, len - 1) : str;
    }

    /**
     * 读数据时获取字段集合信息
     */
    private static Map<Integer, Field> getReadFieldMap(Class<?> tClass) {
        Map<Integer, Field> fieldMap = new HashMap<>();
        if (null != tClass) {
            for (Field field : tClass.getDeclaredFields()) {
                ExcelCol col = field.getAnnotation(ExcelCol.class);
                if (null != col) {
                    field.setAccessible(true);
                    fieldMap.put(col.readIndex(), field);
                }
            }
        }
        return fieldMap;
    }

    /**
     * 读写数据时获取区域字段集合信息
     */
    private static Map<Field, SheetRegion> getRegionFieldMap(Class<?> tClass) {
        Map<Field, SheetRegion> fieldMap = new HashMap<>();
        if (null != tClass) {
            for (Field field : tClass.getDeclaredFields()) {
                ExcelRegion region = field.getAnnotation(ExcelRegion.class);
                if (null != region) {
                    field.setAccessible(true);
                    SheetRegion sheetRegion = SheetRegion.builder()
                            .startRowIndex(region.startRowIndex())
                            .endRowIndex(region.endRowIndex())
                            .startColIndex(region.startColIndex())
                            .endColIndex(region.endColIndex())
                            .build();
                    fieldMap.put(field, sheetRegion);
                }
            }
        }
        return fieldMap;
    }

    /**
     * 写数据时获取字段集合信息
     */
    private static Map<Integer, Field> getWriteFieldMap(Class<?> tClass) {
        Map<Integer, Field> fieldMap = new HashMap<>();
        if (null != tClass) {
            for (Field field : tClass.getDeclaredFields()) {
                ExcelCol col = field.getAnnotation(ExcelCol.class);
                if (null != col) {
                    field.setAccessible(true);
                    fieldMap.put(col.writeStartIndex(), field);
                }
            }
        }
        return fieldMap;
    }

    /**
     * 判断对象的所有字段值是否为空
     */
    private static <T> boolean isAllFieldNull(T t) {
        boolean res = true;
        Class<?> tClass = t.getClass();
        for (Field field : tClass.getDeclaredFields()) {
            field.setAccessible(true);
            try {
                Object fieldValue = field.get(t);
                if (!Const.FIELDS_DEFAULT.contains(field.getName()) && null != fieldValue && !Const.PTN_BAR_MID.equals(fieldValue)) {
                    res = false;
                }
            } catch (Exception ignored) {
            }
        }
        return res;
    }

    /**
     * 设置字段值
     */
    private static <T> void setFieldValue(Field field, String value, T t) throws Exception {
        if (null != field) {
            String type = field.getType().toString();
            if (null == value || StrUtil.isBlank(value) || Const.PTN_BAR_MID.equals(value)) {
                field.set(t, null);
            } else if (type.endsWith(JavaType.INTEGER.getType())) {
                field.set(t, Integer.parseInt(value));
            } else if (type.endsWith(JavaType.DOUBLE.getType())) {
                Double v = null;
                try {
                    v = Double.parseDouble(value);
                } catch (Exception ignored) {

                }
                field.set(t, v);
            } else {
                field.set(t, value);
            }
        }
    }

    /**
     * 自动处理cell内容（有上下标信息字符串自动处理成富文本）
     */
    private static String getCellValue(Cell cell, XSSFWorkbook workbook, String typeName, String[] validValueRange) {
        String res;
        CellType cellType = cell.getCellType();
        if (CellType.NUMERIC.equals(cellType)) {
            double cellValue = cell.getNumericCellValue();
            if (typeName.endsWith(JavaType.STRING.getType())) {
                res = String.valueOf((long) cellValue);
            } else {
                res = String.valueOf(cellValue);
            }
        } else if (CellType.STRING.equals(cellType)) {
            String stringCellValue = cell.getStringCellValue();
            if (validValueRange.length == 0 || Arrays.asList(validValueRange).contains(stringCellValue)) {
                XSSFFont font;
                XSSFRichTextString rts = (XSSFRichTextString) cell.getRichStringCellValue();
                if (rts.numFormattingRuns() > 1) {
                    StringBuilder value = new StringBuilder();
                    for (int i = 0; i < rts.numFormattingRuns(); i++) {
                        int runLength = rts.getLengthOfFormattingRun(i);
                        int runIndex = rts.getIndexOfFormattingRun(i);
                        String temp = rts.toString().substring(runIndex, (runIndex + runLength));
                        try {
                            font = rts.getFontOfFormattingRun(i);
                            font.getTypeOffset();
                        } catch (NullPointerException e) {
                            font = workbook.getFontAt(XSSFFont.DEFAULT_CHARSET);
                            font.setTypeOffset(XSSFFont.SS_NONE);
                        }
                        temp = addTagInfo(temp, font.getTypeOffset());
                        value.append(temp);
                    }
                    res = value.toString();
                } else {
                    res = stringCellValue;
                }
            } else {
                res = null;
            }
        } else {
            res = Const.PTN_EMPTY;
        }
        return res;
    }

    /**
     * 处理有上下标的字符串
     */
    private static String addTagInfo(String str, short typeOffset) {
        if (typeOffset == XSSFFont.SS_SUPER) {
            str = String.format(Const.TPL_TAG, Const.TAG_SUP_START, str, Const.TAG_SUP_END);
        }
        if (typeOffset == XSSFFont.SS_SUB) {
            str = String.format(Const.TPL_TAG, Const.TAG_SUB_START, str, Const.TAG_SUB_END);
        }
        return str;
    }

    /**
     * 有上下标信息字符串处理成富文本
     *
     * @param str 字符串
     * @return 处理后的富文本
     */
    private static XSSFRichTextString richTextString(XSSFWorkbook workbook, String str, List<List<int[]>> tagIndexArr) {
        XSSFRichTextString richTextString = new XSSFRichTextString(str);
        List<int[]> subs = tagIndexArr.get(0);
        List<int[]> sups = tagIndexArr.get(1);
        if (subs.size() > 0) {
            XSSFFont font = workbook.createFont();
            font.setTypeOffset(XSSFFont.SS_SUB);
            for (int[] pair : subs) {
                richTextString.applyFont(pair[0], pair[1], font);
            }
        }
        if (sups.size() > 0) {
            XSSFFont font = workbook.createFont();
            font.setTypeOffset(XSSFFont.SS_SUPER);
            for (int[] pair : sups) {
                richTextString.applyFont(pair[0], pair[1], font);
            }
        }
        return richTextString;
    }

    /**
     * 获取下一对标签的index，不存在这些标签就返回null
     *
     * @param str 字符串
     * @param tag SUB_START或者SUP_START
     * @return int[]中有两个元素，第一个是开始标签的index，第二个元素是结束标签的index
     */
    private static int[] getNextTagsIndex(String str, String tag) {
        int firstStart = str.indexOf(tag);
        if (firstStart > -1) {
            int firstEnd = 0;
            if (tag.equals(Const.TAG_SUB_START)) {
                firstEnd = str.indexOf(Const.TAG_SUB_END);
            } else if (tag.equals(Const.TAG_SUP_START)) {
                firstEnd = str.indexOf(Const.TAG_SUP_END);
            }
            if (firstEnd > firstStart) {
                return new int[]{firstStart, firstEnd};
            }
        }
        return new int[]{};
    }

    /**
     * 移除下一对sub或者sup或者u或者strong或者em标签
     *
     * @param str 字符串
     * @param tag SUB_START或者SUP_START
     * @return 返回移除后的字符串
     */
    private static String removeNextTags(String str, String tag) {
        str = str.replaceFirst(tag, Const.PTN_EMPTY);
        if (tag.equals(Const.TAG_SUB_START)) {
            str = str.replaceFirst(Const.TAG_SUB_END, Const.PTN_EMPTY);
        } else if (tag.equals(Const.TAG_SUP_START)) {
            str = str.replaceFirst(Const.TAG_SUP_END, Const.PTN_EMPTY);
        }
        return str;
    }

    /**
     * 判断是不是包含sub、sup标签
     *
     * @param str 字符串
     * @return 返回是否包含
     */
    private static boolean containTag(String str) {
        return (str.contains(Const.TAG_SUB_START) && str.contains(Const.TAG_SUB_END)) || (str.contains(Const.TAG_SUP_START) && str.contains(Const.TAG_SUP_END));
    }

    /**
     * 处理字符串，得到每个sub、sup标签的开始和对应的结束的标签的index，方便后面根据这个标签做字体操作
     *
     * @param str          字符串
     * @param tagIndexList 传一个新建的空list进来，方法结束的时候会存储好标签位置信息。
     *                     <br>tagIndexList.get(0)存放的sub
     *                     <br>tagIndexList.get(1)存放的是sup
     * @return 返回sub、sup处理完之后的字符串
     */
    private static String getIndexes(String str, List<List<int[]>> tagIndexList) {
        List<int[]> subs = new ArrayList<>(), sups = new ArrayList<>();
        while (true) {
            int[] sub_pair = getNextTagsIndex(str, Const.TAG_SUB_START), sup_pair = getNextTagsIndex(str, Const.TAG_SUP_START);
            boolean subFirst = false, supFirst = false;
            List<Integer> a = new ArrayList<>();
            if (sub_pair.length > 0) {
                a.add(sub_pair[0]);
            }
            if (sup_pair.length > 0) {
                a.add(sup_pair[0]);
            }
            Collections.sort(a);
            if (sub_pair.length > 0) {
                if (sub_pair[0] == Integer.parseInt(a.get(0).toString())) {
                    subFirst = true;
                }
            }
            if (sup_pair.length > 0) {
                if (sup_pair[0] == Integer.parseInt(a.get(0).toString())) {
                    supFirst = true;
                }
            }
            if (subFirst) {
                str = removeNextTags(str, Const.TAG_SUB_START);
                // <sub>标签被去掉之后，结束标签需要相应往前移动
                sub_pair[1] = sub_pair[1] - Const.TAG_SUB_START.length();
                subs.add(sub_pair);
                continue;
            }
            if (supFirst) {
                str = removeNextTags(str, Const.TAG_SUP_START);
                // <sup>标签被去掉之后，结束标签需要相应往前移动
                sup_pair[1] = sup_pair[1] - Const.TAG_SUP_START.length();
                sups.add(sup_pair);
                continue;
            }
            if (sub_pair.length == 0 && sup_pair.length == 0) {
                break;
            }
        }
        tagIndexList.add(subs);
        tagIndexList.add(sups);
        return str;
    }

}
