package com.deepoove.poi.util;

import com.deepoove.poi.data.RenderData;
import com.deepoove.poi.exception.RenderException;
import com.deepoove.poi.render.compute.ReadMapAccessor;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.expression.ExpressionParser;
import org.springframework.expression.common.TemplateParserContext;
import org.springframework.expression.spel.SpelCompilerMode;
import org.springframework.expression.spel.SpelParserConfiguration;
import org.springframework.expression.spel.standard.SpelExpressionParser;
import org.springframework.expression.spel.support.StandardEvaluationContext;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@SuppressWarnings("unused")
public class ExcelTemplateRenderer {

    private static final Logger log = LoggerFactory.getLogger(ExcelTemplateRenderer.class);
    private final ExpressionParser parser;

    private Workbook workbook;

    // 动态行数据名称 占位符号 [$xxxx]
    private static final Pattern DYNAMIC_ROW_DATA_PLACEHOLDOR = Pattern.compile("(\\[\\s*\\$.*?\\s*\\])");
    TemplateParserContext templateRowContext = new TemplateParserContext("[", "]");
    // 行书写占位符号  [ xxx ]
    private static final Pattern ROW_PLACEHOLDOR = Pattern.compile("(\\[\\s*.*?\\s*\\])");

    {
        ClassLoader loader = getClass().getClassLoader();
        SpelParserConfiguration config = new SpelParserConfiguration(SpelCompilerMode.IMMEDIATE, loader,
            true, true, 1024);
        this.parser = new SpelExpressionParser(config);
    }

    public void render(String templatePath, Map<String, Object> model)
        throws Exception {
        InputStream inputStream = Files.newInputStream(Paths.get(templatePath));
        this.render(inputStream, model);
        inputStream.close();
    }

    public void render(InputStream inputStream, Map<String, Object> model)
        throws Exception {
        if (inputStream == null) {
            throw new RenderException("InputStream is null, unable to render. ");
        }
        this.workbook = WorkbookFactory.create(inputStream);

        for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            processSheet(sheet, model);
        }
    }

    @SuppressWarnings("unchecked")
    private void processSheet(Sheet sheet, Map<String, Object> model) {
        StandardEvaluationContext context = new StandardEvaluationContext(model);
        context.addPropertyAccessor(new ReadMapAccessor());
        int rowNumber = 0;
        // Traverse the entire sheet and process the data dynamic rendering.
        while (true) {
            Row row = sheet.getRow(rowNumber);
            if (sheet.getLastRowNum() < rowNumber) {
                break;
            }
            if (row == null) {
                rowNumber++;
                continue;
            }
            /*
              1. Traverse the entire row
                   Use DYNAMIC_ROWDATA-PLACEHOLDOR to find the data placeholder symbol and its corresponding data.
                   If the array data exists, render the row. If the array data does not exist, clear the row data.
                   Find the data placeholder symbol and clear it.
                   Find all the [] placeholder symbols, record the content of each cell, and the corresponding placeholder symbol.
             */
            Map<Integer, CellData> templateRowData = new HashMap<>();
            List<Object> dataList = null;
            boolean isDynamicRow = false;
            for (Cell cell : row) {
                if (cell.getCellType() == CellType.STRING) {
                    CellData cellData = new CellData();
                    String value = cell.getStringCellValue();
                    Matcher matcher = DYNAMIC_ROW_DATA_PLACEHOLDOR.matcher(value);
                    if (matcher.find()) {
                        isDynamicRow = true;
                        String variableName = matcher.group(1);
                        String tempVariableName = variableName.replaceFirst("\\$", "");
                        Object tempObj = this.execExpression(tempVariableName, context, templateRowContext);
                        if (tempObj == null) {
                            log.warn("Dynamic row data is null: {}", variableName);
                        } else if (tempObj instanceof List<?>) {
                            dataList = (List<Object>) tempObj;
                        } else if (tempObj.getClass().isArray()) {
                            dataList = Collections.singletonList(tempObj);
                        } else {
                            log.warn("Dynamic row data is not a list: {}", variableName);
                            dataList = null;
                        }
                        value = value.replace(variableName, "");
                    }

                    Matcher dataMatcher = ROW_PLACEHOLDOR.matcher(value);
                    while (dataMatcher.find()) {
                        String placeholder = dataMatcher.group(1);
                        cellData.addPlaceholdersEl(placeholder);
                        value = value.replace(placeholder, "%s");
                    }

                    cellData.setRow(rowNumber);
                    cellData.setTemplateValue(value);
                    templateRowData.put(cell.getColumnIndex(), cellData);
                }
            }
            // 2、 Processing dynamic row rendering: Traverse the data rendering rows based on the found array data and
            // the corresponding row and column expressions.
            if (dataList == null) {
                if (isDynamicRow) {
                    ExcelUtils.removeOneRowAndShift(sheet, row);
                    rowNumber--;
                }
                rowNumber++;
                continue;
            }
            for (Object data : dataList) {
                Map<String, Object> map = null;
                Row newRow = ExcelUtils.insertRowAndShiftBelow(sheet, ++rowNumber);
                ExcelUtils.copyRow(row, newRow, true);
                Map<String, Object> variables = (Map<String, Object>) ReflectionUtils.getValue("variables", context);
                variables.clear();
                if (data != null) {
                    try {
                        TlBeanUtil beanUtil = new TlBeanUtil();
                        if (!(data instanceof String || TlBeanUtil.isPrimitive(data))) {
                            map = beanUtil.beanToMap(data, RenderData.class, 0);
                            model.putAll(map);
                        }
                    } catch (Exception ignore) {
                    }
                }

                for (Cell newCell : newRow) {
                    CellData cellData = templateRowData.get(newCell.getColumnIndex());
                    if (cellData == null) {
                        continue;
                    }
                    Object[] valueList = new Object[cellData.getPlaceholdersEl().size()];
                    for (int i = 0; i < cellData.getPlaceholdersEl().size(); i++) {
                        String placeholder = cellData.getPlaceholdersEl().get(i);
                        Object value = this.execExpression(placeholder, context, templateRowContext);
                        valueList[i] = value;
                    }
                    newCell.setCellValue(String.format(cellData.getTemplateValue(), valueList));
                }
                if (map != null && !map.isEmpty()) {
                    map.forEach((key, value) -> model.remove(key));
                }
            }
            ExcelUtils.removeOneRowAndShift(sheet, row);
        }

        // Traverse the entire sheet for static placeholder {{}} symbol processing
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() == CellType.STRING) {
                    String value = cell.getStringCellValue();
                    if (value.contains("{{") && value.contains("}}")) {
                        Object newValue = execExpression(value, context, null);
                        ExcelUtils.setCellValue(cell, newValue);
                    }
                }
            }
        }
    }

    private Object execExpression(String el, StandardEvaluationContext context, TemplateParserContext parserContext) {
        if (parserContext == null) {
            parserContext = new TemplateParserContext("{{", "}}");
        }
        try {
            return parser.parseExpression(el, parserContext).getValue(context);
        } catch (Exception e) {
            log.error("Error replacing placeholders 【{}】 in Excel template: {}", el, e.getMessage());
            return null;
        }
    }

    public Workbook getWorkBook() {
        return workbook;
    }

    public void save(String outputFilePath) {
        try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
            workbook.write(fos);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public byte[] getBytes() {
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            workbook.write(bos);
            return bos.toByteArray();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    protected static class CellData {
        private Integer row;
        private Integer col;
        // 占位符好模板字符串
        private String templateValue;
        // 占位符号表达式列表
        private List<String> placeholdersEl = new ArrayList<>();

        public CellData() {
        }

        public CellData(Integer row, Integer col, String value) {
            this.row = row;
            this.col = col;
            this.templateValue = value;
        }

        public Integer getRow() {
            return row;
        }

        public void setRow(Integer row) {
            this.row = row;
        }

        public Integer getCol() {
            return col;
        }

        public void setCol(Integer col) {
            this.col = col;
        }

        public String getTemplateValue() {
            return templateValue;
        }

        public void setTemplateValue(String templateValue) {
            this.templateValue = templateValue;
        }

        public void addPlaceholdersEl(String placeholdersEl) {
            this.placeholdersEl.add(placeholdersEl);
        }

        public List<String> getPlaceholdersEl() {
            return placeholdersEl;
        }

        public void setPlaceholdersEl(List<String> placeholdersEl) {
            this.placeholdersEl = placeholdersEl;
        }
    }
}