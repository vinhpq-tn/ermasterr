package org.insightech.er.editor.model.dbexport.excel.sheet_generator;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.insightech.er.editor.model.ERDiagram;
import org.insightech.er.editor.model.ObjectModel;
import org.insightech.er.editor.model.dbexport.excel.ExportToExcelManager.LoopDefinition;
import org.insightech.er.editor.model.diagram_contents.element.connection.Relation;
import org.insightech.er.editor.model.diagram_contents.element.node.table.TableView;
import org.insightech.er.editor.model.diagram_contents.element.node.table.column.NormalColumn;
import org.insightech.er.editor.model.progress_monitor.ProgressMonitor;
import org.insightech.er.util.Format;
import org.insightech.er.util.POIUtils;
import org.insightech.er.util.POIUtils.CellLocation;

public abstract class AbstractSheetGenerator {

    private static final int MAX_SHEET_NAME_LENGTH = 26;

    protected static final String KEYWORD_ORDER = "$ORD";

    protected static final String KEYWORD_LOGICAL_TABLE_NAME = "$LTN";

    protected static final String KEYWORD_PHYSICAL_TABLE_NAME = "$PTN";

    protected static final String KEYWORD_LOGICAL_COLUMN_NAME = "$LCN";

    protected static final String KEYWORD_PHYSICAL_COLUMN_NAME = "$PCN";

    protected static final String KEYWORD_TYPE = "$TYP";

    protected static final String KEYWORD_TYPE_EMBEDDED = "$TYE";

    protected static final String KEYWORD_LENGTH = "$LEN";

    protected static final String KEYWORD_DECIMAL = "$DEC";

    protected static final String KEYWORD_PRIMARY_KEY = "$PK";

    protected static final String KEYWORD_NOT_NULL = "$NN";

    protected static final String KEYWORD_UNIQUE_KEY = "$UK";

    protected static final String KEYWORD_FOREIGN_KEY = "$FK";

    protected static final String KEYWORD_LOGICAL_REFERENCE_TABLE_KEY = "$LRFTC";

    protected static final String KEYWORD_PHYSICAL_REFERENCE_TABLE_KEY = "$PRFTC";

    protected static final String KEYWORD_LOGICAL_REFERENCE_TABLE = "$LRFT";

    protected static final String KEYWORD_PHYSICAL_REFERENCE_TABLE = "$PRFT";

    protected static final String KEYWORD_LOGICAL_REFERENCE_KEY = "$LRFC";

    protected static final String KEYWORD_PHYSICAL_REFERENCE_KEY = "$PRFC";

    protected static final String KEYWORD_AUTO_INCREMENT = "$INC";

    protected static final String KEYWORD_DESCRIPTION = "$CDSC";

    protected static final String KEYWORD_DEFAULT_VALUE = "$DEF";

    protected static final String KEYWORD_LOGICAL_FOREIGN_KEY_NAME = "$LFKN";

    protected static final String KEYWORD_PHYSICAL_FOREIGN_KEY_NAME = "$PFKN";

    protected static final String KEYWORD_TABLE_DESCRIPTION = "$TDSC";

    private static final String[] KEYWORDS_OF_COLUMN = {KEYWORD_ORDER, KEYWORD_LOGICAL_TABLE_NAME, KEYWORD_PHYSICAL_TABLE_NAME, KEYWORD_LOGICAL_COLUMN_NAME, KEYWORD_PHYSICAL_COLUMN_NAME, KEYWORD_TYPE, KEYWORD_TYPE_EMBEDDED, KEYWORD_LENGTH, KEYWORD_DECIMAL, KEYWORD_PRIMARY_KEY, KEYWORD_NOT_NULL, KEYWORD_UNIQUE_KEY, KEYWORD_FOREIGN_KEY, KEYWORD_LOGICAL_REFERENCE_TABLE_KEY, KEYWORD_PHYSICAL_REFERENCE_TABLE_KEY, KEYWORD_LOGICAL_REFERENCE_TABLE, KEYWORD_PHYSICAL_REFERENCE_TABLE, KEYWORD_LOGICAL_REFERENCE_KEY, KEYWORD_PHYSICAL_REFERENCE_KEY, KEYWORD_AUTO_INCREMENT, KEYWORD_DEFAULT_VALUE, KEYWORD_DESCRIPTION, KEYWORD_LOGICAL_FOREIGN_KEY_NAME, KEYWORD_PHYSICAL_FOREIGN_KEY_NAME};

    protected static final String[] FIND_KEYWORDS_OF_COLUMN = {KEYWORD_LOGICAL_COLUMN_NAME, KEYWORD_PHYSICAL_COLUMN_NAME};

    protected Map<String, String> keywordsValueMap;

    public static class ColumnTemplate {
        public Map<Integer, String> columnTemplateMap = new HashMap<Integer, String>();

        public List<HSSFCellStyle> topRowCellStyleList;

        public List<HSSFCellStyle> middleRowCellStyleList;

        public List<HSSFCellStyle> bottomRowCellStyleList;
    }

    public static class MatrixCellStyle {
        public HSSFCellStyle headerTemplateCellStyle;

        public HSSFCellStyle style11;
        public HSSFCellStyle style12;
        public HSSFCellStyle style13;
        public HSSFCellStyle style21;
        public HSSFCellStyle style22;
        public HSSFCellStyle style23;
        public HSSFCellStyle style31;
        public HSSFCellStyle style32;
        public HSSFCellStyle style33;
    }

    protected Map<String, String> buildKeywordsValueMap(final HSSFSheet wordsSheet, final int columnNo, final String[] keywords) {
        final Map<String, String> keywordsValueMap = new HashMap<String, String>();

        for (final String keyword : keywords) {
            final CellLocation location = POIUtils.findCell(wordsSheet, keyword, columnNo);
            if (location != null) {
                final HSSFRow row = wordsSheet.getRow(location.r);

                final HSSFCell cell = row.getCell(location.c + 2);
                final String value = cell.getRichStringCellValue().getString();

                if (value != null) {
                    keywordsValueMap.put(keyword, value);
                }
            }
        }

        return keywordsValueMap;
    }

    protected String getValue(final Map<String, String> keywordsValueMap, final String keyword, final Object obj) {
        if (obj instanceof Boolean) {
            if (Boolean.TRUE.equals(obj)) {
                final String value = keywordsValueMap.get(keyword);

                if (value != null && !"".equals(value)) {
                    return value;
                }
            } else {
                return "";
            }
        }

        if (obj == null) {
            return "";
        }

        return obj.toString();
    }

    protected void setColumnData(final Map<String, String> keywordsValueMap, final ColumnTemplate columnTemplate, final HSSFRow row, final NormalColumn normalColumn, final TableView tableView, final int order) {

        for (final int columnNum : columnTemplate.columnTemplateMap.keySet()) {
            final HSSFCell cell = row.createCell(columnNum);
            final String template = columnTemplate.columnTemplateMap.get(columnNum);

            String value = null;
            if (KEYWORD_ORDER.equals(template)) {
                value = String.valueOf(order);

            } else {
                value = getColumnValue(keywordsValueMap, normalColumn, tableView, template);
            }

            try {
                final double num = Double.parseDouble(value);
                cell.setCellValue(num);

            } catch (final NumberFormatException e) {
                final HSSFRichTextString text = new HSSFRichTextString(value);
                cell.setCellValue(text);
            }
        }
    }

    private String getColumnValue(final Map<String, String> keywordsValueMap, final NormalColumn normalColumn, final TableView tableView, final String template) {
        String str = template;

        for (final String keyword : KEYWORDS_OF_COLUMN) {
            str = str.replaceAll("\\" + keyword, Matcher.quoteReplacement(getKeywordValue(keywordsValueMap, normalColumn, tableView, keyword)));
        }

        return str;
    }

    private String getKeywordValue(final Map<String, String> keywordsValueMap, final NormalColumn normalColumn, final TableView tableView, final String keyword) {
        Object obj = null;

        if (KEYWORD_LOGICAL_TABLE_NAME.equals(keyword)) {
            obj = tableView.getLogicalName();

        } else if (KEYWORD_PHYSICAL_TABLE_NAME.equals(keyword)) {
            obj = tableView.getPhysicalName();

        } else if (KEYWORD_LOGICAL_COLUMN_NAME.equals(keyword)) {
            obj = normalColumn.getLogicalName();

        } else if (KEYWORD_PHYSICAL_COLUMN_NAME.equals(keyword)) {
            obj = normalColumn.getPhysicalName();

        } else if (KEYWORD_TYPE.equals(keyword)) {
            if (normalColumn.getType() == null) {
                obj = null;
            } else {
                obj = Format.formatType(normalColumn.getType(), normalColumn.getTypeData(), tableView.getDiagram().getDatabase(), false);
            }
        } else if (KEYWORD_TYPE_EMBEDDED.equals(keyword)) {
            if (normalColumn.getType() == null) {
                obj = null;
            } else {
                obj = Format.formatType(normalColumn.getType(), normalColumn.getTypeData(), tableView.getDiagram().getDatabase(), true);
            }
        } else if (KEYWORD_LENGTH.equals(keyword)) {
            obj = normalColumn.getTypeData().getLength();

        } else if (KEYWORD_DECIMAL.equals(keyword)) {
            obj = normalColumn.getTypeData().getDecimal();

        } else if (KEYWORD_PRIMARY_KEY.equals(keyword)) {
            obj = normalColumn.isPrimaryKey();

        } else if (KEYWORD_NOT_NULL.equals(keyword)) {
            obj = normalColumn.isNotNull();

        } else if (KEYWORD_FOREIGN_KEY.equals(keyword)) {
            final List<Relation> relationList = normalColumn.getRelationList();

            if (relationList == null || relationList.isEmpty()) {
                obj = false;

            } else {
                obj = true;
            }

        } else if (KEYWORD_LOGICAL_REFERENCE_TABLE_KEY.equals(keyword)) {
            final List<Relation> relationList = normalColumn.getRelationList();

            if (relationList == null || relationList.isEmpty()) {
                obj = null;

            } else {
                final Relation relation = relationList.get(0);

                final TableView referencedTable = relation.getSourceTableView();
                obj = referencedTable.getLogicalName() + "." + normalColumn.getReferencedColumn(relation).getLogicalName();
            }

        } else if (KEYWORD_PHYSICAL_REFERENCE_TABLE_KEY.equals(keyword)) {
            final List<Relation> relationList = normalColumn.getRelationList();

            if (relationList == null || relationList.isEmpty()) {
                obj = null;

            } else {
                final Relation relation = relationList.get(0);

                final TableView referencedTable = relation.getSourceTableView();
                obj = referencedTable.getPhysicalName() + "." + normalColumn.getReferencedColumn(relation).getPhysicalName();
            }

        } else if (KEYWORD_LOGICAL_REFERENCE_TABLE.equals(keyword)) {
            final List<Relation> relationList = normalColumn.getRelationList();

            if (relationList == null || relationList.isEmpty()) {
                obj = null;

            } else {
                final TableView referencedTable = relationList.get(0).getSourceTableView();
                obj = referencedTable.getLogicalName();
            }

        } else if (KEYWORD_PHYSICAL_REFERENCE_TABLE.equals(keyword)) {
            final List<Relation> relationList = normalColumn.getRelationList();

            if (relationList == null || relationList.isEmpty()) {
                obj = null;

            } else {
                final TableView referencedTable = relationList.get(0).getSourceTableView();
                obj = referencedTable.getPhysicalName();
            }

        } else if (KEYWORD_LOGICAL_REFERENCE_KEY.equals(keyword)) {
            final List<Relation> relationList = normalColumn.getRelationList();

            if (relationList == null || relationList.isEmpty()) {
                obj = null;

            } else {
                final Relation relation = relationList.get(0);

                obj = normalColumn.getReferencedColumn(relation).getLogicalName();
            }

        } else if (KEYWORD_PHYSICAL_REFERENCE_KEY.equals(keyword)) {
            final List<Relation> relationList = normalColumn.getRelationList();

            if (relationList == null || relationList.isEmpty()) {
                obj = null;

            } else {
                final Relation relation = relationList.get(0);

                obj = normalColumn.getReferencedColumn(relation).getPhysicalName();
            }

        } else if (KEYWORD_LOGICAL_FOREIGN_KEY_NAME.equals(keyword)) {
            obj = normalColumn.getLogicalName();

        } else if (KEYWORD_PHYSICAL_FOREIGN_KEY_NAME.equals(keyword)) {
            obj = normalColumn.getPhysicalName();

        } else if (KEYWORD_UNIQUE_KEY.equals(keyword)) {
            obj = normalColumn.isUniqueKey();

        } else if (KEYWORD_DESCRIPTION.equals(keyword)) {
            obj = normalColumn.getDescription();

        } else if (KEYWORD_DEFAULT_VALUE.equals(keyword)) {
            obj = normalColumn.getDefaultValue();

        } else if (KEYWORD_AUTO_INCREMENT.equals(keyword)) {
            obj = normalColumn.isAutoIncrement();

        }

        return getValue(keywordsValueMap, keyword, obj);
    }

    protected ColumnTemplate loadColumnTemplate(final HSSFWorkbook workbook, final HSSFSheet templateSheet, final CellLocation location) {
        if (location == null) {
            return null;
        }

        final ColumnTemplate columnTemplate = new ColumnTemplate();

        final HSSFRow row = templateSheet.getRow(location.r);
        final HSSFRow bottomRow = templateSheet.getRow(location.r + 1);

        for (int colNum = row.getFirstCellNum(); colNum <= row.getLastCellNum(); colNum++) {

            final HSSFCell cell = row.getCell(colNum);

            if (cell != null) {
                columnTemplate.columnTemplateMap.put(colNum, cell.getRichStringCellValue().getString());
            }
        }

        columnTemplate.topRowCellStyleList = POIUtils.copyCellStyle(workbook, row);
        columnTemplate.middleRowCellStyleList = POIUtils.copyCellStyle(workbook, row);
        columnTemplate.bottomRowCellStyleList = POIUtils.copyCellStyle(workbook, row);

        for (short i = 0; i < columnTemplate.middleRowCellStyleList.size(); i++) {
            final HSSFCellStyle middleRowCellStyle = columnTemplate.middleRowCellStyleList.get(i);
            if (middleRowCellStyle != null) {
                final HSSFCellStyle topRowCellStyle = columnTemplate.topRowCellStyleList.get(i);
                final HSSFCellStyle bottomRowCellStyle = columnTemplate.bottomRowCellStyleList.get(i);

                final HSSFCell bottomCell = bottomRow.getCell(row.getFirstCellNum() + i);

                topRowCellStyle.setBorderBottom(bottomCell.getCellStyle().getBorderTop());
                middleRowCellStyle.setBorderTop(bottomCell.getCellStyle().getBorderTop());
                middleRowCellStyle.setBorderBottom(bottomCell.getCellStyle().getBorderTop());
                bottomRowCellStyle.setBorderTop(bottomCell.getCellStyle().getBorderTop());
                bottomRowCellStyle.setBorderBottom(bottomCell.getCellStyle().getBorderBottom());
            }
        }

        return columnTemplate;
    }

    protected void setCellStyle(final ColumnTemplate columnTemplate, final HSSFSheet sheet, final int firstRowNum, final int rowSize, final int firstColNum) {

        sheet.removeRow(sheet.getRow(firstRowNum + rowSize));

        final HSSFRow bottomRowTemplate = sheet.getRow(firstRowNum + rowSize + 1);
        sheet.removeRow(bottomRowTemplate);

        for (int r = firstRowNum + 1; r < firstRowNum + rowSize; r++) {
            final HSSFRow row = sheet.getRow(r);

            for (int i = 0; i < columnTemplate.middleRowCellStyleList.size(); i++) {
                final HSSFCell cell = row.getCell(firstColNum + i);
                if (cell != null && columnTemplate.middleRowCellStyleList.get(i) != null) {
                    cell.setCellStyle(columnTemplate.middleRowCellStyleList.get(i));
                }
            }
        }

        if (rowSize > 0) {
            final HSSFRow topRow = sheet.getRow(firstRowNum);

            for (int i = 0; i < columnTemplate.topRowCellStyleList.size(); i++) {
                final HSSFCell cell = topRow.getCell(firstColNum + i);
                if (cell != null) {
                    if (columnTemplate.topRowCellStyleList.get(i) != null) {
                        cell.setCellStyle(columnTemplate.topRowCellStyleList.get(i));
                    }
                }
            }

            final HSSFRow bottomRow = sheet.getRow(firstRowNum + rowSize - 1);

            for (int i = 0; i < columnTemplate.bottomRowCellStyleList.size(); i++) {
                final HSSFCell bottomRowCell = bottomRow.getCell(firstColNum + i);
                if (bottomRowCell != null) {
                    if (columnTemplate.bottomRowCellStyleList.get(i) != null) {
                        bottomRowCell.setCellStyle(columnTemplate.bottomRowCellStyleList.get(i));
                    }
                }
            }

        } else {
            final HSSFRow bottomRow = sheet.getRow(firstRowNum - 1);

            if (bottomRow != null) {
                for (int i = 0; i < columnTemplate.bottomRowCellStyleList.size(); i++) {
                    final HSSFCell bottomRowCell = bottomRow.getCell(firstColNum + i);

                    if (bottomRowCell != null) {
                        final HSSFCellStyle bottomRowCellStyle = bottomRowCell.getCellStyle();
                        if (columnTemplate.bottomRowCellStyleList.get(i) != null) {
                            bottomRowCellStyle.setBorderBottom(columnTemplate.bottomRowCellStyleList.get(i).getBorderBottom());
                        }
                    }
                }
            }
        }

        final List<CellRangeAddress> regionList = POIUtils.getMergedRegionList(sheet, firstRowNum);

        for (int r = firstRowNum + 1; r < firstRowNum + rowSize; r++) {
            POIUtils.copyMergedRegion(sheet, regionList, r);
        }
    }

    public static HSSFSheet createNewSheet(final HSSFWorkbook workbook, final int sheetNo, final String name, final Map<String, Integer> sheetNameMap) {
        final HSSFSheet sheet = workbook.cloneSheet(sheetNo);
        final int newSheetNo = workbook.getSheetIndex(sheet);

        workbook.setSheetName(newSheetNo, decideSheetName(name, sheetNameMap));

        return sheet;
    }

    public static String decideSheetName(String name, final Map<String, Integer> sheetNameMap) {
        if (name.length() > MAX_SHEET_NAME_LENGTH) {
            name = name.substring(0, MAX_SHEET_NAME_LENGTH);
        }

        String sheetName = null;

        Integer sameNameNum = sheetNameMap.get(name.toUpperCase());
        if (sameNameNum == null) {
            sameNameNum = 0;
            sheetName = name;

        } else {
            do {
                sameNameNum++;
                sheetName = name + "(" + sameNameNum + ")";
            } while (sheetNameMap.containsKey(sheetName.toUpperCase()));
        }

        sheetNameMap.put(name.toUpperCase(), sameNameNum);

        return sheetName;
    }

    public void init(final HSSFSheet wordsSheet) {
        keywordsValueMap = buildKeywordsValueMap(wordsSheet, getKeywordsColumnNo(), getKeywords());
    }

    public abstract void generate(ProgressMonitor monitor, HSSFWorkbook workbook, int sheetNo, boolean useLogicalNameAsSheetName, Map<String, Integer> sheetNameMap, Map<String, ObjectModel> sheetObjectMap, ERDiagram diagram, Map<String, LoopDefinition> loopDefinitionMap) throws InterruptedException;

    public abstract int count(ERDiagram diagram);

    public abstract String getTemplateSheetName();

    public abstract int getKeywordsColumnNo();

    public abstract String[] getKeywords();
}
