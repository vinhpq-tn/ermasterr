package org.insightech.er.editor.model.dbexport.excel.sheet_generator;

import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.insightech.er.editor.model.ERDiagram;
import org.insightech.er.editor.model.ObjectModel;
import org.insightech.er.editor.model.dbexport.excel.ExportToExcelManager.LoopDefinition;
import org.insightech.er.editor.model.diagram_contents.element.node.table.ERTable;
import org.insightech.er.editor.model.diagram_contents.element.node.table.column.NormalColumn;
import org.insightech.er.editor.model.diagram_contents.element.node.table.index.Index;
import org.insightech.er.editor.model.diagram_contents.element.node.table.unique_key.ComplexUniqueKey;
import org.insightech.er.editor.model.progress_monitor.ProgressMonitor;
import org.insightech.er.util.Format;
import org.insightech.er.util.POIUtils;
import org.insightech.er.util.POIUtils.CellLocation;

public class TableSheetGenerator extends AbstractSheetGenerator {

    private static final String KEYWORD_LOGICAL_INDEX_MATRIX = "$LIDX";

    private static final String KEYWORD_PHYSICAL_INDEX_MATRIX = "$PIDX";

    private static final String KEYWORD_LOGICAL_COMPLEX_UNIQUE_KEY_MATRIX = "$LCUK";

    private static final String KEYWORD_PHYSICAL_COMPLEX_UNIQUE_KEY_MATRIX = "$PCUK";

    private static final String KEYWORD_TABLE_CONSTRAINT = "$TCON";

    private static final String[] FIND_KEYWORDS_OF_FK_COLUMN = {KEYWORD_LOGICAL_FOREIGN_KEY_NAME, KEYWORD_PHYSICAL_FOREIGN_KEY_NAME};

    private ColumnTemplate columnTemplate;

    private ColumnTemplate fkColumnTemplate;

    private MatrixCellStyle physicalIndexMatrixCellStyle;

    private MatrixCellStyle logicalIndexMatrixCellStyle;

    private MatrixCellStyle physicalComplexUniqueKeyMatrixCellStyle;

    private MatrixCellStyle logicalComplexUniqueKeyMatrixCellStyle;

    protected void clear() {
        columnTemplate = null;
        fkColumnTemplate = null;
        physicalIndexMatrixCellStyle = null;
        logicalIndexMatrixCellStyle = null;
        physicalComplexUniqueKeyMatrixCellStyle = null;
        logicalComplexUniqueKeyMatrixCellStyle = null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void generate(final ProgressMonitor monitor, final HSSFWorkbook workbook, final int sheetNo, final boolean useLogicalNameAsSheetName, final Map<String, Integer> sheetNameMap, final Map<String, ObjectModel> sheetObjectMap, final ERDiagram diagram, final Map<String, LoopDefinition> loopDefinitionMap) throws InterruptedException {
        clear();

        List<ERTable> nodeSet = null;

        if (diagram.getCurrentCategory() != null) {
            nodeSet = diagram.getCurrentCategory().getTableContents();
        } else {
            nodeSet = diagram.getDiagramContents().getContents().getTableSet().getList();
        }

        for (final ERTable table : nodeSet) {
            String name = null;
            if (useLogicalNameAsSheetName) {
                name = table.getLogicalName();
            } else {
                name = table.getPhysicalName();
            }

            final HSSFSheet newSheet = createNewSheet(workbook, sheetNo, name, sheetNameMap);

            final String sheetName = workbook.getSheetName(workbook.getSheetIndex(newSheet));
            monitor.subTaskWithCounter("[Table] " + sheetName);

            sheetObjectMap.put(sheetName, table);

            setTableData(workbook, newSheet, table);

            monitor.worked(1);
        }
    }

    public void setTableData(final HSSFWorkbook workbook, final HSSFSheet sheet, final ERTable table) {
        POIUtils.replace(sheet, KEYWORD_LOGICAL_TABLE_NAME, getValue(keywordsValueMap, KEYWORD_LOGICAL_TABLE_NAME, table.getLogicalName()));

        POIUtils.replace(sheet, KEYWORD_PHYSICAL_TABLE_NAME, getValue(keywordsValueMap, KEYWORD_PHYSICAL_TABLE_NAME, table.getPhysicalName()));

        POIUtils.replace(sheet, KEYWORD_TABLE_DESCRIPTION, getValue(keywordsValueMap, KEYWORD_TABLE_DESCRIPTION, table.getDescription()));

        POIUtils.replace(sheet, KEYWORD_TABLE_CONSTRAINT, getValue(keywordsValueMap, KEYWORD_TABLE_CONSTRAINT, table.getConstraint()));

        final CellLocation cellLocation = POIUtils.findCell(sheet, FIND_KEYWORDS_OF_COLUMN);

        if (cellLocation != null) {
            int rowNum = cellLocation.r;
            final HSSFRow templateRow = sheet.getRow(rowNum);

            if (columnTemplate == null) {
                columnTemplate = loadColumnTemplate(workbook, sheet, cellLocation);
            }

            int order = 1;

            for (final NormalColumn normalColumn : table.getExpandedColumns()) {
                final HSSFRow row = POIUtils.insertRow(sheet, rowNum++);
                setColumnData(keywordsValueMap, columnTemplate, row, normalColumn, table, order);
                order++;
            }

            setCellStyle(columnTemplate, sheet, cellLocation.r, rowNum - cellLocation.r, templateRow.getFirstCellNum());
        }

        final CellLocation fkCellLocation = POIUtils.findCell(sheet, FIND_KEYWORDS_OF_FK_COLUMN);

        if (fkCellLocation != null) {
            int rowNum = fkCellLocation.r;
            final HSSFRow templateRow = sheet.getRow(rowNum);

            if (fkColumnTemplate == null) {
                fkColumnTemplate = loadColumnTemplate(workbook, sheet, fkCellLocation);
            }

            int order = 1;

            for (final NormalColumn normalColumn : table.getExpandedColumns()) {
                if (normalColumn.isForeignKey()) {
                    final HSSFRow row = POIUtils.insertRow(sheet, rowNum++);
                    setColumnData(keywordsValueMap, fkColumnTemplate, row, normalColumn, table, order);
                    order++;
                }
            }

            setCellStyle(fkColumnTemplate, sheet, fkCellLocation.r, rowNum - fkCellLocation.r, templateRow.getFirstCellNum());
        }

        this.setIndexMatrix(workbook, sheet, table);
        this.setComplexUniqueKeyMatrix(workbook, sheet, table);
    }

    private void setIndexMatrix(final HSSFWorkbook workbook, final HSSFSheet sheet, final ERTable table) {
        final CellLocation logicalIndexCellLocation = POIUtils.findCell(sheet, KEYWORD_LOGICAL_INDEX_MATRIX);

        if (logicalIndexCellLocation != null) {
            if (logicalIndexMatrixCellStyle == null) {
                logicalIndexMatrixCellStyle = this.createMatrixCellStyle(workbook, sheet, logicalIndexCellLocation);
            }
            setIndexMatrix(workbook, sheet, table, logicalIndexCellLocation, logicalIndexMatrixCellStyle, true);
        }

        final CellLocation physicalIndexCellLocation = POIUtils.findCell(sheet, KEYWORD_PHYSICAL_INDEX_MATRIX);

        if (physicalIndexCellLocation != null) {
            if (physicalIndexMatrixCellStyle == null) {
                physicalIndexMatrixCellStyle = this.createMatrixCellStyle(workbook, sheet, physicalIndexCellLocation);
            }
            setIndexMatrix(workbook, sheet, table, physicalIndexCellLocation, physicalIndexMatrixCellStyle, false);
        }
    }

    private void setComplexUniqueKeyMatrix(final HSSFWorkbook workbook, final HSSFSheet sheet, final ERTable table) {
        final CellLocation logicalCellLocation = POIUtils.findCell(sheet, KEYWORD_LOGICAL_COMPLEX_UNIQUE_KEY_MATRIX);

        if (logicalCellLocation != null) {
            if (logicalComplexUniqueKeyMatrixCellStyle == null) {
                logicalComplexUniqueKeyMatrixCellStyle = this.createMatrixCellStyle(workbook, sheet, logicalCellLocation);
            }
            setComplexUniqueKeyMatrix(workbook, sheet, table, logicalCellLocation, logicalComplexUniqueKeyMatrixCellStyle, true);
        }

        final CellLocation physicalCellLocation = POIUtils.findCell(sheet, KEYWORD_PHYSICAL_COMPLEX_UNIQUE_KEY_MATRIX);

        if (physicalCellLocation != null) {
            if (physicalComplexUniqueKeyMatrixCellStyle == null) {
                physicalComplexUniqueKeyMatrixCellStyle = this.createMatrixCellStyle(workbook, sheet, physicalCellLocation);
            }

            this.setComplexUniqueKeyMatrix(workbook, sheet, table, physicalCellLocation, physicalComplexUniqueKeyMatrixCellStyle, false);
        }
    }

    private void setIndexMatrixColor(final HSSFWorkbook workbook, final HSSFCellStyle indexStyle) {
        indexStyle.setFillForegroundColor(HSSFColor.WHITE.index);
        final HSSFFont font = workbook.getFontAt(indexStyle.getFontIndex());
        font.setColor(HSSFColor.BLACK.index);
    }

    private MatrixCellStyle createMatrixCellStyle(final HSSFWorkbook workbook, final HSSFSheet sheet, final CellLocation matrixCellLocation) {

        final int matrixRowNum = matrixCellLocation.r;
        final int matrixColumnNum = matrixCellLocation.c;

        final HSSFRow matrixHeaderTemplateRow = sheet.getRow(matrixRowNum);
        final HSSFCell matrixHeaderTemplateCell = matrixHeaderTemplateRow.getCell(matrixColumnNum);

        final MatrixCellStyle matrixCellStyle = new MatrixCellStyle();

        matrixCellStyle.headerTemplateCellStyle = matrixHeaderTemplateCell.getCellStyle();

        matrixCellStyle.style11 = this.createMatrixCellStyle(workbook, matrixCellStyle.headerTemplateCellStyle, false, true, true, false);

        matrixCellStyle.style12 = this.createMatrixCellStyle(workbook, matrixCellStyle.headerTemplateCellStyle, false, true, true, true);

        matrixCellStyle.style13 = this.createMatrixCellStyle(workbook, matrixCellStyle.headerTemplateCellStyle, false, false, true, true);

        matrixCellStyle.style21 = this.createMatrixCellStyle(workbook, matrixCellStyle.headerTemplateCellStyle, true, true, true, false);

        matrixCellStyle.style22 = this.createMatrixCellStyle(workbook, matrixCellStyle.headerTemplateCellStyle, true, true, true, true);
        setIndexMatrixColor(workbook, matrixCellStyle.style22);

        matrixCellStyle.style23 = this.createMatrixCellStyle(workbook, matrixCellStyle.headerTemplateCellStyle, true, false, true, true);
        setIndexMatrixColor(workbook, matrixCellStyle.style23);

        matrixCellStyle.style31 = this.createMatrixCellStyle(workbook, matrixCellStyle.headerTemplateCellStyle, true, true, false, false);

        matrixCellStyle.style32 = this.createMatrixCellStyle(workbook, matrixCellStyle.headerTemplateCellStyle, true, true, false, true);
        setIndexMatrixColor(workbook, matrixCellStyle.style32);

        matrixCellStyle.style33 = this.createMatrixCellStyle(workbook, matrixCellStyle.headerTemplateCellStyle, true, false, false, true);
        setIndexMatrixColor(workbook, matrixCellStyle.style33);

        return matrixCellStyle;
    }

    private HSSFCellStyle createMatrixCellStyle(final HSSFWorkbook workbook, final HSSFCellStyle matrixHeaderTemplateCellStyle, final boolean top, final boolean right, final boolean bottom, final boolean left) {
        final HSSFCellStyle cellStyle = POIUtils.copyCellStyle(workbook, matrixHeaderTemplateCellStyle);

        if (top) {
            cellStyle.setBorderTop(CellStyle.BORDER_THIN);
        }
        if (right) {
            cellStyle.setBorderRight(CellStyle.BORDER_THIN);
        }
        if (bottom) {
            cellStyle.setBorderBottom(CellStyle.BORDER_THIN);
        }
        if (left) {
            cellStyle.setBorderLeft(CellStyle.BORDER_THIN);
        }

        return cellStyle;
    }

    private void setIndexMatrix(final HSSFWorkbook workbook, final HSSFSheet sheet, final ERTable table, final CellLocation cellLocation, final MatrixCellStyle matrixCellStyle, final boolean isLogical) {

        int rowNum = cellLocation.r;
        final int columnNum = cellLocation.c;

        final HSSFRow headerTemplateRow = sheet.getRow(rowNum);
        final HSSFCell headerTemplateCell = headerTemplateRow.getCell(columnNum);

        final int num = table.getIndexes().size();

        if (num == 0) {
            headerTemplateRow.removeCell(headerTemplateCell);

            final HSSFRow row = sheet.getRow(rowNum - 1);
            if (row != null) {
                final HSSFCell cell = row.getCell(columnNum);
                if (cell != null) {
                    cell.getCellStyle().setBorderBottom(headerTemplateCell.getCellStyle().getBorderBottom());
                }
            }
            return;
        }

        final HSSFRow headerRow = sheet.createRow(rowNum++);

        for (int i = 0; i < num + 1; i++) {
            final HSSFCell cell = headerRow.createCell(columnNum + i);

            if (i == 0) {
                cell.setCellStyle(matrixCellStyle.style11);

            } else {
                final Index index = table.getIndexes().get(i - 1);
                final HSSFRichTextString text = new HSSFRichTextString(index.getName());
                cell.setCellValue(text);

                if (i != num) {
                    cell.setCellStyle(matrixCellStyle.style12);
                } else {
                    cell.setCellStyle(matrixCellStyle.style13);
                }
            }
        }

        final int columnSize = table.getExpandedColumns().size();
        for (int j = 0; j < columnSize; j++) {
            final NormalColumn normalColumn = table.getExpandedColumns().get(j);

            final HSSFRow row = POIUtils.insertRow(sheet, rowNum++);

            for (int i = 0; i < num + 1; i++) {
                final HSSFCell cell = row.createCell(columnNum + i);

                if (i == 0) {
                    String columnName = null;
                    if (isLogical) {
                        columnName = normalColumn.getLogicalName();
                    } else {
                        columnName = normalColumn.getPhysicalName();
                    }

                    final HSSFRichTextString text = new HSSFRichTextString(columnName);
                    cell.setCellValue(text);
                    cell.setCellStyle(headerTemplateCell.getCellStyle());

                    if (j != columnSize - 1) {
                        cell.setCellStyle(matrixCellStyle.style21);
                    } else {
                        cell.setCellStyle(matrixCellStyle.style31);
                    }

                } else {
                    final Index index = table.getIndexes().get(i - 1);
                    final List<NormalColumn> indexColumnList = index.getColumns();

                    final int indexNo = indexColumnList.indexOf(normalColumn);
                    if (indexNo != -1) {
                        cell.setCellValue(indexNo + 1);
                    }

                    if (i != num) {
                        if (j != columnSize - 1) {
                            cell.setCellStyle(matrixCellStyle.style22);
                        } else {
                            cell.setCellStyle(matrixCellStyle.style32);
                        }

                    } else {
                        if (j != columnSize - 1) {
                            cell.setCellStyle(matrixCellStyle.style23);
                        } else {
                            cell.setCellStyle(matrixCellStyle.style33);
                        }
                    }
                }
            }
        }
    }

    private void setComplexUniqueKeyMatrix(final HSSFWorkbook workbook, final HSSFSheet sheet, final ERTable table, final CellLocation cellLocation, final MatrixCellStyle matrixCellStyle, final boolean isLogical) {

        int rowNum = cellLocation.r;
        final int columnNum = cellLocation.c;

        final HSSFRow headerTemplateRow = sheet.getRow(rowNum);
        final HSSFCell headerTemplateCell = headerTemplateRow.getCell(columnNum);

        final int num = table.getComplexUniqueKeyList().size();

        if (num == 0) {
            headerTemplateRow.removeCell(headerTemplateCell);

            final HSSFRow row = sheet.getRow(rowNum - 1);
            if (row != null) {
                final HSSFCell cell = row.getCell(columnNum);
                if (cell != null) {
                    cell.getCellStyle().setBorderBottom(headerTemplateCell.getCellStyle().getBorderBottom());
                }
            }
            return;
        }

        final HSSFRow headerRow = sheet.createRow(rowNum++);

        for (int i = 0; i < num + 1; i++) {
            final HSSFCell cell = headerRow.createCell(columnNum + i);

            if (i == 0) {
                cell.setCellStyle(matrixCellStyle.style11);

            } else {
                final ComplexUniqueKey complexUniqueKey = table.getComplexUniqueKeyList().get(i - 1);
                final HSSFRichTextString text = new HSSFRichTextString(Format.null2blank(complexUniqueKey.getUniqueKeyName()));
                cell.setCellValue(text);

                if (i != num) {
                    cell.setCellStyle(matrixCellStyle.style12);
                } else {
                    cell.setCellStyle(matrixCellStyle.style13);
                }
            }
        }

        final int columnSize = table.getExpandedColumns().size();
        for (int j = 0; j < columnSize; j++) {
            final NormalColumn normalColumn = table.getExpandedColumns().get(j);

            final HSSFRow row = POIUtils.insertRow(sheet, rowNum++);

            for (int i = 0; i < num + 1; i++) {
                final HSSFCell cell = row.createCell(columnNum + i);

                if (i == 0) {
                    String columnName = null;
                    if (isLogical) {
                        columnName = normalColumn.getLogicalName();
                    } else {
                        columnName = normalColumn.getPhysicalName();
                    }

                    final HSSFRichTextString text = new HSSFRichTextString(columnName);
                    cell.setCellValue(text);
                    cell.setCellStyle(headerTemplateCell.getCellStyle());

                    if (j != columnSize - 1) {
                        cell.setCellStyle(matrixCellStyle.style21);
                    } else {
                        cell.setCellStyle(matrixCellStyle.style31);
                    }

                } else {
                    final ComplexUniqueKey complexUniqueKey = table.getComplexUniqueKeyList().get(i - 1);
                    final List<NormalColumn> targetColumnList = complexUniqueKey.getColumnList();

                    final int indexNo = targetColumnList.indexOf(normalColumn);
                    if (indexNo != -1) {
                        cell.setCellValue(indexNo + 1);
                    }

                    if (i != num) {
                        if (j != columnSize - 1) {
                            cell.setCellStyle(matrixCellStyle.style22);
                        } else {
                            cell.setCellStyle(matrixCellStyle.style32);
                        }

                    } else {
                        if (j != columnSize - 1) {
                            cell.setCellStyle(matrixCellStyle.style23);
                        } else {
                            cell.setCellStyle(matrixCellStyle.style33);
                        }
                    }
                }
            }
        }
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getTemplateSheetName() {
        return "table_template";
    }

    @Override
    public String[] getKeywords() {
        return new String[] {KEYWORD_LOGICAL_TABLE_NAME, KEYWORD_PHYSICAL_TABLE_NAME, KEYWORD_TABLE_DESCRIPTION, KEYWORD_TABLE_CONSTRAINT, KEYWORD_ORDER, KEYWORD_LOGICAL_COLUMN_NAME, KEYWORD_PHYSICAL_COLUMN_NAME, KEYWORD_TYPE, KEYWORD_LENGTH, KEYWORD_DECIMAL, KEYWORD_PRIMARY_KEY, KEYWORD_NOT_NULL, KEYWORD_UNIQUE_KEY, KEYWORD_FOREIGN_KEY, KEYWORD_LOGICAL_REFERENCE_TABLE_KEY, KEYWORD_PHYSICAL_REFERENCE_TABLE_KEY, KEYWORD_LOGICAL_REFERENCE_TABLE, KEYWORD_PHYSICAL_REFERENCE_TABLE, KEYWORD_LOGICAL_REFERENCE_KEY, KEYWORD_PHYSICAL_REFERENCE_KEY, KEYWORD_AUTO_INCREMENT, KEYWORD_DEFAULT_VALUE, KEYWORD_DESCRIPTION, KEYWORD_LOGICAL_INDEX_MATRIX, KEYWORD_PHYSICAL_INDEX_MATRIX, KEYWORD_LOGICAL_FOREIGN_KEY_NAME, KEYWORD_PHYSICAL_FOREIGN_KEY_NAME};
    }

    @Override
    public int getKeywordsColumnNo() {
        return 0;
    }

    @Override
    public int count(final ERDiagram diagram) {
        return diagram.getDiagramContents().getContents().getTableSet().getList().size();
    }
}
