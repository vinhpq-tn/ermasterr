package org.insightech.er.editor.model.dbexport.excel.sheet_generator;

import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.insightech.er.editor.model.ERDiagram;
import org.insightech.er.editor.model.ObjectModel;
import org.insightech.er.editor.model.dbexport.excel.ExportToExcelManager.LoopDefinition;
import org.insightech.er.editor.model.diagram_contents.element.node.table.ERTable;
import org.insightech.er.editor.model.diagram_contents.element.node.table.column.ColumnSet;
import org.insightech.er.editor.model.diagram_contents.element.node.table.column.NormalColumn;
import org.insightech.er.editor.model.progress_monitor.ProgressMonitor;
import org.insightech.er.util.POIUtils;
import org.insightech.er.util.POIUtils.CellLocation;

public class ColumnSheetGenerator extends AbstractSheetGenerator {

    private static final String KEYWORD_SHEET_NAME = "$SHTN";

    private ColumnTemplate columnTemplate;

    private void clear() {
        columnTemplate = null;
    }

    public void setAllColumnsData(final ProgressMonitor monitor, final HSSFWorkbook workbook, final HSSFSheet sheet, final ERDiagram diagram) throws InterruptedException {
        clear();

        final CellLocation cellLocation = POIUtils.findCell(sheet, FIND_KEYWORDS_OF_COLUMN);

        if (cellLocation != null) {
            int rowNum = cellLocation.r;
            final HSSFRow templateRow = sheet.getRow(rowNum);

            if (columnTemplate == null) {
                columnTemplate = loadColumnTemplate(workbook, sheet, cellLocation);
            }

            int order = 1;

            for (final ERTable table : diagram.getDiagramContents().getContents().getTableSet()) {

                if (diagram.getCurrentCategory() != null && !diagram.getCurrentCategory().contains(table)) {
                    continue;
                }

                monitor.subTaskWithCounter(sheet.getSheetName() + " - " + table.getName());

                for (final NormalColumn normalColumn : table.getExpandedColumns()) {

                    final HSSFRow row = POIUtils.insertRow(sheet, rowNum++);
                    setColumnData(keywordsValueMap, columnTemplate, row, normalColumn, table, order);
                    order++;
                }

                monitor.worked(1);
            }

            setCellStyle(columnTemplate, sheet, cellLocation.r, rowNum - cellLocation.r, templateRow.getFirstCellNum());
        }
    }

    public String getSheetName() {
        String name = keywordsValueMap.get(KEYWORD_SHEET_NAME);

        if (name == null) {
            name = "all attributes";
        }

        return name;
    }

    @Override
    public void generate(final ProgressMonitor monitor, final HSSFWorkbook workbook, final int sheetNo, final boolean useLogicalNameAsSheetName, final Map<String, Integer> sheetNameMap, final Map<String, ObjectModel> sheetObjectMap, final ERDiagram diagram, final Map<String, LoopDefinition> loopDefinitionMap) throws InterruptedException {
        final String name = getSheetName();
        final HSSFSheet newSheet = createNewSheet(workbook, sheetNo, name, sheetNameMap);

        final String sheetName = workbook.getSheetName(workbook.getSheetIndex(newSheet));

        sheetObjectMap.put(sheetName, new ColumnSet());

        setAllColumnsData(monitor, workbook, newSheet, diagram);
    }

    @Override
    public String getTemplateSheetName() {
        return "column_template";
    }

    @Override
    public int getKeywordsColumnNo() {
        return 20;
    }

    @Override
    public String[] getKeywords() {
        return new String[] {KEYWORD_LOGICAL_TABLE_NAME, KEYWORD_PHYSICAL_TABLE_NAME, KEYWORD_TABLE_DESCRIPTION, KEYWORD_ORDER, KEYWORD_LOGICAL_COLUMN_NAME, KEYWORD_PHYSICAL_COLUMN_NAME, KEYWORD_TYPE, KEYWORD_LENGTH, KEYWORD_DECIMAL, KEYWORD_PRIMARY_KEY, KEYWORD_NOT_NULL, KEYWORD_UNIQUE_KEY, KEYWORD_FOREIGN_KEY, KEYWORD_LOGICAL_REFERENCE_TABLE_KEY, KEYWORD_PHYSICAL_REFERENCE_TABLE_KEY, KEYWORD_LOGICAL_REFERENCE_TABLE, KEYWORD_PHYSICAL_REFERENCE_TABLE, KEYWORD_LOGICAL_REFERENCE_KEY, KEYWORD_PHYSICAL_REFERENCE_KEY, KEYWORD_AUTO_INCREMENT, KEYWORD_DEFAULT_VALUE, KEYWORD_DESCRIPTION, KEYWORD_SHEET_NAME};
    }

    @Override
    public int count(final ERDiagram diagram) {
        int count = 0;

        for (final ERTable table : diagram.getDiagramContents().getContents().getTableSet()) {

            if (diagram.getCurrentCategory() != null && !diagram.getCurrentCategory().contains(table)) {
                continue;
            }

            count++;
        }

        return count;
    }

}
