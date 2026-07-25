package org.insightech.er.editor.model.dbexport.excel.sheet_generator;

import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.insightech.er.editor.model.ERDiagram;
import org.insightech.er.editor.model.ObjectModel;
import org.insightech.er.editor.model.dbexport.excel.ExportToExcelManager.LoopDefinition;
import org.insightech.er.editor.model.diagram_contents.element.node.table.ERTable;
import org.insightech.er.editor.model.diagram_contents.element.node.table.column.NormalColumn;
import org.insightech.er.editor.model.diagram_contents.element.node.table.index.Index;
import org.insightech.er.editor.model.progress_monitor.ProgressMonitor;
import org.insightech.er.util.POIUtils;
import org.insightech.er.util.POIUtils.CellLocation;

public class IndexSheetGenerator extends AbstractSheetGenerator {

    private static final String KEYWORD_PHYSICAL_INDEX_NAME = "$PIN";

    private static final String KEYWORD_INDEX_TYPE = "$ITYP";

    private static final String KEYWORD_UNIQUE_INDEX = "$IU";

    private static final String KEYWORD_INDEX_DESCRIPTION = "$IDSC";

    private ColumnTemplate columnTemplate;

    protected void clear() {
        columnTemplate = null;
    }

    @Override
    public void generate(final ProgressMonitor monitor, final HSSFWorkbook workbook, final int sheetNo, final boolean useLogicalNameAsSheetName, final Map<String, Integer> sheetNameMap, final Map<String, ObjectModel> sheetObjectMap, final ERDiagram diagram, final Map<String, LoopDefinition> loopDefinitionMap) throws InterruptedException {
        clear();

        for (final ERTable table : diagram.getDiagramContents().getContents().getTableSet()) {
            if (diagram.getCurrentCategory() != null && !diagram.getCurrentCategory().contains(table)) {
                continue;
            }

            for (final Index index : table.getIndexes()) {
                final String name = index.getName();

                final HSSFSheet newSheet = createNewSheet(workbook, sheetNo, name, sheetNameMap);

                final String sheetName = workbook.getSheetName(workbook.getSheetIndex(newSheet));
                monitor.subTaskWithCounter("[Index] " + sheetName);

                sheetObjectMap.put(sheetName, index);

                setIndexData(workbook, newSheet, index);
                monitor.worked(1);
            }
        }
    }

    /**
     * �C���f�b�N�X�V�[�g�Ƀf�[�^��ݒ肵�܂�.
     * 
     * @param workbook
     * @param sheet
     * @param index
     */
    public void setIndexData(final HSSFWorkbook workbook, final HSSFSheet sheet, final Index index) {
        POIUtils.replace(sheet, KEYWORD_PHYSICAL_INDEX_NAME, getValue(keywordsValueMap, KEYWORD_PHYSICAL_INDEX_NAME, index.getName()));
        POIUtils.replace(sheet, KEYWORD_INDEX_TYPE, getValue(keywordsValueMap, KEYWORD_INDEX_TYPE, index.getType()));
        POIUtils.replace(sheet, KEYWORD_UNIQUE_INDEX, getValue(keywordsValueMap, KEYWORD_UNIQUE_INDEX, !index.isNonUnique()));
        POIUtils.replace(sheet, KEYWORD_PHYSICAL_TABLE_NAME, getValue(keywordsValueMap, KEYWORD_PHYSICAL_TABLE_NAME, index.getTable().getPhysicalName()));
        POIUtils.replace(sheet, KEYWORD_LOGICAL_TABLE_NAME, getValue(keywordsValueMap, KEYWORD_LOGICAL_TABLE_NAME, index.getTable().getLogicalName()));
        POIUtils.replace(sheet, KEYWORD_INDEX_DESCRIPTION, getValue(keywordsValueMap, KEYWORD_INDEX_DESCRIPTION, index.getDescription()));

        final CellLocation cellLocation = POIUtils.findCell(sheet, FIND_KEYWORDS_OF_COLUMN);

        if (cellLocation != null) {
            int rowNum = cellLocation.r;
            final HSSFRow templateRow = sheet.getRow(rowNum);

            if (columnTemplate == null) {
                columnTemplate = loadColumnTemplate(workbook, sheet, cellLocation);
            }

            int order = 1;

            for (final NormalColumn normalColumn : index.getColumns()) {
                final HSSFRow row = POIUtils.insertRow(sheet, rowNum++);
                setColumnData(keywordsValueMap, columnTemplate, row, normalColumn, index.getTable(), order);
                order++;
            }

            setCellStyle(columnTemplate, sheet, cellLocation.r, rowNum - cellLocation.r, templateRow.getFirstCellNum());
        }
    }

    @Override
    public String getTemplateSheetName() {
        return "index_template";
    }

    @Override
    public String[] getKeywords() {
        return new String[] {KEYWORD_PHYSICAL_INDEX_NAME, KEYWORD_INDEX_TYPE, KEYWORD_UNIQUE_INDEX, KEYWORD_INDEX_DESCRIPTION, KEYWORD_ORDER, KEYWORD_LOGICAL_TABLE_NAME, KEYWORD_PHYSICAL_TABLE_NAME, KEYWORD_TABLE_DESCRIPTION, KEYWORD_LOGICAL_COLUMN_NAME, KEYWORD_PHYSICAL_COLUMN_NAME, KEYWORD_TYPE, KEYWORD_LENGTH, KEYWORD_DECIMAL, KEYWORD_PRIMARY_KEY, KEYWORD_NOT_NULL, KEYWORD_UNIQUE_KEY, KEYWORD_FOREIGN_KEY, KEYWORD_LOGICAL_REFERENCE_TABLE_KEY, KEYWORD_PHYSICAL_REFERENCE_TABLE_KEY, KEYWORD_LOGICAL_REFERENCE_TABLE, KEYWORD_PHYSICAL_REFERENCE_TABLE, KEYWORD_LOGICAL_REFERENCE_KEY, KEYWORD_PHYSICAL_REFERENCE_KEY, KEYWORD_AUTO_INCREMENT, KEYWORD_DEFAULT_VALUE, KEYWORD_DESCRIPTION};
    }

    @Override
    public int getKeywordsColumnNo() {
        return 4;
    }

    @Override
    public int count(final ERDiagram diagram) {
        int count = 0;

        for (final ERTable table : diagram.getDiagramContents().getContents().getTableSet()) {
            if (diagram.getCurrentCategory() != null && !diagram.getCurrentCategory().contains(table)) {
                continue;
            }
            count += table.getIndexes().size();
        }

        return count;
    }

}
