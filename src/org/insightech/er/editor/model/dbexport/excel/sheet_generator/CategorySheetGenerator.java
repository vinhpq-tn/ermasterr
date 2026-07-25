package org.insightech.er.editor.model.dbexport.excel.sheet_generator;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.insightech.er.editor.model.ERDiagram;
import org.insightech.er.editor.model.ObjectModel;
import org.insightech.er.editor.model.dbexport.excel.ExportToExcelManager.LoopDefinition;
import org.insightech.er.editor.model.diagram_contents.element.node.category.Category;
import org.insightech.er.editor.model.diagram_contents.element.node.table.ERTable;
import org.insightech.er.editor.model.progress_monitor.ProgressMonitor;
import org.insightech.er.util.POIUtils;

public class CategorySheetGenerator extends TableSheetGenerator {

    @Override
    public void generate(final ProgressMonitor monitor, final HSSFWorkbook workbook, final int sheetNo, final boolean useLogicalNameAsSheetName, final Map<String, Integer> sheetNameMap, final Map<String, ObjectModel> sheetObjectMap, final ERDiagram diagram, final Map<String, LoopDefinition> loopDefinitionMap) throws InterruptedException {
        clear();

        if (diagram.getCurrentCategory() != null) {
            return;
        }

        final LoopDefinition loopDefinition = loopDefinitionMap.get(getTemplateSheetName());
        final HSSFSheet oldSheet = workbook.getSheetAt(sheetNo);

        final List<ERTable> allTables = new ArrayList<ERTable>(diagram.getDiagramContents().getContents().getTableSet().getList());

        for (final Category category : diagram.getDiagramContents().getSettings().getCategorySetting().getSelectedCategories()) {
            final HSSFSheet newSheet = createNewSheet(workbook, sheetNo, category.getName(), sheetNameMap);

            final String sheetName = workbook.getSheetName(workbook.getSheetIndex(newSheet));
            monitor.subTaskWithCounter("[Category] " + sheetName);

            sheetObjectMap.put(sheetName, category);

            boolean first = true;

            for (final ERTable table : category.getTableContents()) {
                if (allTables.contains(table)) {
                    allTables.remove(table);
                    monitor.worked(1);
                }

                if (first) {
                    first = false;

                } else {
                    POIUtils.copyRow(oldSheet, newSheet, loopDefinition.startLine - 1, oldSheet.getLastRowNum(), newSheet.getLastRowNum() + loopDefinition.spaceLine + 1);
                }

                setTableData(workbook, newSheet, table);

                newSheet.setRowBreak(newSheet.getLastRowNum() + loopDefinition.spaceLine);
            }

            if (first) {
                int rowIndex = loopDefinition.startLine - 1;

                while (rowIndex <= newSheet.getLastRowNum()) {
                    final HSSFRow row = newSheet.getRow(rowIndex);
                    if (row != null) {
                        newSheet.removeRow(row);
                    }

                    rowIndex++;
                }
            }

            monitor.worked(1);
        }

        if (!allTables.isEmpty()) {
            final HSSFSheet newSheet = createNewSheet(workbook, sheetNo, loopDefinition.sheetName, sheetNameMap);

            final String sheetName = workbook.getSheetName(workbook.getSheetIndex(newSheet));

            sheetObjectMap.put(sheetName, diagram.getDiagramContents().getContents().getTableSet());

            boolean first = true;

            for (final ERTable table : allTables) {
                monitor.subTaskWithCounter("[Category] " + newSheet.getSheetName() + " - " + table.getName());

                if (first) {
                    first = false;

                } else {
                    POIUtils.copyRow(oldSheet, newSheet, loopDefinition.startLine - 1, oldSheet.getLastRowNum(), newSheet.getLastRowNum() + loopDefinition.spaceLine + 1);
                }

                setTableData(workbook, newSheet, table);
                newSheet.setRowBreak(newSheet.getLastRowNum() + loopDefinition.spaceLine);

                monitor.worked(1);
            }
        }
    }

    @Override
    public String getTemplateSheetName() {
        return "category_template";
    }

    @Override
    public int count(final ERDiagram diagram) {
        int count = diagram.getDiagramContents().getSettings().getCategorySetting().getSelectedCategories().size();

        count += diagram.getDiagramContents().getContents().getTableSet().getList().size();

        return count;
    }

}
