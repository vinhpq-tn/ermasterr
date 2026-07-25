package org.insightech.er.editor.model.dbexport.testdata.impl;

import java.io.File;
import java.io.IOException;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.insightech.er.editor.model.ERDiagram;
import org.insightech.er.editor.model.dbexport.testdata.TestDataCreator;
import org.insightech.er.editor.model.diagram_contents.element.node.table.ERTable;
import org.insightech.er.editor.model.diagram_contents.element.node.table.column.NormalColumn;
import org.insightech.er.editor.model.testdata.RepeatTestData;
import org.insightech.er.editor.model.testdata.RepeatTestDataDef;
import org.insightech.er.util.Format;
import org.insightech.er.util.POIUtils;
import org.insightech.er.util.io.FileUtils;

public class DBUnitXLSTestDataCreator extends TestDataCreator {

    private HSSFWorkbook workbook;

    private Set<String> sheetNames;

    private HSSFSheet sheet;

    private int rowNum = 0;

    public DBUnitXLSTestDataCreator() {}

    @Override
    protected void openFile() throws IOException {
        workbook = new HSSFWorkbook();
        sheetNames = new HashSet<String>();
    }

    @Override
    protected void write() throws Exception {
        super.write();

        final File file = new File(FileUtils.getFile(baseDir, exportTestDataSetting.getExportFilePath()), testData.getName() + ".xls");

        file.getParentFile().mkdirs();

        POIUtils.writeExcelFile(file, workbook);
    }

    @Override
    protected void closeFile() throws IOException {}

    @Override
    protected boolean skipTable(final ERTable table) {
        final String sheetName = table.getPhysicalName();

        if (sheetNames.contains(sheetName)) {
            return true;
        }

        sheetNames.add(sheetName);

        return false;
    }

    @Override
    protected void writeTableHeader(final ERDiagram diagram, final ERTable table) {
        final String sheetName = table.getPhysicalName();
        sheet = workbook.createSheet(sheetName);

        rowNum = 0;
        final HSSFRow row = sheet.createRow(rowNum++);

        int col = 0;

        for (final NormalColumn column : table.getExpandedColumns()) {
            final HSSFCell cell = row.createCell(col++);
            cell.setCellValue(new HSSFRichTextString(column.getPhysicalName()));
        }
    }

    @Override
    protected void writeTableFooter(final ERTable table) {}

    @Override
    protected void writeDirectTestData(final ERTable table, final Map<NormalColumn, String> data, final String database) {
        final HSSFRow row = sheet.createRow(rowNum++);

        int col = 0;

        for (final NormalColumn column : table.getExpandedColumns()) {
            final HSSFCell cell = row.createCell(col++);

            final String value = Format.null2blank(data.get(column));

            if (value == null || "null".equals(value.toLowerCase())) {

            } else {
                cell.setCellValue(new HSSFRichTextString(value));
            }
        }
    }

    @Override
    protected void writeRepeatTestData(final ERTable table, final RepeatTestData repeatTestData, final String database) {

        for (int i = 0; i < repeatTestData.getTestDataNum(); i++) {
            final HSSFRow row = sheet.createRow(rowNum++);

            int col = 0;

            for (final NormalColumn column : table.getExpandedColumns()) {
                final HSSFCell cell = row.createCell(col++);

                final RepeatTestDataDef repeatTestDataDef = repeatTestData.getDataDef(column);

                final String value = getMergedRepeatTestDataValue(i, repeatTestDataDef, column);

                if (value == null || "null".equals(value.toLowerCase())) {

                } else {
                    cell.setCellValue(new HSSFRichTextString(value));
                }
            }
        }

    }

}
