package org.insightech.er.editor.model.dbexport.excel.sheet_generator;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.insightech.er.editor.model.ERDiagram;
import org.insightech.er.editor.model.ObjectModel;
import org.insightech.er.editor.model.StringObjectModel;
import org.insightech.er.editor.model.dbexport.excel.ExportToExcelManager.LoopDefinition;
import org.insightech.er.editor.model.progress_monitor.ProgressMonitor;
import org.insightech.er.editor.model.tracking.ChangeTracking;
import org.insightech.er.util.POIUtils;
import org.insightech.er.util.POIUtils.CellLocation;

public class HistorySheetGenerator extends AbstractSheetGenerator {

    private static final String KEYWORD_DATE = "$DATE";

    private static final String KEYWORD_CONTENTS = "$CON";

    private static final String KEYWORD_DATE_FORMAT = "$FMT";

    private static final String KEYWORD_SHEET_NAME = "$SHTN";

    private static final String[] FIND_KEYWORDS_LIST = {KEYWORD_DATE, KEYWORD_CONTENTS};

    /**
     * {@inheritDoc}
     */
    @Override
    public void generate(final ProgressMonitor monitor, final HSSFWorkbook workbook, final int sheetNo, final boolean useLogicalNameAsSheetName, final Map<String, Integer> sheetNameMap, final Map<String, ObjectModel> sheetObjectMap, final ERDiagram diagram, final Map<String, LoopDefinition> loopDefinitionMap) throws InterruptedException {

        String sheetName = getSheetName();

        final HSSFSheet newSheet = createNewSheet(workbook, sheetNo, sheetName, sheetNameMap);

        sheetName = workbook.getSheetName(workbook.getSheetIndex(newSheet));
        monitor.subTaskWithCounter(sheetName);

        sheetObjectMap.put(sheetName, new StringObjectModel(sheetName));

        setHistoryListData(workbook, newSheet, sheetObjectMap, diagram);
        monitor.worked(1);
    }

    public void setHistoryListData(final HSSFWorkbook workbook, final HSSFSheet sheet, final Map<String, ObjectModel> sheetObjectMap, final ERDiagram diagram) {
        final CellLocation cellLocation = POIUtils.findCell(sheet, FIND_KEYWORDS_LIST);

        if (cellLocation != null) {
            int rowNum = cellLocation.r;
            final HSSFRow templateRow = sheet.getRow(rowNum);

            final ColumnTemplate columnTemplate = loadColumnTemplate(workbook, sheet, cellLocation);
            int order = 1;

            final HSSFFont linkCellFont = null;
            final int linkCol = -1;

            for (final ChangeTracking changeTracking : diagram.getChangeTrackingList().getList()) {
                final HSSFRow row = POIUtils.insertRow(sheet, rowNum++);

                for (final int columnNum : columnTemplate.columnTemplateMap.keySet()) {
                    final HSSFCell cell = row.createCell(columnNum);
                    final String template = columnTemplate.columnTemplateMap.get(columnNum);

                    String value = null;
                    if (KEYWORD_ORDER.equals(template)) {
                        value = String.valueOf(order);

                    } else {
                        if (KEYWORD_DATE.equals(template)) {
                            final DateFormat format = new SimpleDateFormat(keywordsValueMap.get(KEYWORD_DATE_FORMAT));
                            try {
                                value = format.format(changeTracking.getUpdatedDate());

                            } catch (final Exception e) {
                                value = changeTracking.getUpdatedDate().toString();
                            }

                        } else if (KEYWORD_CONTENTS.equals(template)) {
                            value = changeTracking.getComment();
                        }

                        final HSSFRichTextString text = new HSSFRichTextString(value);
                        cell.setCellValue(text);
                    }

                    order++;
                }
            }

            setCellStyle(columnTemplate, sheet, cellLocation.r, rowNum - cellLocation.r, templateRow.getFirstCellNum());

            if (linkCol != -1) {
                for (int row = cellLocation.r; row < rowNum; row++) {
                    final HSSFCell cell = sheet.getRow(row).getCell(linkCol);
                    cell.getCellStyle().setFont(linkCellFont);
                }
            }
        }
    }

    public String getSheetName() {
        String name = keywordsValueMap.get(KEYWORD_SHEET_NAME);

        if (name == null) {
            name = "dialog.title.change.tracking";
        }

        return name;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getTemplateSheetName() {
        return "history_template";
    }

    @Override
    public String[] getKeywords() {
        return new String[] {KEYWORD_DATE, KEYWORD_CONTENTS, KEYWORD_ORDER, KEYWORD_DATE_FORMAT, KEYWORD_SHEET_NAME};
    }

    @Override
    public int getKeywordsColumnNo() {
        return 28;
    }

    @Override
    public int count(final ERDiagram diagram) {
        return 1;
    }
}
