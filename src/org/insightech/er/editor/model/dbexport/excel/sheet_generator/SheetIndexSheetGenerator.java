package org.insightech.er.editor.model.dbexport.excel.sheet_generator;

import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFHyperlink;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Font;
import org.insightech.er.ResourceString;
import org.insightech.er.editor.model.ERDiagram;
import org.insightech.er.editor.model.ObjectModel;
import org.insightech.er.editor.model.dbexport.excel.ExportToExcelManager.LoopDefinition;
import org.insightech.er.editor.model.progress_monitor.ProgressMonitor;
import org.insightech.er.util.POIUtils;
import org.insightech.er.util.POIUtils.CellLocation;

public class SheetIndexSheetGenerator extends AbstractSheetGenerator {

    private static final String KEYWORD_SHEET_TYPE = "$SHTT";

    private static final String KEYWORD_NAME = "$NAM";

    private static final String KEYWORD_DESCRIPTION = "$DSC";

    private static final String KEYWORD_SHEET_NAME = "$SHTN";

    private static final String[] FIND_KEYWORDS_LIST = {KEYWORD_SHEET_TYPE, KEYWORD_NAME, KEYWORD_DESCRIPTION};

    /**
     * {@inheritDoc}
     */
    @Override
    public void generate(final ProgressMonitor monitor, final HSSFWorkbook workbook, final int sheetNo, final boolean useLogicalNameAsSheetName, final Map<String, Integer> sheetNameMap, final Map<String, ObjectModel> sheetObjectMap, final ERDiagram diagram, final Map<String, LoopDefinition> loopDefinitionMap) throws InterruptedException {

        final HSSFSheet sheet = workbook.getSheetAt(sheetNo);

        setSheetListData(workbook, sheet, sheetObjectMap, diagram);
    }

    public void setSheetListData(final HSSFWorkbook workbook, final HSSFSheet sheet, final Map<String, ObjectModel> sheetObjectMap, final ERDiagram diagram) {
        final CellLocation cellLocation = POIUtils.findCell(sheet, FIND_KEYWORDS_LIST);

        if (cellLocation != null) {
            int rowNum = cellLocation.r;
            final HSSFRow templateRow = sheet.getRow(rowNum);

            final ColumnTemplate columnTemplate = loadColumnTemplate(workbook, sheet, cellLocation);
            int order = 1;

            HSSFFont linkCellFont = null;
            int linkCol = -1;

            for (final Map.Entry<String, ObjectModel> entry : sheetObjectMap.entrySet()) {
                final String sheetName = entry.getKey();
                final ObjectModel objectModel = entry.getValue();

                final HSSFRow row = POIUtils.insertRow(sheet, rowNum++);

                for (final int columnNum : columnTemplate.columnTemplateMap.keySet()) {
                    final HSSFCell cell = row.createCell(columnNum);
                    final String template = columnTemplate.columnTemplateMap.get(columnNum);

                    String value = null;
                    if (KEYWORD_ORDER.equals(template)) {
                        value = String.valueOf(order);

                    } else {
                        if (KEYWORD_SHEET_TYPE.equals(template)) {
                            value = ResourceString.getResourceString("label.object.type." + objectModel.getObjectType());

                        } else if (KEYWORD_NAME.equals(template)) {
                            value = sheetName;
                            final HSSFHyperlink link = new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
                            link.setAddress("'" + sheetName + "'!A1");
                            cell.setHyperlink(link);

                            if (linkCellFont == null) {
                                linkCol = columnNum;

                                linkCellFont = POIUtils.copyFont(workbook, cell.getCellStyle().getFont(workbook));

                                linkCellFont.setColor(HSSFColor.BLUE.index);
                                linkCellFont.setUnderline(Font.U_SINGLE);
                            }

                        } else if (KEYWORD_DESCRIPTION.equals(template)) {
                            value = objectModel.getDescription();
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
            name = "List of sheets";
        }

        return name;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getTemplateSheetName() {
        return "sheet_index_template";
    }

    @Override
    public String[] getKeywords() {
        return new String[] {KEYWORD_SHEET_TYPE, KEYWORD_NAME, KEYWORD_DESCRIPTION, KEYWORD_ORDER, KEYWORD_SHEET_NAME};
    }

    @Override
    public int getKeywordsColumnNo() {
        return 24;
    }

    @Override
    public int count(final ERDiagram diagram) {
        return 1;
    }
}
