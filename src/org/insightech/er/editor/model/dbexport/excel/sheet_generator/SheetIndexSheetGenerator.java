package org.insightech.er.editor.model.dbexport.excel.sheet_generator;

import java.util.Map;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
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
    public void generate(final ProgressMonitor monitor, final XSSFWorkbook workbook, final int sheetNo, final boolean useLogicalNameAsSheetName, final Map<String, Integer> sheetNameMap, final Map<String, ObjectModel> sheetObjectMap, final ERDiagram diagram, final Map<String, LoopDefinition> loopDefinitionMap) throws InterruptedException {

        final XSSFSheet sheet = workbook.getSheetAt(sheetNo);

        setSheetListData(workbook, sheet, sheetObjectMap, diagram);
    }

    public void setSheetListData(final XSSFWorkbook workbook, final XSSFSheet sheet, final Map<String, ObjectModel> sheetObjectMap, final ERDiagram diagram) {
        final CellLocation cellLocation = POIUtils.findCell(sheet, FIND_KEYWORDS_LIST);

        if (cellLocation != null) {
            int rowNum = cellLocation.r;
            final XSSFRow templateRow = sheet.getRow(rowNum);

            final ColumnTemplate columnTemplate = loadColumnTemplate(workbook, sheet, cellLocation);
            int order = 1;

            XSSFFont linkCellFont = null;
            int linkCol = -1;

            for (final Map.Entry<String, ObjectModel> entry : sheetObjectMap.entrySet()) {
                final String sheetName = entry.getKey();
                final ObjectModel objectModel = entry.getValue();

                final XSSFRow row = POIUtils.insertRow(sheet, rowNum++);

                for (final int columnNum : columnTemplate.columnTemplateMap.keySet()) {
                    final XSSFCell cell = row.createCell(columnNum);
                    final String template = columnTemplate.columnTemplateMap.get(columnNum);

                    String value = null;
                    if (KEYWORD_ORDER.equals(template)) {
                        value = String.valueOf(order);

                    } else {
                        if (KEYWORD_SHEET_TYPE.equals(template)) {
                            value = ResourceString.getResourceString("label.object.type." + objectModel.getObjectType());

                        } else if (KEYWORD_NAME.equals(template)) {
                        	value = sheetName;
                        	XSSFCreationHelper createHelper = workbook.getCreationHelper();
                            final XSSFHyperlink link = createHelper.createHyperlink(HyperlinkType.DOCUMENT);
                            link.setAddress("'" + sheetName + "'!A1");
                            cell.setHyperlink(link);

                            if (linkCellFont == null) {
                                linkCol = columnNum;

                                linkCellFont = POIUtils.copyFont(workbook, cell.getCellStyle().getFont());

                                linkCellFont.setColor(new XSSFColor(new byte[] {0, 0, (byte) 255}));
                                linkCellFont.setUnderline(Font.U_SINGLE);
                            }

                        } else if (KEYWORD_DESCRIPTION.equals(template)) {
                            value = objectModel.getDescription();
                        }

                        final XSSFRichTextString text = new XSSFRichTextString(value);
                        cell.setCellValue(text);
                    }

                    order++;
                }
            }

            setCellStyle(columnTemplate, sheet, cellLocation.r, rowNum - cellLocation.r, templateRow.getFirstCellNum());

            if (linkCol != -1) {
                for (int row = cellLocation.r; row < rowNum; row++) {
                    final XSSFCell cell = sheet.getRow(row).getCell(linkCol);
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
