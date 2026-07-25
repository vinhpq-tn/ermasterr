package org.insightech.er.editor.model.dbexport.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.insightech.er.ResourceString;
import org.insightech.er.editor.model.ERDiagram;
import org.insightech.er.editor.model.ObjectModel;
import org.insightech.er.editor.model.StringObjectModel;
import org.insightech.er.editor.model.dbexport.AbstractExportManager;
import org.insightech.er.editor.model.dbexport.excel.sheet_generator.AbstractSheetGenerator;
import org.insightech.er.editor.model.dbexport.excel.sheet_generator.AllIndicesSheetGenerator;
import org.insightech.er.editor.model.dbexport.excel.sheet_generator.AllSequencesSheetGenerator;
import org.insightech.er.editor.model.dbexport.excel.sheet_generator.AllTablesSheetGenerator;
import org.insightech.er.editor.model.dbexport.excel.sheet_generator.AllTriggerSheetGenerator;
import org.insightech.er.editor.model.dbexport.excel.sheet_generator.AllViewSheetGenerator;
import org.insightech.er.editor.model.dbexport.excel.sheet_generator.CategorySheetGenerator;
import org.insightech.er.editor.model.dbexport.excel.sheet_generator.ColumnSheetGenerator;
import org.insightech.er.editor.model.dbexport.excel.sheet_generator.HistorySheetGenerator;
import org.insightech.er.editor.model.dbexport.excel.sheet_generator.IndexSheetGenerator;
import org.insightech.er.editor.model.dbexport.excel.sheet_generator.PictureSheetGenerator;
import org.insightech.er.editor.model.dbexport.excel.sheet_generator.SequenceSheetGenerator;
import org.insightech.er.editor.model.dbexport.excel.sheet_generator.SheetIndexSheetGenerator;
import org.insightech.er.editor.model.dbexport.excel.sheet_generator.TableSheetGenerator;
import org.insightech.er.editor.model.dbexport.excel.sheet_generator.TriggerSheetGenerator;
import org.insightech.er.editor.model.dbexport.excel.sheet_generator.ViewSheetGenerator;
import org.insightech.er.editor.model.dbexport.image.ExportToImageManager;
import org.insightech.er.editor.model.dbexport.image.ImageInfo;
import org.insightech.er.editor.model.dbexport.image.ImageInfoSet;
import org.insightech.er.editor.model.progress_monitor.ProgressMonitor;
import org.insightech.er.editor.model.settings.export.ExportExcelSetting;
import org.insightech.er.preference.PreferenceInitializer;
import org.insightech.er.preference.page.template.TemplatePreferencePage;
import org.insightech.er.util.Check;
import org.insightech.er.util.POIUtils;
import org.insightech.er.util.io.FileUtils;

public class ExportToExcelManager extends AbstractExportManager {

    private static final String WORDS_SHEET_NAME = "words";

    private static final String LOOPS_SHEET_NAME = "loops";

    private final Map<String, Integer> sheetNameMap;

    private final Map<String, ObjectModel> sheetObjectMap;

    private static final List<AbstractSheetGenerator> SHHET_GENERATOR_LIST = new ArrayList<AbstractSheetGenerator>();

    static {
        SHHET_GENERATOR_LIST.add(new TableSheetGenerator());
        SHHET_GENERATOR_LIST.add(new IndexSheetGenerator());
        SHHET_GENERATOR_LIST.add(new SequenceSheetGenerator());
        SHHET_GENERATOR_LIST.add(new ViewSheetGenerator());
        SHHET_GENERATOR_LIST.add(new TriggerSheetGenerator());
        SHHET_GENERATOR_LIST.add(new ColumnSheetGenerator());
        SHHET_GENERATOR_LIST.add(new AllTablesSheetGenerator());
        SHHET_GENERATOR_LIST.add(new AllIndicesSheetGenerator());
        SHHET_GENERATOR_LIST.add(new AllSequencesSheetGenerator());
        SHHET_GENERATOR_LIST.add(new AllViewSheetGenerator());
        SHHET_GENERATOR_LIST.add(new AllTriggerSheetGenerator());
        SHHET_GENERATOR_LIST.add(new CategorySheetGenerator());
        SHHET_GENERATOR_LIST.add(new HistorySheetGenerator());
    }

    public static class LoopDefinition {

        public int startLine;

        public int spaceLine;

        public String sheetName;

        public LoopDefinition(final int startLine, final int spaceLine, final String sheetName) {
            this.startLine = startLine;
            this.spaceLine = spaceLine;
            this.sheetName = sheetName;
        }
    }

    private PictureSheetGenerator pictureSheetGenerator;

    private SheetIndexSheetGenerator sheetIndexSheetGenerator;

    private final Map<String, LoopDefinition> loopDefinitionMap;

    private final ExportExcelSetting exportExcelSetting;

    private HSSFWorkbook workbook;

    private File excelFile;

    public ExportToExcelManager(final ExportExcelSetting exportExcelSetting) throws FileNotFoundException {
        super("dialog.message.export.excel");

        this.exportExcelSetting = exportExcelSetting;

        sheetNameMap = new HashMap<String, Integer>();
        sheetObjectMap = new LinkedHashMap<String, ObjectModel>();

        loopDefinitionMap = new HashMap<String, LoopDefinition>();
    }

    @Override
    public void init(final ERDiagram diagram, final File projectDir) throws Exception {
        super.init(diagram, projectDir);

        excelFile = FileUtils.getFile(this.projectDir, exportExcelSetting.getExcelOutput());
        excelFile.getParentFile().mkdirs();

        // this.backup(this.excelFile, true);

        InputStream templateStream = null;

        try {
            templateStream = getSelectedTemplate();
            workbook = loadTemplateWorkbook(templateStream, this.diagram);

        } finally {
            if (templateStream != null) {
                templateStream.close();
            }
        }

        // check whether the file is not opened by another process.
        POIUtils.writeExcelFile(excelFile, workbook);
    }

    private InputStream getSelectedTemplate() throws FileNotFoundException {
        if (!Check.isEmpty(exportExcelSetting.getExcelTemplatePath())) {
            return new FileInputStream(FileUtils.getFile(projectDir, exportExcelSetting.getExcelTemplatePath()));
        }

        final String lang = exportExcelSetting.getUsedDefaultTemplateLang();

        if ("en".equals(lang)) {
            return TemplatePreferencePage.getDefaultExcelTemplateEn();

        } else if ("zh".equals(lang)) {
            return TemplatePreferencePage.getDefaultExcelTemplateZh();

        } else if ("ja".equals(lang)) {
            return TemplatePreferencePage.getDefaultExcelTemplateJa();

        }

        final String templateName = exportExcelSetting.getExcelTemplate();

        final File file = new File(PreferenceInitializer.getTemplatePath(templateName));

        return new FileInputStream(file);
    }

    @Override
    protected int getTotalTaskCount() {
        return countSheetFromTemplate(workbook, diagram);
    }

    @Override
    protected void doProcess(final ProgressMonitor monitor) throws Exception {
        if (exportExcelSetting.isPutERDiagramOnExcel()) {
            pictureSheetGenerator = createPictureSheetGenerator(monitor, workbook);
        }

        createSheetFromTemplate(monitor, workbook, diagram, exportExcelSetting.isUseLogicalNameAsSheet());

        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            workbook.getSheetAt(i).setSelected(false);
        }

        if (workbook.getNumberOfSheets() > 0) {
            workbook.getSheetAt(0).setSelected(true);
            workbook.setActiveSheet(0);
            workbook.setFirstVisibleTab(0);
        }

        POIUtils.writeExcelFile(excelFile, workbook);
    }

    private PictureSheetGenerator createPictureSheetGenerator(final ProgressMonitor monitor, final HSSFWorkbook workbook) throws Exception {
        final ImageInfoSet imageInfoSet = ExportToImageManager.outputImage(diagram, exportExcelSetting.getCategory(), projectDir, monitor);

        final ImageInfo imageInfo = imageInfoSet.getDiagramImageInfo();

        return new PictureSheetGenerator(workbook, imageInfo.getImageData(), imageInfo.getExcelPictureType());
    }

    private HSSFWorkbook loadTemplateWorkbook(final InputStream template, final ERDiagram diagram) throws IOException {

        final HSSFWorkbook workbook = POIUtils.readExcelBook(template);

        if (workbook == null) {
            throw new IOException(ResourceString.getResourceString("error.read.file"));
        }

        final HSSFSheet wordsSheet = workbook.getSheet(WORDS_SHEET_NAME);

        if (wordsSheet == null) {
            throw new IOException(ResourceString.getResourceString("error.not.found.words.sheet"));
        }

        final HSSFSheet loopsSheet = workbook.getSheet(LOOPS_SHEET_NAME);

        if (loopsSheet == null) {
            throw new IOException(ResourceString.getResourceString("error.not.found.loops.sheet"));
        }

        initLoopDefinitionMap(loopsSheet);

        for (final AbstractSheetGenerator sheetGenerator : SHHET_GENERATOR_LIST) {
            sheetGenerator.init(wordsSheet);
        }

        sheetIndexSheetGenerator = new SheetIndexSheetGenerator();
        sheetIndexSheetGenerator.init(wordsSheet);

        return workbook;
    }

    private void initLoopDefinitionMap(final HSSFSheet loopsSheet) {
        for (int i = 2; i <= loopsSheet.getLastRowNum(); i++) {
            final String templateSheetName = POIUtils.getCellValue(loopsSheet, i, 0);
            if (templateSheetName == null) {
                break;
            }

            final int firstLine = POIUtils.getIntCellValue(loopsSheet, i, 1);
            final int spaceLine = POIUtils.getIntCellValue(loopsSheet, i, 2);
            final String sheetName = POIUtils.getCellValue(loopsSheet, i, 3);

            loopDefinitionMap.put(templateSheetName, new LoopDefinition(firstLine, spaceLine, sheetName));
        }
    }

    private AbstractSheetGenerator getSheetGenerator(final String templateSheetName) {
        for (final AbstractSheetGenerator sheetGenerator : SHHET_GENERATOR_LIST) {
            if (sheetGenerator.getTemplateSheetName().equals(templateSheetName)) {
                return sheetGenerator;
            }
        }

        return null;
    }

    private void initSheetNameMap(final HSSFWorkbook workbook) {
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            final String sheetName = workbook.getSheetName(i);
            sheetNameMap.put(sheetName.toUpperCase(), 0);
        }
    }

    private void createSheetFromTemplate(final ProgressMonitor monitor, final HSSFWorkbook workbook, final ERDiagram diagram, final boolean useLogicalNameAsSheetName) throws InterruptedException {
        initSheetNameMap(workbook);

        int originalSheetNum = workbook.getNumberOfSheets();

        int sheetIndexSheetNo = -1;

        while (originalSheetNum > 0) {
            final String templateSheetName = workbook.getSheetName(0);

            final AbstractSheetGenerator sheetGenerator = getSheetGenerator(templateSheetName);

            if (sheetGenerator != null) {
                sheetGenerator.generate(monitor, workbook, 0, useLogicalNameAsSheetName, sheetNameMap, sheetObjectMap, diagram, loopDefinitionMap);
                workbook.removeSheetAt(0);

            } else {
                if (!isExcludeTarget(templateSheetName)) {
                    moveSheet(workbook, 0);
                    final HSSFSheet sheet = workbook.getSheetAt(workbook.getNumberOfSheets() - 1);

                    sheetObjectMap.put(templateSheetName, new StringObjectModel(templateSheetName));

                    if (pictureSheetGenerator != null) {
                        pictureSheetGenerator.setImage(workbook, sheet);
                    }

                    if (sheetIndexSheetGenerator.getTemplateSheetName().equals(templateSheetName)) {
                        sheetIndexSheetNo = workbook.getNumberOfSheets() - originalSheetNum;

                        String name = sheetIndexSheetGenerator.getSheetName();

                        name = AbstractSheetGenerator.decideSheetName(name, sheetNameMap);

                        monitor.subTaskWithCounter(name);

                        workbook.setSheetName(workbook.getNumberOfSheets() - 1, name);
                    } else {
                        monitor.subTaskWithCounter(sheet.getSheetName());
                    }

                } else {
                    monitor.subTaskWithCounter("Removing template sheet");
                    workbook.removeSheetAt(0);
                }

                monitor.worked(1);
            }

            originalSheetNum--;
        }

        if (sheetIndexSheetNo != -1) {
            sheetIndexSheetGenerator.generate(monitor, workbook, sheetIndexSheetNo, useLogicalNameAsSheetName, sheetNameMap, sheetObjectMap, diagram, loopDefinitionMap);
        }
    }

    public static HSSFSheet moveSheet(final HSSFWorkbook workbook, final int sheetNo) {
        final HSSFSheet oldSheet = workbook.getSheetAt(sheetNo);
        final String sheetName = oldSheet.getSheetName();

        final HSSFSheet newSheet = workbook.cloneSheet(sheetNo);
        final int newSheetNo = workbook.getSheetIndex(newSheet);

        workbook.removeSheetAt(sheetNo);

        workbook.setSheetName(newSheetNo - 1, sheetName);

        return newSheet;
    }

    private int countSheetFromTemplate(final HSSFWorkbook workbook, final ERDiagram diagram) {
        int count = 0;

        for (int sheetNo = 0; sheetNo < workbook.getNumberOfSheets(); sheetNo++) {
            final String templateSheetName = workbook.getSheetName(sheetNo);

            final AbstractSheetGenerator sheetGenerator = getSheetGenerator(templateSheetName);

            if (sheetGenerator != null) {
                count += sheetGenerator.count(diagram);

            } else {
                count++;
            }
        }

        if (exportExcelSetting.isPutERDiagramOnExcel()) {
            count += 1;
        }

        return count;
    }

    private boolean isExcludeTarget(final String templateSheetName) {
        if (WORDS_SHEET_NAME.equals(templateSheetName) || LOOPS_SHEET_NAME.equals(templateSheetName)) {
            return true;
        }

        return false;
    }

    @Override
    public File getOutputFileOrDir() {
        return FileUtils.getFile(projectDir, exportExcelSetting.getExcelOutput());
    }
}
