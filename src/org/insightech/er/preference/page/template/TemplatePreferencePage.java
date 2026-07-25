package org.insightech.er.preference.page.template;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.eclipse.jface.preference.PreferencePage;
import org.eclipse.swt.SWT;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.layout.GridLayout;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.widgets.Composite;
import org.eclipse.swt.widgets.Control;
import org.eclipse.ui.IWorkbench;
import org.eclipse.ui.IWorkbenchPreferencePage;
import org.insightech.er.ERDiagramActivator;
import org.insightech.er.ResourceString;
import org.insightech.er.Resources;
import org.insightech.er.common.widgets.CompositeFactory;
import org.insightech.er.preference.PreferenceInitializer;
import org.insightech.er.util.io.IOUtils;

public class TemplatePreferencePage extends PreferencePage implements IWorkbenchPreferencePage {

    private static final String DEFAULT_TEMPLATE_FILE_EN = "template_en.xls";

    private static final String DEFAULT_TEMPLATE_FILE_ZH = "template_zh.xls";

    private static final String DEFAULT_TEMPLATE_FILE_JA = "template_ja.xls";

    private TemplateFileListEditor fileListEditor;

    @Override
    public void init(final IWorkbench workbench) {}

    @Override
    protected Control createContents(final Composite parent) {
        final Composite composite = new Composite(parent, SWT.NONE);
        final GridLayout layout = new GridLayout();
        layout.numColumns = 2;
        composite.setLayout(layout);

        final Composite buttonComposite = CompositeFactory.createChildComposite(composite, 2, 2);
        createButtonComposite(buttonComposite);

        CompositeFactory.fillLine(composite, Resources.PREFERENCE_PAGE_MARGIN_TOP);

        fileListEditor = new TemplateFileListEditor(PreferenceInitializer.TEMPLATE_FILE_LIST, ResourceString.getResourceString("label.custom.tempplate"), composite);
        fileListEditor.load();

        CompositeFactory.fillLine(composite);

        CompositeFactory.createLabel(composite, "dialog.message.template.file.store", 2);

        return composite;
    }

    private void createButtonComposite(final Composite composite) {
        final GridLayout layout = new GridLayout();
        layout.numColumns = 4;
        composite.setLayout(layout);

        final Button buttonEn = new Button(composite, SWT.NONE);
        buttonEn.setText(ResourceString.getResourceString("label.button.download.template.en"));
        buttonEn.addSelectionListener(new SelectionAdapter() {

            /**
             * {@inheritDoc}
             */
            @Override
            public void widgetSelected(final SelectionEvent e) {
                download(DEFAULT_TEMPLATE_FILE_EN);
            }
        });

        final Button buttonZh = new Button(composite, SWT.NONE);
        buttonZh.setText(ResourceString.getResourceString("label.button.download.template.zh"));
        buttonZh.addSelectionListener(new SelectionAdapter() {

            /**
             * {@inheritDoc}
             */
            @Override
            public void widgetSelected(final SelectionEvent e) {
                download(DEFAULT_TEMPLATE_FILE_ZH);
            }
        });

        final Button buttonJa = new Button(composite, SWT.NONE);
        buttonJa.setText(ResourceString.getResourceString("label.button.download.template.ja"));
        buttonJa.addSelectionListener(new SelectionAdapter() {

            /**
             * {@inheritDoc}
             */
            @Override
            public void widgetSelected(final SelectionEvent e) {
                download(DEFAULT_TEMPLATE_FILE_JA);
            }
        });

    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected void performDefaults() {
        fileListEditor.loadDefault();

        super.performDefaults();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean performOk() {
        fileListEditor.store();

        return super.performOk();
    }

    private void download(final String fileName) {
        final String filePath = ERDiagramActivator.showSaveDialog(null, fileName, fileName, new String[] {".xls"}, true);

        if (filePath != null) {
            InputStream in = null;
            OutputStream out = null;
            try {
                in = this.getClass().getResourceAsStream("/" + fileName);
                out = new FileOutputStream(filePath);

                IOUtils.copy(in, out);
            } catch (final IOException ioe) {
                ERDiagramActivator.showExceptionDialog(ioe);

            } finally {
                if (in != null) {
                    try {
                        in.close();
                    } catch (final IOException e1) {
                        ERDiagramActivator.showExceptionDialog(e1);
                    }

                }
                if (out != null) {
                    try {
                        out.close();
                    } catch (final IOException e1) {
                        ERDiagramActivator.showExceptionDialog(e1);
                    }
                }

            }
        }
    }

    public static InputStream getDefaultExcelTemplateEn() {
        return TemplatePreferencePage.class.getResourceAsStream("/" + DEFAULT_TEMPLATE_FILE_EN);
    }

    public static InputStream getDefaultExcelTemplateZh() {
        return TemplatePreferencePage.class.getResourceAsStream("/" + DEFAULT_TEMPLATE_FILE_ZH);
    }

    public static InputStream getDefaultExcelTemplateJa() {
        return TemplatePreferencePage.class.getResourceAsStream("/" + DEFAULT_TEMPLATE_FILE_JA);
    }
}
