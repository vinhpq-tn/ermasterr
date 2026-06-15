package org.insightech.er.editor.model.dbexport.excel.sheet_generator;

import java.awt.Dimension;

import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.insightech.er.util.POIUtils;
import org.insightech.er.util.POIUtils.CellLocation;

public class PictureSheetGenerator {

    private static final String KEYWORD_ER = "$ER";

    private final byte[] imageBuffer;

    private int pictureIndex;

    private final int excelPictureType;

    public PictureSheetGenerator(final XSSFWorkbook workbook, final byte[] imageBuffer, final int excelPictureType) {
        this.imageBuffer = imageBuffer;
        this.excelPictureType = excelPictureType;

        if (this.imageBuffer != null) {
            pictureIndex = workbook.addPicture(this.imageBuffer, this.excelPictureType);
        }
    }

    public void setImage(final XSSFWorkbook workbook, final XSSFSheet sheet) {
        final CellLocation cellLocation = POIUtils.findMatchCell(sheet, "\\" + KEYWORD_ER + ".*");

        if (cellLocation != null) {
            int width = -1;
            int height = -1;

            final String value = POIUtils.getCellValue(sheet, cellLocation);

            final int startIndex = value.indexOf("(");
            if (startIndex != -1) {
                final int middleIndex = value.indexOf(",", startIndex + 1);
                if (middleIndex != -1) {
                    width = Integer.parseInt(value.substring(startIndex + 1, middleIndex).trim());
                    height = Integer.parseInt(value.substring(middleIndex + 1, value.length() - 1).trim());
                }
            }

            this.setImage(workbook, sheet, cellLocation, width, height);
        }
    }

    private void setImage(final XSSFWorkbook workbook, final XSSFSheet sheet, final CellLocation cellLocation, int width, int height) {
        POIUtils.setCellValue(sheet, cellLocation, "");

        if (imageBuffer != null) {
            final XSSFDrawing patriarch = sheet.createDrawingPatriarch();
            final XSSFClientAnchor preferredSize = getPreferredSize(sheet, new XSSFClientAnchor(0, 0, 0, 0, (short) cellLocation.c, cellLocation.r, (short) 0, 0), width, height);
            final XSSFPicture picture = patriarch.createPicture(preferredSize, pictureIndex);

            final Dimension dimension = picture.getImageDimension();
            final float rate = (float) dimension.width / (float) dimension.height;
            final float specifiedRate = (float) width / (float) height;

            if (width == -1 || height == -1) {
                width = dimension.width;
                height = dimension.height;

            } else {
                if (rate > specifiedRate) {
                    if (dimension.width > width) {
                        height = (int) (width / rate);

                    } else {
                        width = dimension.width;
                        height = dimension.height;
                    }

                } else {
                    if (dimension.height > height) {
                        width = (int) (height * rate);

                    } else {
                        width = dimension.width;
                        height = dimension.height;
                    }
                }
            }
        }
    }

    public XSSFClientAnchor getPreferredSize(final XSSFSheet sheet, final XSSFClientAnchor anchor, final int width, final int height) {
        float w = 0.0F;
        w += getColumnWidthInPixels(sheet, anchor.getCol1()) * (1 - anchor.getDx1() / 1024);

        short col2 = (short) (anchor.getCol1() + 1);
        int dx2 = 0;
        for (; w < width; w += getColumnWidthInPixels(sheet, col2++));
        if (w > width) {
            col2--;
            final float cw = getColumnWidthInPixels(sheet, col2);
            final float delta = w - width;
            dx2 = (int) (((cw - delta) / cw) * 1024F);
        }
        anchor.setCol2(col2);
        anchor.setDx2(dx2);
        float h = 0.0F;
        h += (1 - anchor.getDy1() / 256) * getRowHeightInPixels(sheet, anchor.getRow1());
        int row2 = anchor.getRow1() + 1;
        int dy2 = 0;
        for (; h < height; h += getRowHeightInPixels(sheet, row2++));
        if (h > height) {
            row2--;
            final float ch = getRowHeightInPixels(sheet, row2);
            final float delta = h - height;
            dy2 = (int) (((ch - delta) / ch) * 256F);
        }
        anchor.setRow2(row2);
        anchor.setDy2(dy2);
        return anchor;
    }

    private float getColumnWidthInPixels(final XSSFSheet sheet, final int column) {
        final int cw = sheet.getColumnWidth(column);
        final float px = getPixelWidth(sheet, column);
        return cw / px;
    }

    private float getRowHeightInPixels(final XSSFSheet sheet, final int i) {
        final XSSFRow row = sheet.getRow(i);
        float height;
        if (row != null) {
            height = row.getHeight();
        } else {
            height = sheet.getDefaultRowHeight();
        }

        return height / 15F;
    }

    private float getPixelWidth(final XSSFSheet sheet, final int column) {
        final int def = sheet.getDefaultColumnWidth() * 256;
        final int cw = sheet.getColumnWidth(column);
        return cw != def ? 36.56F : 32F;
    }

}
