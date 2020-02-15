package com.paipeng.pdf;


import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.PDPageTree;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.encryption.AccessPermission;
import org.apache.pdfbox.pdmodel.encryption.StandardProtectionPolicy;
import org.apache.pdfbox.pdmodel.font.PDFont;
import org.apache.pdfbox.pdmodel.font.PDType0Font;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.pdmodel.graphics.color.PDColor;
import org.apache.pdfbox.pdmodel.graphics.color.PDDeviceCMYK;
import org.apache.pdfbox.pdmodel.graphics.color.PDDeviceRGB;
import org.apache.pdfbox.pdmodel.graphics.image.LosslessFactory;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.pdfbox.pdmodel.interactive.annotation.PDAnnotationSquareCircle;
import org.apache.pdfbox.pdmodel.interactive.annotation.PDBorderEffectDictionary;
import org.apache.pdfbox.pdmodel.interactive.annotation.PDBorderStyleDictionary;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.PDFTextStripperByArea;
import org.apache.pdfbox.util.Matrix;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import static org.apache.pdfbox.pdmodel.interactive.annotation.PDBorderStyleDictionary.STYLE_SOLID;

public class PDFUtils {
    /**
     * 创建PDF
     *
     * @param path 文件路径
     * @param text 文本内容
     */
    public static void createPDF(String path, String fontPath, String text) {
        PDDocument doc = new PDDocument();

        int leading = 40;
        try {
            PDPage page = new PDPage();
            doc.addPage(page);

            PDFont font = PDType0Font.load(doc, new File(fontPath));

            PDPageContentStream content = new PDPageContentStream(doc, page);
            content.beginText();
            content.setFont(font, 12);
            content.newLineAtOffset(40, 780);
            String[] arr = text.split("\\n", -1);

            int num = 44;
            for (int i = 0; i < arr.length; i++) {
                int count = arr[i].length() / num;

                for (int j = 0; j < count; j++) {
                    content.showText(arr[i].substring(num * j, num * (j + 1)));
                    //content.newLine();
                    content.newLineAtOffset(0, -leading);
                }

                content.showText(arr[i].substring(num * count, arr[i].length()));
                //content.newLine();
                content.newLineAtOffset(0, -leading);
            }

            content.endText();
            content.close();

            doc.save(path);
            doc.close();
        } catch (Exception e) {
            System.out.println(e.toString());
        }
    }

    /**
     * 读取PDF
     *
     * @param path 文件路径
     */
    public static String readPDF(String path) {
        try {
            PDDocument doc = PDDocument.load(new File(path));
            PDFTextStripper textStripper = new PDFTextStripper();
            doc.close();
            return textStripper.getText(doc);
        } catch (IOException e) {
            e.printStackTrace();
            return "";
        }
    }

    /**
     * 获得文档总页数
     *
     * @param path 文件路径
     */
    public static int getTotalPage(String path) {
        try {
            PDDocument doc = PDDocument.load(new File(path));
            int pages = doc.getNumberOfPages();
            doc.close();
            return pages;
        } catch (Exception e) {
            return 0;
        }
    }

    /**
     * 读取PDF的某几页
     *
     * @param path      文件路径
     * @param startPage 开始页码
     * @param endPage   结束页码
     */
    public static String readPdfPage(String path, int startPage, int endPage) {
        try {
            PDDocument doc = PDDocument.load(new File(path));
            PDFTextStripper textStripper = new PDFTextStripper();
            textStripper.setStartPage(startPage);
            textStripper.setEndPage(endPage);
            String str = textStripper.getText(doc);
            doc.close();
            return str;
        } catch (IOException e) {
            e.printStackTrace();
            return "";
        }
    }

    /**
     * 读取PDF中的Table
     *
     * @param path 文件路径
     */
    public static String readPDFTable(String path) {
        StringBuilder sb = new StringBuilder();

        try {
            PDDocument doc = PDDocument.load(new File(path));

            PDPageTree pages = doc.getDocumentCatalog().getPages();

            String regionName = "page";

            for (int i = 0; i < pages.getCount(); i++) {
                PDPage page = pages.get(i);
                Rectangle rect = new Rectangle(0, 0, 2000, 2000);

                regionName = regionName + String.valueOf(i);

                PDFTextStripperByArea stripper = new PDFTextStripperByArea();
                stripper.addRegion(regionName, rect);
                stripper.extractRegions(page);

                sb.append(stripper.getTextForRegion(regionName));
            }

            doc.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

        return sb.toString();
    }

    /**
     * pdf转成图片
     *
     * @param path      文件路径
     * @param imagePath 图片路径
     */
    public static void pdfTranslateImage(String path, String imagePath) {
        PDDocument doc = null;

        try {
            doc = PDDocument.load(new File(path));

            PDFRenderer renderer = new PDFRenderer(doc);

            BufferedImage image = renderer.renderImage(0);

            ImageIO.write(image, "JPEG", new File(imagePath));

            doc.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void insertImage(PDDocument pdDocument, int pageNumber,  BufferedImage bufferedImage, float x, float y, int dpi) throws IOException {
        PDPage page = pdDocument.getDocumentCatalog().getPages().get(pageNumber);

        PDImageXObject imageXObject = LosslessFactory.createFromImage(pdDocument, bufferedImage);

        float scale = 72f / dpi;

        PDPageContentStream contentStream = new PDPageContentStream(pdDocument, page, PDPageContentStream.AppendMode.APPEND, false);
        contentStream.drawImage(imageXObject, x, y, imageXObject.getWidth() * scale, imageXObject.getHeight() * scale);
        contentStream.close();
    }


    public static void insertImage(PDDocument pdDocument, PDPage pdPage,  BufferedImage bufferedImage, float x, float y, int dpi) throws IOException {
        PDImageXObject imageXObject = LosslessFactory.createFromImage(pdDocument, bufferedImage);
        float scale = 72f / dpi;
        PDPageContentStream contentStream = new PDPageContentStream(pdDocument, pdPage, PDPageContentStream.AppendMode.APPEND, false);
        contentStream.drawImage(imageXObject, x, y, imageXObject.getWidth() * scale, imageXObject.getHeight() * scale);
        contentStream.close();
    }


    public static void insertImage(PDDocument pdDocument, PDPage pdPage,  PDImageXObject imageXObject, float x, float y, int dpi) throws IOException {
        float scale = 72f / dpi;
        PDPageContentStream contentStream = new PDPageContentStream(pdDocument, pdPage, PDPageContentStream.AppendMode.APPEND, false);
        contentStream.drawImage(imageXObject, x, y, imageXObject.getWidth() * scale, imageXObject.getHeight() * scale);
        contentStream.close();
    }


    public static void insertImage(PDDocument pdDocument, PDPage pdPage,  BufferedImage bufferedImage, float x, float y, float width, float height) throws IOException {
        PDImageXObject imageXObject = LosslessFactory.createFromImage(pdDocument, bufferedImage);
        PDPageContentStream contentStream = new PDPageContentStream(pdDocument, pdPage, PDPageContentStream.AppendMode.APPEND, false);
        if (width != 0 && height == 0) {
            height = imageXObject.getHeight() * width /imageXObject.getWidth();
        } else if (width == 0 && height != 0) {
            width = imageXObject.getWidth() * height /imageXObject.getHeight();
        } else if (width == 0 && height == 0) {
            width = imageXObject.getWidth();
            height = imageXObject.getHeight();
        }
        contentStream.drawImage(imageXObject, x, y, width, height);
        contentStream.close();
    }


    public static void encryptPDF(PDDocument pdDocument, String ownerPassword, String userPassword) throws IOException {
        int keyLength = 256;

        AccessPermission ap = new AccessPermission();

        // disable printing, everything else is allowed
        ap.setCanPrint(false);
        ap.setCanAssembleDocument(false);
        ap.setCanExtractContent(false);
        ap.setCanModify(false);
        ap.setCanPrintDegraded(false);
        ap.setReadOnly();
        ap.setCanModifyAnnotations(false);

        // Owner password (to open the file with all permissions) is "12345"
        // User password (to open the file but with restricted permissions, is empty here)
        StandardProtectionPolicy spp = new StandardProtectionPolicy(ownerPassword, userPassword, ap);
        spp.setEncryptionKeyLength(keyLength);
        spp.setPermissions(ap);
        pdDocument.protect(spp);

        //pdDocument.save("/Users/paipeng/testee.pdf");

    }

    public static void insertText(PDDocument pdDocument, int pageNumber, String text, int tx, int ty, String fontPath, int fontSize, int fontColor) throws IOException{
        PDPage page = pdDocument.getDocumentCatalog().getPages().get(pageNumber);
        if (page != null) {
            PDFont font = PDType0Font.load(pdDocument, new File(fontPath));

            PDPageContentStream pdPageContentStream = new PDPageContentStream(pdDocument, page, PDPageContentStream.AppendMode.APPEND, false);
            pdPageContentStream.beginText();
            pdPageContentStream.setFont(font, fontSize);
            pdPageContentStream.setNonStrokingColor(new PDColor(new float[]{0,1/255,1f/255}, PDDeviceRGB.INSTANCE));
            pdPageContentStream.newLineAtOffset(tx, ty);

            pdPageContentStream.showText(text);


            pdPageContentStream.endText();
            pdPageContentStream.close();
        }

    }

    public static void insertText(PDDocument pdDocument, PDPage page, String text, float offsetX, float offsetY, float contentWidth, float contentHeight, String fontPath, int fontSize, int fontColor) throws IOException{


        PDFont font = PDType0Font.load(pdDocument, new File(fontPath));

        float titleWidth = font.getStringWidth(text) / 1000 * fontSize;
        float titleHeight = font.getFontDescriptor().getFontBoundingBox().getHeight() / 1000 * fontSize;

        float shiftX = (contentWidth - titleWidth)/2;
        float shiftY = 0;//(contentHeight - titleHeight)/2;


        PDPageContentStream pdPageContentStream = new PDPageContentStream(pdDocument, page, PDPageContentStream.AppendMode.APPEND, false);
        pdPageContentStream.beginText();
        pdPageContentStream.setFont(font, fontSize);
        pdPageContentStream.setNonStrokingColor(new PDColor(new float[]{0,1/255,1f/255}, PDDeviceRGB.INSTANCE));
        pdPageContentStream.newLineAtOffset(offsetX + shiftX, offsetY + shiftY);

        pdPageContentStream.showText(text);

        pdPageContentStream.endText();
        pdPageContentStream.close();
    }

    public static void insertTextWithRotate(PDDocument pdDocument, PDPage page, String text, float offsetX, float offsetY, float contentWidth, float contentHeight, String fontPath, int fontSize, int fontColor) throws IOException{
        PDFont font = PDType0Font.load(pdDocument, new File(fontPath));

        float titleWidth = font.getStringWidth(text) / 1000 * fontSize;
        float titleHeight = font.getFontDescriptor().getFontBoundingBox().getHeight() / 1000 * fontSize;

        float shiftX = (contentWidth - titleWidth)/2;
        float shiftY = titleHeight/2;

        float centeredXPosition = offsetX + shiftX + titleWidth/2 ;//(page.getMediaBox().getWidth() - fontSize/1000f)/2f;

        float centeredYPosition = offsetY + shiftY + titleHeight/2;//(page.getMediaBox().getHeight() - (stringWidth*fontSize)/1000f)/3f;


        PDPageContentStream pdPageContentStream = new PDPageContentStream(pdDocument, page, PDPageContentStream.AppendMode.APPEND, false);
        pdPageContentStream.beginText();
        pdPageContentStream.setFont(font, fontSize);
        pdPageContentStream.setNonStrokingColor(new PDColor(new float[]{0,1/255,1f/255}, PDDeviceRGB.INSTANCE));
        pdPageContentStream.newLineAtOffset(offsetX + shiftX, offsetY + shiftY);


        pdPageContentStream.setTextMatrix(Matrix.getRotateInstance(90*Math.PI*0.25, centeredXPosition,
                centeredYPosition));


        pdPageContentStream.showText(text);

        pdPageContentStream.endText();
        pdPageContentStream.close();
    }

    public static void drawRect(PDDocument pdDocument, PDPage pdPage, float x, float y, float width, float height) throws IOException {
        //PDPage page = pdDocument.getDocumentCatalog().getPages().get(pageNumber);
        if (pdPage != null) {
            PDPageContentStream pdPageContentStream = new PDPageContentStream(pdDocument, pdPage, PDPageContentStream.AppendMode.APPEND, false);
            pdPageContentStream.setNonStrokingColor(Color.getHSBColor(61f/100, 50f/100, 90f/100));
            pdPageContentStream.addRect(x, y, width, height);
            //pdPageContentStream.fill();
            //pdPageContentStream.setLineDashPattern(new float[]{1,2,1}, 0);
            //pdPageContentStream.setStrokingColor(Color.DARK_GRAY);
            pdPageContentStream.setStrokingColor(1.0f, 0, 0, 0);
            pdPageContentStream.stroke();

            pdPageContentStream.close();
        }
    }

    public static void drawPrintFocus(PDDocument pdDocument, PDPage pdPage, float border) throws IOException {
        //PDPage page = pdDocument.getDocumentCatalog().getPages().get(pageNumber);
        if (pdPage != null) {
            PDRectangle pdRectangle = pdPage.getMediaBox();

            System.out.println("pdRectangle: " + pdRectangle.toString());

            int length = 7;
            int length2 = 5;


            float bxa[] = new float[4];
            float bya[] = new float[4];


            // left bottom
            float bx = border/2;
            float by = border/2;
            bxa[0] = bx;
            bya[0] = by;
            // left top
            bxa[1] = border/2;
            bya[1] = pdRectangle.getUpperRightY() - border/2;
            // right top
            bxa[2] = pdRectangle.getUpperRightX() - border/2;
            bya[2] = pdRectangle.getUpperRightY() - border/2;
            // right bottom
            bxa[3] = pdRectangle.getUpperRightX() - border/2;
            bya[3] = border/2;


            PDPageContentStream pdPageContentStream = new PDPageContentStream(pdDocument, pdPage, PDPageContentStream.AppendMode.APPEND, false);
            pdPageContentStream.setLineWidth(0.5f);
            pdPageContentStream.setStrokingColor(1.0f, 1, 1, 0);

            for (int i = 0; i < bxa.length; i++) {
                pdPageContentStream.moveTo(bxa[i] - length / 2.0f, bya[i]);
                pdPageContentStream.lineTo(bxa[i] + length / 2f, bya[i]);
                pdPageContentStream.closeAndStroke();

                pdPageContentStream.moveTo(bxa[i], bya[i] - length / 2f);
                pdPageContentStream.lineTo(bxa[i], bya[i] + length / 2f);
                pdPageContentStream.closeAndStroke();

                PDAnnotationSquareCircle circle = new PDAnnotationSquareCircle(PDAnnotationSquareCircle.SUB_TYPE_CIRCLE);
                PDRectangle position = new PDRectangle();
                position.setLowerLeftX(bxa[i] - length2 / 2.0f);
                position.setLowerLeftY(bya[i] - length2 / 2f);
                position.setUpperRightX(bxa[i] + length2 / 2.0f);
                position.setUpperRightY(bya[i] + length2 / 2f);
                circle.setRectangle(position);
                //circle.setInteriorColor(new PDColor(new float[]{0,1/255,1f/255}, PDDeviceRGB.INSTANCE));
                PDBorderEffectDictionary pdBorderEffectDictionary = new PDBorderEffectDictionary();

                pdBorderEffectDictionary.setStyle(STYLE_SOLID);
                circle.setBorderEffect(pdBorderEffectDictionary);

                PDBorderStyleDictionary thickness = new PDBorderStyleDictionary();
                thickness.setWidth((float)0.5);
                circle.setColor(new PDColor(new float[]{1f, 1f, 1f, 0}, PDDeviceCMYK.INSTANCE));
                circle.setBorderStyle(thickness);

                pdPage.getAnnotations().add(circle);
            }

            pdPageContentStream.close();
        }
    }

    public static void drawPrintColor(PDDocument pdDocument, PDPage pdPage, float border) throws IOException {
        //PDPage page = pdDocument.getDocumentCatalog().getPages().get(pageNumber);
        if (pdPage != null) {
            PDRectangle pdRectangle = pdPage.getMediaBox();

            PDPageContentStream pdPageContentStream = new PDPageContentStream(pdDocument, pdPage, PDPageContentStream.AppendMode.APPEND, false);

            int length = 6;
            float bx = border/2 - length/2f;
            float by = pdRectangle.getUpperRightY() - border/2 - length - 10;

            pdPageContentStream.setNonStrokingColor(1f, 0, 0, 0);
            pdPageContentStream.addRect(bx, by, length, length);
            pdPageContentStream.fill();


            by = pdRectangle.getUpperRightY() - border/2 - length*2 - 10;
            pdPageContentStream.setNonStrokingColor(0f, 1, 0, 0);
            pdPageContentStream.addRect(bx, by, length, length);
            pdPageContentStream.fill();

            by = pdRectangle.getUpperRightY() - border/2 - length*3 - 10;
            pdPageContentStream.setNonStrokingColor(0f, 0, 1, 0);
            pdPageContentStream.addRect(bx, by, length, length);
            pdPageContentStream.fill();

            by = pdRectangle.getUpperRightY() - border/2 - length*4 - 10;
            pdPageContentStream.setNonStrokingColor(0f, 0, 0, 1);
            pdPageContentStream.addRect(bx, by, length, length);
            pdPageContentStream.fill();


            pdPageContentStream.close();
        }
    }


    public static void drawPrintText(PDDocument pdDocument, PDPage pdPage, float border, String pageTitle, String firstSerialNumber, String lastSerialNumber) throws IOException {
        //PDPage page = pdDocument.getDocumentCatalog().getPages().get(pageNumber);
        if (pdPage != null) {
            PDRectangle pdRectangle = pdPage.getMediaBox();

            PDPageContentStream pdPageContentStream = new PDPageContentStream(pdDocument, pdPage, PDPageContentStream.AppendMode.APPEND, false);

            int length = 6;
            float bx = border/2 - length/2f;
            float by = pdRectangle.getUpperRightY() - border/2 - length - 10;

            int fontSize = 6;
            PDFont font = PDType1Font.HELVETICA;
            //PDFont font = PDType0Font.load(pdDocument, new File(PdfUtils.class.getResource("/pdf/msyh.ttf").getPath()));



            if (pageTitle != null) {
                pdPageContentStream.beginText();
                pdPageContentStream.setFont(font, fontSize);
                pdPageContentStream.setNonStrokingColor(new PDColor(new float[]{0, 0, 0, 1}, PDDeviceCMYK.INSTANCE));
                float titleWidth = font.getStringWidth(pageTitle) / 1000 * fontSize;
                float titleHeight = font.getFontDescriptor().getFontBoundingBox().getHeight() / 1000 * fontSize;
                pdPageContentStream.newLineAtOffset((pdRectangle.getWidth() - titleWidth) / 2f, border / 2 - titleHeight / 3);

                pdPageContentStream.showText(pageTitle);
                pdPageContentStream.endText();
            }

            if (firstSerialNumber != null) {
                pdPageContentStream.beginText();
                pdPageContentStream.setFont(font, fontSize);
                pdPageContentStream.setNonStrokingColor(new PDColor(new float[]{0, 0, 0, 1}, PDDeviceCMYK.INSTANCE));
                //float titleWidth = font.getStringWidth(firstSerialNumber) / 1000 * fontSize;
                float titleHeight = font.getFontDescriptor().getFontBoundingBox().getHeight() / 1000 * fontSize;
                pdPageContentStream.newLineAtOffset(border + 5, pdRectangle.getUpperRightY() - border / 2 - titleHeight / 3);
                pdPageContentStream.showText(firstSerialNumber);
                pdPageContentStream.endText();
            }
            if (lastSerialNumber != null) {
                pdPageContentStream.beginText();
                pdPageContentStream.setFont(font, fontSize);
                pdPageContentStream.setNonStrokingColor(new PDColor(new float[]{0, 0, 0, 1}, PDDeviceCMYK.INSTANCE));
                float titleWidth = font.getStringWidth(firstSerialNumber) / 1000 * fontSize;
                float titleHeight = font.getFontDescriptor().getFontBoundingBox().getHeight() / 1000 * fontSize;
                pdPageContentStream.newLineAtOffset(pdRectangle.getUpperRightX() - border - 5 - titleWidth, pdRectangle.getUpperRightY() -border / 2 - titleHeight / 3);
                pdPageContentStream.showText(lastSerialNumber);
                pdPageContentStream.endText();
            }



            pdPageContentStream.close();
        }
    }

    public static void drawLayoutLevel3(PDDocument pdDocument, PDPage pdPage, float border) throws IOException {

        List<Point> points = new ArrayList<>();

        float contentWidth = pdPage.getMediaBox().getUpperRightX() - border * 2;
        float contentHeight = pdPage.getMediaBox().getUpperRightY() - border * 2;


        Point point = new Point();
        point.x = (int)border;
        point.y = (int)(border + contentHeight * 2/5);


        Point point1 = new Point();
        point1.x = (int)(border + contentWidth);
        point1.y = (int)(border + contentHeight * 2/5);

        points.add(point);
        points.add(point1);



        point = new Point();
        point.x = (int)border;
        point.y = (int)(border + contentHeight * 4/5);


        point1 = new Point();
        point1.x = (int)(border + contentWidth);
        point1.y = (int)(border + contentHeight * 4/5);

        points.add(point);
        points.add(point1);


        Point point2 = new Point();
        point2.x = (int)(border + contentWidth/2);
        point2.y = (int)(border);
        points.add(point2);

        Point point3 = new Point();
        point3.x = (int)(border + contentWidth/2);
        point3.y = (int)(border + contentHeight);


        points.add(point3);



        point = new Point();
        point.x = (int)border;
        point.y = (int)(border + contentHeight * 1/5);


        point1 = new Point();
        point1.x = (int)(border + contentWidth);
        point1.y = (int)(border + contentHeight * 1/5);

        points.add(point);
        points.add(point1);

        point = new Point();
        point.x = (int)border;
        point.y = (int)(border + contentHeight * 3/5);


        point1 = new Point();
        point1.x = (int)(border + contentWidth);
        point1.y = (int)(border + contentHeight * 3/5);

        points.add(point);
        points.add(point1);


        PDPageContentStream pdPageContentStream = new PDPageContentStream(pdDocument, pdPage, PDPageContentStream.AppendMode.APPEND, false);
        pdPageContentStream.setLineWidth(0.5f);
        pdPageContentStream.setStrokingColor(1.0f, 1, 1, 0);

        for (int i = 0; i < points.size(); i+=2) {
            pdPageContentStream.moveTo(points.get(i).x, points.get(i).y);
            pdPageContentStream.lineTo(points.get(i+1).x, points.get(i+1).y);
            pdPageContentStream.closeAndStroke();
        }


        pdPageContentStream.close();
    }
}