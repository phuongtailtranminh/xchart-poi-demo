package me.phuongtm.demo.xchartworddemo;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.knowm.xchart.BitmapEncoder;
import org.knowm.xchart.QuickChart;
import org.knowm.xchart.XYChart;

import java.io.*;

public class Application {

    public static final String IMAGES_DOCX = "images.docx";

    public static void main(String[] args) throws IOException, InvalidFormatException {
        writeDoc(drawSimpleLineChart(), IMAGES_DOCX);
    }

    private static byte[] drawSimpleLineChart() throws IOException {
        double[] xData = new double[] { 0.0, 1.0, 2.0 };
        double[] yData = new double[] { 2.0, 1.0, 0.0 };
        // Create Chart
        XYChart chart = QuickChart.getChart("Sample Chart", "X", "Y", "y(x)", xData, yData);
        return BitmapEncoder.getBitmapBytes(chart, BitmapEncoder.BitmapFormat.PNG);
    }

    private static void writeDoc(byte[] imageBytes, String fileName) throws InvalidFormatException, IOException {
        InputStream is = new ByteArrayInputStream(imageBytes);
        XWPFDocument doc = new XWPFDocument();
        XWPFParagraph p = doc.createParagraph();
        XWPFRun r = p.createRun();
        r.addPicture(is, XWPFDocument.PICTURE_TYPE_PNG, fileName, Units.toEMU(500), Units.toEMU(500));
        try (FileOutputStream out = new FileOutputStream(fileName)) {
            doc.write(out);
        }
    }

}
