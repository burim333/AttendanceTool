package PDF_classes;

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfImportedPage;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfWriter;
import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import javax.swing.JOptionPane;

/**
 *
 * @author Burim Sadriu Copyright© 2020, All Rights Reserved
 */
public class PeriodicReportPDF {

    private static PdfImportedPage page;
    private static int year;
    private static int month;
    private static int day;
    private static String name = "*";
    private static String publicStartDate = "ÅÅÅÅ-MM-DD";
    private static String absence = "4";
    private static String pathPDF = System.getProperty("user.home") + "/Skrivbord/Attendance Tool/Dokument/absencePDF.pdf";
    private static String pathNewPDF = " ";
    private static String fontPath;
    private static int index = 0;
    private static PdfContentByte cb;
    private static PdfContentByte cb2;
    public static BaseFont f_cn;
    public static Document document;
    public static PdfWriter writer;
    private static InputStream templateInputStream;
    public static PdfReader reader;
    private String dag;
    private static int[] x_pos;
    private static int[] y_pos;
    private boolean create_pdf;

    public PeriodicReportPDF() {
        //fontPath = getClass().getResource("/documents/Caviardreams.ttf").toString();
        System.out.println("periodicReport: "+ pathPDF);
        x_pos = new int[32];
        y_pos = new int[61];

        x_pos[1] = 120;
        x_pos[2] = 130;
        x_pos[3] = 140;
        x_pos[4] = 150;
        x_pos[5] = 160;
        x_pos[6] = 170;
        x_pos[7] = 180;
        x_pos[8] = 190;
        x_pos[9] = 200;
        x_pos[10] = 210;
        x_pos[11] = 220;
        x_pos[12] = 230;
        x_pos[13] = 240;
        x_pos[14] = 250;
        x_pos[15] = 260;
        x_pos[16] = 270;
        x_pos[17] = 280;
        x_pos[18] = 290;
        x_pos[19] = 300;
        x_pos[20] = 310;
        x_pos[21] = 320;
        x_pos[22] = 330;
        x_pos[23] = 340;
        x_pos[24] = 350;
        x_pos[25] = 360;
        x_pos[26] = 370;
        x_pos[27] = 380;
        x_pos[28] = 390;
        x_pos[29] = 400;
        x_pos[30] = 410;
        x_pos[31] = 420;

        y_pos[0] = 690;
        y_pos[1] = 668;
        y_pos[2] = 648;
        y_pos[3] = 626;
        y_pos[4] = 606;
        y_pos[5] = 584;
        y_pos[6] = 562;
        y_pos[7] = 541;
        y_pos[8] = 520;
        y_pos[9] = 500;
        y_pos[10] = 479;
        y_pos[11] = 459;
        y_pos[12] = 438;
        y_pos[13] = 418;
        y_pos[14] = 397;
        y_pos[15] = 377;
        y_pos[16] = 357;
        y_pos[17] = 336;
        y_pos[18] = 315;
        y_pos[19] = 295;
        y_pos[20] = 274;
        y_pos[21] = 253;
        y_pos[22] = 232;
        y_pos[23] = 211;
        y_pos[24] = 191;
        y_pos[25] = 170;
        y_pos[26] = 150;
        y_pos[27] = 130;
        y_pos[28] = 109;
        y_pos[29] = 90;
        //andra sidan
        y_pos[30] = 687;
        y_pos[31] = 666;
        y_pos[32] = 646;
        y_pos[33] = 625;
        y_pos[34] = 606;
        y_pos[35] = 584;
        y_pos[36] = 562;
        y_pos[37] = 541;
        y_pos[38] = 520;
        y_pos[39] = 500;
        y_pos[40] = 479;
        y_pos[41] = 459;
        y_pos[42] = 438;
        y_pos[43] = 418;
        y_pos[44] = 397;
        y_pos[45] = 377;
        y_pos[46] = 357;
        y_pos[47] = 336;
        y_pos[48] = 315;
        y_pos[49] = 295;
        y_pos[50] = 274;
        y_pos[51] = 253;
        y_pos[52] = 232;
        y_pos[53] = 211;
        y_pos[54] = 191;
        y_pos[55] = 170;
        y_pos[56] = 150;
        y_pos[57] = 130;
        y_pos[58] = 109;
        y_pos[59] = 90;

    }

    //skapa en kopia av original pdf:et och spara det i vald grupp med deltagarens namn.
    public boolean createPDF(String fileName) throws FileNotFoundException, IOException {

        try {
            fileName = fileName.replace("/", "").replace("\\", "").replace("!", "").replace("?", "").replace("*", "").replace(":", "").replace("?", "").replace("<", "").replace(">", "").replace("|", "").trim();
            //fontPath = getClass().getResource("/documents/Caviardreams.ttf").toString();
            f_cn = BaseFont.createFont(fontPath, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);

            pathNewPDF = System.getProperty("user.home") + "/Skrivbord/Attendance Tool/Periodiska rapporter/Periodisk rapport_" + fileName + ".pdf";

            document = new Document(PageSize.A4);
            writer = PdfWriter.getInstance(document, new FileOutputStream(pathNewPDF));
            document.open();
            cb = writer.getDirectContent();

            // Load existing PDF
            templateInputStream = new FileInputStream(pathPDF);
            reader = new PdfReader(templateInputStream);

            document.newPage();
            page = writer.getImportedPage(reader, 1);
            cb.addTemplate(page, 0, 1);

        } catch (Exception ex) {
            JOptionPane.showMessageDialog(null, "Periodiska rapporter kunde inte skapas!\n\n"
                    + "1. Startdatumet (" + fileName + ") får inte innehålla speciella tecken (t.ex. '!' ?' '*'  '/'  '\'  :).\n"
                    + "2. Äldre periodisk rapport ( Periodisk rapport_" + fileName + " ) får inte vara öppen.\n"
                    + "3. Säkerställ att PDF-dokumentet ’absencePDF.pdf’ finns i ’Attendance Tool’ → ’Dokument’.\n"
                    + "4. Om ovanstående inte hjälper, starta om datorn och försök igen.\n\n\n", "Meddelande", JOptionPane.DEFAULT_OPTION);
            return false;
        }
        return true;
    }

    //Lägg till info om vem pdf:et kommer från
    public boolean setStaticInfo(String startDate) {
        publicStartDate = startDate;
        //create_pdf = true;
        try {
            create_pdf = createPDF(startDate);

        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "Ett fel inträffade när programmet försökte\n"
                    + "skapa periodiska rapporter!\n", "Meddelande", JOptionPane.DEFAULT_OPTION);
            return false;
        }

        cb.beginText();
        cb.setFontAndSize(f_cn, 9);

        //Startdatum
        cb.setTextMatrix(130, 748); //(x-pos, y-pos)
        cb.showText(startDate);
        cb.endText();
        return create_pdf;
    }

    public void addNameToPDF(String namn, int row) {
        if (namn.length() > 22) {
            namn = namn.substring(0, Math.min(namn.length(), 22));
        }
        if (row == 30) {
            document.newPage();
            page = writer.getImportedPage(reader, 2);
            cb.addTemplate(page, 0, 2);
            cb.beginText();
            cb.setFontAndSize(f_cn, 9);
            cb.setTextMatrix(130, 749); //(x-pos, y-pos)
            cb.showText(publicStartDate);
            cb.endText();
        }
        cb.beginText();
        cb.setFontAndSize(f_cn, 7);
        cb.setTextMatrix(21, getYPosition(row));
        cb.showText(namn);
        cb.endText();
    }

    //lägg till frånvaro i pdf:et, beroende på excel data och deltagare
    public void addAbsenceToPDF(String rowValue, int day, int indexY, int indexX, int indexXp) throws DocumentException, IOException {
        dag = String.valueOf(day);
        cb.beginText();
        cb.setFontAndSize(f_cn, 6);
        if (rowValue.toLowerCase().equals("p")) {
            if (indexXp != 0) { //lägg inte till en komma i början
                cb.setTextMatrix(getXPosition(indexXp) + 242, getYPosition(indexY));
                cb.showText(",");
            }
            cb.setTextMatrix(getXPosition(indexXp) + 234, getYPosition(indexY));
            cb.showText(dag);
            cb.endText();
        } else {
            if (indexX != 0) {
                cb.setTextMatrix(getXPosition(indexX) + 8, getYPosition(indexY));
                cb.showText(",");
            }
            cb.setTextMatrix(getXPosition(indexX), getYPosition(indexY));
            cb.showText(dag);
            cb.endText();
        }
    }

    public void closeAbsencePDF() {
        document.close();
    }

    //stäng dokumentet och öppna den nya pdf filen
    public void openAbsencePDF() throws IOException {
        //öppna pdf filen 
        Desktop desktop = Desktop.getDesktop();
        File file = new File(pathNewPDF);
        //desktop.open(file);
    }

    //hämta positionen av rutorna där data från Excelen ska fyllas i.
    public static int getXPosition(int day) {
        //x positionen för varje dag, i dag-rutorna i pdf:et
        return x_pos[day];
    }

    public static int getYPosition(int row) {
        //Första sidan
        if (row > 59) {
            return y_pos[59];
        } else {
            return y_pos[row];
        }
    }
}
