import com.itextpdf.text.Document;
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
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Calendar;
import java.util.Date;
import javax.swing.JOptionPane;

/**
 *
 * @author Burim Sadriu Copyright© 2020, All Rights Reserved
 */
public class PresenceReportPDF {

    private static PdfImportedPage page;
    private static int year;
    private static int month;
    private static int day;
    private static String stringMonth = "fel";
    private static String stringYear = "fel";
    private static String activity = GUI.prefs.get("aktivitetString", " ");
    private static String name = "fel";
    private static String eAddress = GUI.prefs.get("mailString", " ");
    private static String phone = GUI.prefs.get("telString", " ");
    private static String organizer = GUI.prefs.get("anordString", " ");
    //static String absence = "4";
    private static String pathPDF = System.getProperty("user.home") + "/Skrivbord/Attendance Tool/Dokument/monthreport.pdf";
    private static String pathNewPDF = " ";
    //static String fontPath;
    private static int index = 0;
    private static PdfContentByte cb;
    private static PdfContentByte cb2;
    static BaseFont f_cn;
    static Document document;
    static PdfWriter writer;
    private static InputStream templateInputStream;
    static PdfReader reader;
    private static int[] x_pos;

    public PresenceReportPDF() {
        //fontPath = getClass().getResource("/documents/Caviardreams.ttf").toString();
        System.out.println("PresencePDF: " + pathPDF);

        eAddress = GUI.prefs.get("mailString", " ");
        phone = GUI.prefs.get("telString", " ");
        organizer = GUI.prefs.get("anordString", " ");
        activity = GUI.prefs.get("aktivitetString", " ");

        x_pos = new int[32];
        //x positionen för varje dag, i dag-rutorna i pdf:et
        x_pos[1] = 35;
        x_pos[2] = 60;
        x_pos[3] = 85;
        x_pos[4] = 110;
        x_pos[5] = 135;
        x_pos[6] = 160;
        x_pos[7] = 185;
        x_pos[8] = 210;
        x_pos[9] = 235;
        x_pos[10] = 260;
        x_pos[11] = 285;
        x_pos[12] = 310;
        x_pos[13] = 335;
        x_pos[14] = 360;
        x_pos[15] = 385;
        x_pos[16] = 410;
        x_pos[17] = 435;
        x_pos[18] = 460;
        x_pos[19] = 485;
        x_pos[20] = 510;
        x_pos[21] = 535;
        x_pos[22] = 560;
        x_pos[23] = 585;
        x_pos[24] = 610;
        x_pos[25] = 635;
        x_pos[26] = 660;
        x_pos[27] = 685;
        x_pos[28] = 710;
        x_pos[29] = 735;
        x_pos[30] = 760;
        x_pos[31] = 785;

    }

    //skapa en kopia av original pdf:et och spara det i aktuell grupp med deltagarens namn
    public static void initializePDF(String fileName, String group) throws FileNotFoundException, IOException {
        PdfReader.unethicalreading = true;
        eAddress = GUI.prefs.get("mailString", " ");
        phone = GUI.prefs.get("telString", " ");
        organizer = GUI.prefs.get("anordString", " ");
        activity = GUI.prefs.get("aktivitetString", " ");
        fileName = fileName.replace("/", "").replace("\\", "").replace("!", "").replace("?", "").replace("*", "").replace(":", "").replace("?", "").replace("<", "").replace(">", "").replace("|", "").trim();

        try {
            if (!Files.isDirectory(Paths.get(System.getProperty("user.home") + "/Skrivbord/Attendance Tool/Månadsrapporter/" + group))) {
                File dir = new File(System.getProperty("user.home") + "/Skrivbord/Attendance Tool/Månadsrapporter/" + group);
                dir.mkdir();
            }
            pathNewPDF = System.getProperty("user.home") + "/Skrivbord/Attendance Tool/Månadsrapporter/" + group + "/" + fileName + ".pdf";

            document = new Document(PageSize.A4.rotate());//PageSize.A4 för portrait
            writer = PdfWriter.getInstance(document, new FileOutputStream(pathNewPDF));
            document.open();
            cb = writer.getDirectContent();
            f_cn = BaseFont.createFont(BaseFont.COURIER, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            // Load existing PDF
            templateInputStream = new FileInputStream(pathPDF);
            reader = new PdfReader(templateInputStream);

        } catch (Exception ex) {
            JOptionPane.showMessageDialog(null, "1. Äldre månadsrapport för " + fileName + " får inte vara öppen.\n"
                    + "2. Deltagarens namn, " + fileName + ", får inte innehålla speciella tecken (t.ex. '!' ?' '*'  '/'  '\'  :).\n"
                    + "3. Säkerställ att PDF-dokumentet (monthreport.pdf) finns i ’Attendance Tool’ → ’Dokument’.\n"
                    + "4. Om ovanstående inte hjälper, starta om datorn och försök igen.\n\n", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
    }

    //Lägg till info om vem pdf:et kommer från
    public static void setStaticInfo(String deltNamn, String raportMonth, String raportYear, String group) {
        name = deltNamn.replace("/", "").replace("\\", "").replace("*", "").replace(":", "").replace("?", "").replace("<", "").replace(">", "").replace("|", "").trim();;
        Date date = new Date();
        Calendar cal = Calendar.getInstance();
        cal.setTime(date);
        year = cal.get(Calendar.YEAR);
        month = cal.get(Calendar.MONTH) + 1;//månader startar från noll!
        day = cal.get(Calendar.DAY_OF_MONTH);
        stringYear = Integer.toString(year);

        //lägg till en nolla framför för jan - sep.
        if (month > 0 && month < 10) {
            stringMonth = "0" + Integer.toString(month);
        } else {
            stringMonth = Integer.toString(month);
        }
        int n = reader.getNumberOfPages();
        //i är antalet sidor i original pdf:et
        for (int i = 1; i <= n; i++) {
            document.newPage();
            page = writer.getImportedPage(reader, i);
            cb.addTemplate(page, 0, i);
            //lägg till namn, år och datum på första sidan också
            if (i == 1) {
                cb.beginText();
                cb.setFontAndSize(f_cn, 11);

                //Namn
                cb.setTextMatrix(74, 398); //(x-pos, y-pos)
                cb.showText(name);

                //grupp
                cb.setTextMatrix(74, 485); //(x-pos, y-pos)
                cb.showText("Grupp: " + group);

                //År
                //System.out.println(raportYear +", "+stringYear);
                cb.setTextMatrix(671, 398);
                if (!raportYear.equals(" ")) {
                    cb.showText(raportYear);
                } else {
                    cb.showText(stringYear);
                }
                //System.out.println(raportMonth +", "+stringMonth);
                //månad
                cb.setTextMatrix(731, 398);
                if (!raportMonth.equals(" ")) {
                    cb.showText(raportMonth);
                } else {
                    cb.showText(stringMonth);
                }
                cb.endText();
            } //Lägg till företags-information på sida nr 2 (gamla månadsrapporten).
            //Används inte nu.
            else if (i == 2) {
                cb.beginText();
                cb.setFontAndSize(f_cn, 11);
                cb.setTextMatrix(60, 686);
                cb.showText(name);
                cb.setTextMatrix(390, 708);
                if (!raportYear.equals(" ")) {
                    cb.showText(raportYear);
                } else {
                    cb.showText(stringYear);
                }

                cb.setTextMatrix(465, 708);
                // om man valt en annan månad än det som är nu
                if (!raportMonth.equals(" ")) {
                    cb.showText(raportMonth);// annan månad
                } else {
                    cb.showText(stringMonth);
                }

                cb.setTextMatrix(71, 605);
                cb.showText(activity.toUpperCase());
                cb.setTextMatrix(71, 482);
                cb.showText(organizer);
                cb.setTextMatrix(71, 458);
                cb.showText(eAddress);
                cb.setTextMatrix(320, 481);
                cb.showText(phone);
                cb.endText();
            }
        }
    }

    //lägg till frånvaro i pdf:et, beroende på exceldata och deltagare
    public static void addAbsenceToPDF(String rowValue, int index2, String group, String tid) {
        rowValue = rowValue(rowValue, group, tid);
        cb.beginText();
        cb.setTextMatrix(getXPosition(index2), 282);
        cb.showText(rowValue);
        cb.endText();
    }

    //stäng dokumentet och öppna ett fönster av den nya pdf filen
    public void openPDF() throws IOException {
        document.close();

        //öppna pdf filen 
        Desktop desktop = Desktop.getDesktop();
        File file = new File(pathNewPDF);
        //desktop.open(file);
    }

    //hämta data som ska fyllas i pdf:et
    public static String rowValue(String columnValue, String group, String schedule) {
        schedule = schedule.replaceAll("\\s+", "").toLowerCase();
        columnValue = columnValue.toLowerCase();
        if (columnValue.trim().equals("h")) {//hel frånvaro för heltidarna
            return "8";//h = heltid
        } else if (columnValue.trim().equals("d")) {//hel frånvaro för deltidarna
            return "4";//d = deltid
        } else if (columnValue.equals("0.15") || columnValue.equals("0.30") || columnValue.equals("0.45")
                || columnValue.equals("1") || columnValue.equals("1.15") || columnValue.equals("1.30")
                || columnValue.equals("1.45") || columnValue.equals("2") || columnValue.equals("2.15")
                || columnValue.equals("2.30") || columnValue.equals("2.45") || columnValue.equals("3")
                || columnValue.equals("3.15") || columnValue.equals("3.30") || columnValue.equals("3.45")
                || columnValue.equals("4") || columnValue.equals("4.15") || columnValue.equals("4.30")
                || columnValue.equals("4.45") || columnValue.equals("5") || columnValue.equals("5.15")
                || columnValue.equals("5.30") || columnValue.equals("5.45") || columnValue.equals("6")
                || columnValue.equals("6.15") || columnValue.equals("6.30") || columnValue.equals("6.45")
                || columnValue.equals("7") || columnValue.equals("7.15") || columnValue.equals("7.30")
                || columnValue.equals("7.45") || columnValue.equals("8") || columnValue.equals("s")
                || columnValue.equals("v") || columnValue.equals("ö") || columnValue.equals("a")
                || columnValue.equals("af") || columnValue.equals("p")) {
            return columnValue;//minus frånvarotimmar      
        } else if ((columnValue.trim().equals("so")) && (schedule.trim().equals("deltid")
                || schedule.trim().equals("em/fm") || schedule.trim().equals("fm/em")
                || schedule.trim().equals("fm") || schedule.trim().equals("em"))) {
            return "4";//deltid
        } else if ((columnValue.trim().equals("so")) && schedule.trim().equals("heltid")) {
            return "8";//heltid
        }
        return " ";
    }

    //hämta positionen av rutorna där data från excelen ska fyllas i.
    public static int getXPosition(int day) {
        return x_pos[day];
    }
}
