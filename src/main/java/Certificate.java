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
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import javax.swing.JOptionPane;

/**
 *
 * @author Burim Sadriu Copyright© 2020, All Rights Reserved
 */
public class Certificate {

    static PdfImportedPage page;
    static int year;
    static int month;
    static int day;
    static String firstWord = null;
    static String stringMonth = "fel";
    static String stringYear = "fel";
    static String stringDay = "fel";
    //static String activity = GUI.prefs.get("aktivitetString", " ");
    static String name = "fel";
    static String samordnare;
    //static String phone = GUI.prefs.get("telString", " ");
    //static String organizer = GUI.prefs.get("anordString", " ");
    static String absence = "4";
    static String pathPDF = System.getProperty("user.home") + "\\Desktop\\Attendance Tool\\Dokument\\Kursintyg.pdf";
    static String pathNewPDF = " ";
    //static String fontPath;
    static int index = 0;
    static PdfContentByte cb;
    static PdfContentByte cb2;
    static BaseFont f_cn;
    static Document document;
    static PdfWriter writer;
    static InputStream templateInputStream;
    static PdfReader reader;

    public Certificate() {
        //fontPath = getClass().getResource("/documents/Ralewaylight.ttf").toString();
        //System.out.println(fontPath);
        //eAddress = GUI.prefs.get("mailString", " ");
        samordnare = "x y";
        //phone = GUI.prefs.get("telString", " ");
        //organizer = GUI.prefs.get("anordString", " ");
        //activity = GUI.prefs.get("aktivitetString", " ");
    }

    //skapa en kopia av original pdf:et och spara det i vald grupp med deltagarens namn
    public static void initializePDF(String fileName) {
        //eAddress = GUI.prefs.get("mailString", " ");
        samordnare = "x y";
        //phone = GUI.prefs.get("telString", " ");
        //organizer = GUI.prefs.get("anordString", " ");
        //activity = GUI.prefs.get("aktivitetString", " ");
        try {
            pathNewPDF = System.getProperty("user.home") + "\\Desktop\\Attendance Tool\\Intyg\\Kursintyg_" + fileName + ".pdf";

            document = new Document(PageSize.A4);
            writer = PdfWriter.getInstance(document, new FileOutputStream(pathNewPDF));
            document.open();
            cb = writer.getDirectContent();
            //f_cn = BaseFont.createFont(fontPath, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            // Load existing PDF
            templateInputStream = new FileInputStream(pathPDF);
            reader = new PdfReader(templateInputStream);

        } catch (Exception ex) {
            JOptionPane.showMessageDialog(null, "1. Se till att PDF-filen med deltagarens namn inte är öppen.\n"
                    + "2. Säkerställ att 'Kursintyg.pdf' är placerad i ’Attendance Tool’ → ’Dokument’ mappen.\n\n", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
    }

    //Lägg till info om vem pdf:et kommer från
    public static void createCertificate(String deltNamn, String teacher, String group, String aboutPart, String city) throws IOException {
        initializePDF(deltNamn);
        name = deltNamn;

        //dela texten i delar som inte är större än 95 tecken, men dela inte orden.
        String result = " ";
        List<String> matchList = new ArrayList<String>();
        Pattern regex = Pattern.compile(".{1,92}(?:\\s|$)", Pattern.DOTALL);
        Matcher regexMatcher = regex.matcher(aboutPart);
        while (regexMatcher.find()) {
            result = regexMatcher.group().trim();// ta bort tomrum i början
            matchList.add(result);
        }

        int xx = 0;

        if (name.contains(" ")) {
            firstWord = name.substring(0, name.indexOf(" "));
        } else {
            firstWord = name;
        }

        Date date = new Date();
        Calendar cal = Calendar.getInstance();
        cal.setTime(date);
        year = cal.get(Calendar.YEAR);
        month = cal.get(Calendar.MONTH) + 1;//månader startar från noll!
        day = cal.get(Calendar.DAY_OF_MONTH);
        stringYear = Integer.toString(year);
        stringDay = Integer.toString(day);
        int dayToday = Integer.parseInt(stringDay);
        //lägg till en nolla framför för jan - sep.
        if (month > 0 && month < 10) {
            stringMonth = "0" + Integer.toString(month);
        } else {
            stringMonth = Integer.toString(month);
        }
        if (dayToday > 0 && dayToday < 10) {
            stringDay = "0" + stringDay;
        }

        String todayDate = stringYear + "-" + stringMonth + "-" + stringDay;
        //i är antalet sidor i original pdf:et

        document.newPage();
        page = writer.getImportedPage(reader, 1);
        cb.addTemplate(page, 0, 1);

        cb.saveState();
        cb.beginText();
        cb.setFontAndSize(f_cn, 12);

        //deltagare
        cb.setTextMatrix(170, 655); //(x-pos, y-pos)
        cb.showText(name);

        //stad
        cb.setTextMatrix(290, 629);
        cb.showText(city);

        //stad
        cb.setTextMatrix(164, 131);
        cb.showText(city);

        //datum
        cb.setTextMatrix(380, 131);
        cb.showText(todayDate);

        //Text
        cb.setTextMatrix(164, 80);
        cb.showText(samordnare);

        //Lärare
        cb.setTextMatrix(380, 80);
        cb.showText(teacher);

        xx = 300;
        for (String item : matchList) {
            cb.setTextMatrix(48, xx);
            cb.showText(item);
            xx = xx - 15;

        }
        cb.endText();
        cb.restoreState();
        document.close();

        //öppna pdf filen 
        Desktop desktop = Desktop.getDesktop();
        File file = new File(pathNewPDF);
        desktop.open(file);
    }

}
