import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.util.Locale;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JOptionPane;

import org.joda.time.DateTime;
import org.joda.time.DateTimeZone;

/**
 * @author Burim Sadriu Copyright© 2020, All Rights Reserved
 */
//Den här classen kopierar excelfilen när man är klar med närvaro-inskrivningen för månaden.
public final class CopyFiles {

    private FileInputStream instream = null;
    private FileOutputStream outstream = null;
    private ExcelHandler excel = new ExcelHandler();

    public CopyFiles(String month) {
        checkForCopy(month);
    }

    public static String theMonth(int month) {
        String[] monthNames = {"januari", "februari", "mars", "april", "maj", "juni", "juli", "augusti", "september", "oktober", "november", "december"};
        return monthNames[month];
    }

    public void checkForCopy(String month) {

        DateTimeZone timeZone = DateTimeZone.forID("Europe/Stockholm");
        DateTime dateTime = new DateTime(timeZone);
        Locale localeSweden = new Locale("sv", "SE"); // ( language code, country code );
        String monthString = dateTime.monthOfYear().getAsText(localeSweden);
        monthString = monthString.toLowerCase();

        File infile = new File(System.getProperty("user.home") + "\\Desktop\\Attendance Tool\\Dokument\\deltagare.xls");
        File outfile = new File(System.getProperty("user.home") + "\\Desktop\\Attendance Tool\\Dokument\\Tidigare Excel\\" + monthString + "-deltagare.xls");

        int monthInt;
        String theMonth = " ";

        if (!month.equals(" ")) {
            monthInt = Integer.parseInt(month) - 1;
            theMonth = theMonth(monthInt);
        }

        if (!month.equals(" ")) {
            outfile = new File(System.getProperty("user.home") + "\\Desktop\\Attendance Tool\\Dokument\\Tidigare Excel\\" + theMonth + "-deltagare.xls");
        }

        if (infile.exists()) {
            int reply = 0;
            if (!month.equals(" ")) {
                reply = JOptionPane.showConfirmDialog(null, "Är du klar med månads-närvaron för " + theMonth + " månad?", "Spara och nollställ Excelfil", JOptionPane.YES_NO_OPTION, JOptionPane.PLAIN_MESSAGE);
            } else {
                reply = JOptionPane.showConfirmDialog(null, "Är du klar med månads-närvaron för " + monthString + " månad?", "Spara och nollställ Excelfil", JOptionPane.YES_NO_OPTION, JOptionPane.PLAIN_MESSAGE);
            }
            if (reply == JOptionPane.YES_OPTION) {

                if (!outfile.exists()) {
                    try {
                        copyFile(infile, outfile, month);
                        excel.clearExcelData();
                    } catch (IOException | ParseException ex) {
                        Logger.getLogger(CopyFiles.class.getName()).log(Level.SEVERE, null, ex);
                    }
                } else {
                    int reply3 = JOptionPane.showConfirmDialog(null, "En fil med samma namn finns redan, vill du ersätta den?\n\n"
                            + "Tänk på att när månads-närvaron sparas så nollställs\n"
                            + "originalfilen. Det kan innebära att om du skriver över filen\n"
                            + "som redan finns kan du förlora månads-närvaron.\n\n", "Säkerhetsfråga", JOptionPane.YES_NO_OPTION, JOptionPane.PLAIN_MESSAGE);
                    if (reply3 == JOptionPane.YES_OPTION) {
                        copyFile(infile, outfile, month);
                        try {
                            excel.clearExcelData();
                        } catch (IOException | ParseException ex) {
                            Logger.getLogger(CopyFiles.class.getName()).log(Level.SEVERE, null, ex);
                        }
                    } else {
                    }
                }
            }
        } else {
            JOptionPane.showMessageDialog(null, "Excelfilen saknas! Börja med att importera en excelfil.", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
    }

    public void copyFile(File infile, File outfile, String month) {
        try {
            instream = new FileInputStream(infile);
            outstream = new FileOutputStream(outfile);

            byte[] buffer = new byte[1024];

            int length;
            /*copiera från input stream till
             * output stream
             */
            while ((length = instream.read(buffer)) > 0) {
                outstream.write(buffer, 0, length);
            }
            //stäng input/output stream
            instream.close();
            outstream.close();
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "Ett fel inträffade\n"
                    + "Det gick inte att kopiera filen", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
    }
}
