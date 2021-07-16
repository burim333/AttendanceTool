import java.awt.Desktop;
import java.io.File;
import javax.swing.JOptionPane;

/**
 *
 * @author Burim Sadriu Copyright© 2020, All Rights Reserved
 */
public class OpenFoldersAndDoc {

    static String groupsPath = "S:\\330 Region Kristianstad-Blekinge\\Kristianstad Björkhem\\Grupplistor";// Björkhem
    static String schedulePath = "S:\\330 Region Kristianstad-Blekinge\\Hässleholm\\Schema\\Veckoschema Hässleholm.xlsx";
    static String modulePath = "S:\\AlphaCE Alla\\Aktuella uppdrag\\Grundläggande moduler\\Utbildningsmaterial";
    static String modulePath2 = "S:\\330 Region Kristianstad-Blekinge\\Kristianstad\\GM Modul 2\\Modul 2 GM\\Modul 2 - förberedande undervisning i svenska språket";
    static String modulePath3 = "S:\\330 Region Kristianstad-Blekinge\\Kristianstad\\GM Modul 2\\Eget material MODUL 2";
    static String deviationPath = "S:\\330 Region Kristianstad-Blekinge\\Hässleholm\\Admin\\GM\\Avvikelserapporter";
    static String dagJournalPath = "S:\\330 Region Kristianstad-Blekinge\\Hässleholm\\Dagjournal.xlsx";
    static String help = System.getProperty("user.home") + "\\Desktop\\Attendance Tool\\Hjälp";
    static String monthReportPath = System.getProperty("user.home") + "\\Desktop\\Attendance Tool\\Manadsrapporter";
    static String openPeriodRap = System.getProperty("user.home") + "\\Desktop\\Attendance Tool\\Periodiska rapporter";
    static String openMonthRap = System.getProperty("user.home") + "\\Desktop\\Attendance Tool\\Manadsrapporter";
    static String openAbsenceLists = System.getProperty("user.home") + "\\Desktop\\Attendance Tool\\Dagens grupplistor";

    public OpenFoldersAndDoc() {

    }

    public void moduler() {
        try {
            Desktop.getDesktop().open(new File(modulePath3));//modul 2 nya
            Desktop.getDesktop().open(new File(modulePath));//utbildningsmaterial
            //Desktop.getDesktop().open(new File(modulePath2));//modul 2  
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(null, "1. Säkerställ att du har internetuppkoppling.\n"
                    + "2. Säkerställ att du är inloggad på servern.\n"
                    + "3. Kontrollera mappen 'Utbildningsmaterial' i (S:).", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
    }

    public void openReportsFolder() {
        try {
            Desktop.getDesktop().open(new File(monthReportPath));
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(null, "Kontrollera mappen 'Attendance Tool' på skrivbordet.\n", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
    }

    public void openPeriodReportsFolder() {
        try {
            Desktop.getDesktop().open(new File(openPeriodRap));
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(null, "Kontrollera mappen 'Attendance Tool' på skrivbordet.\n", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
    }

    public void openMonthReportsFolder() {
        try {
            Desktop.getDesktop().open(new File(openMonthRap));
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(null, "Kontrollera mappen 'Attendance Tool' på skrivbordet.\n", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
    }

    public void openAbsenceListFolder() {
        try {
            Desktop.getDesktop().open(new File(openAbsenceLists));
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(null, "Kontrollera mappen 'Attendance Tool' på skrivbordet.\n", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
    }

    public void schedule() {
        try {
            Desktop.getDesktop().open(new File(schedulePath));
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(null, "1. Säkerställ att du har internetuppkoppling.\n"
                    + "2. Säkerställ att du är inloggad på servern.\n"
                    + "3. Kontrollera mappen 'Schema' i (S:).\n"
                    + "4. Kolla i mappen om filen 'Veckoschema Hässleholm.xlsx' finns.", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
    }

    public void deviationReport() {
        try {
            Desktop.getDesktop().open(new File(deviationPath));
        } catch (Exception ex) {

            JOptionPane.showMessageDialog(null, "1. Säkerställ att du har internetuppkoppling.\n"
                    + "2. Säkerställ att du är inloggad på servern.\n"
                    + "3. Kontrollera mappen 'Avvikelserapporter' i (S:).", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
    }

    public void groups() {
        try {
            Desktop.getDesktop().open(new File(groupsPath));
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(null, "1. Säkerställ att du har internetuppkoppling.\n"
                    + "2. Säkerställ att du är inloggad på servern.\n"
                    + "3. Kontrollera mappen 'Grupper' i (S:).", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
    }

    public void dagJournal() {
        try {
            Desktop.getDesktop().open(new File(dagJournalPath));
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(null, "1. Säkerställ att du har internetuppkoppling.\n"
                    + "2. Säkerställ att du är inloggad på servern.\n"
                    + "3. Kontrollera mappen 'Hässleholm' i (S:\\330 Region Kristianstad-Blekinge).\n"
                    + "4. Kolla i mappen om filen 'Dagjournal.xlsx' finns.", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
    }

    public void help() {
        try {
            Desktop.getDesktop().open(new File(help));
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(null, "Säkerställ att mappen 'Hjälp' finns i\n"
                    + "mappen 'Attendance Tool' som finns på skrivbordet\n", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
    }
}
