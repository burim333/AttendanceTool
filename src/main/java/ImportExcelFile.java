import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.prefs.Preferences;

/**
 *
 * @author Burim Sadriu Copyright© 2020, All Rights Reserved
 */
public class ImportExcelFile extends JFrame {

    JTextField sourceFileTextField;
    String selectedFilePath = "";
    String seletedFileName = "";
    LookAndFeel previousLF;
    Preferences pref = Preferences.userRoot();
    String path = pref.get("DEFAULT_PATH", "");
    JButton copyButton;
    JButton sourceFileChoserButton;
    JLabel sourceFileJLabel;
    
    public ImportExcelFile() throws HeadlessException, ClassNotFoundException, InstantiationException, IllegalAccessException, UnsupportedLookAndFeelException {
        previousLF = UIManager.getLookAndFeel();
        sourceFileJLabel = new JLabel("Fil: ");
        sourceFileTextField = new JTextField(20);
        sourceFileTextField.setEditable(false);

        sourceFileChoserButton = new JButton("Välj fil");
        sourceFileChoserButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                try {
                    UIManager.setLookAndFeel("com.sun.java.swing.plaf.windows.WindowsLookAndFeel");
                    JFileChooser jFileChooser = new JFileChooser(path);
                    jFileChooser.setDialogTitle("Importera fil");
                    FileNameExtensionFilter xmlfilter = new FileNameExtensionFilter("xlsx, xls, xlsm", "xlsx", "xls", "xlsm");//bara xlsx och xls filer
                    jFileChooser.setFileFilter(xmlfilter);
                    int status = jFileChooser.showOpenDialog(ImportExcelFile.this);
                    if (status == JFileChooser.APPROVE_OPTION) {
                        seletedFileName = jFileChooser.getSelectedFile().getName();
                        selectedFilePath = jFileChooser.getSelectedFile().getAbsolutePath();
                        sourceFileTextField.setText(seletedFileName);
                        UIManager.setLookAndFeel(previousLF);
                        UIManager.getLookAndFeelDefaults().put("TableHeader.font", new Font("Arial", Font.PLAIN, 14));// ändra text i kategorin
                    } else {
                        UIManager.setLookAndFeel(previousLF);
                        UIManager.getLookAndFeelDefaults().put("TableHeader.font", new Font("Arial", Font.PLAIN, 14));// ändra text i kategorin
                    }
                } catch (ClassNotFoundException | InstantiationException | IllegalAccessException | UnsupportedLookAndFeelException ex) {
                    JOptionPane.showMessageDialog(null, "Ett fel inträffade (fel: välj fil)\n"
                            + "Det gick inte att kopiera filen!", "Meddelande", JOptionPane.DEFAULT_OPTION);
                }
            }
        });

        copyButton = new JButton("Kopiera");
        copyButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                if (FilenameUtils.getExtension(selectedFilePath).equals("xlsx")
                        || FilenameUtils.getExtension(selectedFilePath).equals("xlsm")) {
                    File destinationFile = new File((System.getProperty("user.home") + "\\Desktop\\Attendance Tool\\Dokument\\deltagare.xls"));
                    try {
                        File sourceFile = new File(selectedFilePath);
                        String absoluteSourcePath = sourceFile.getAbsolutePath();
                        String sourceFilePath = absoluteSourcePath.substring(0, absoluteSourcePath.lastIndexOf(File.separator));
                        if (!destinationFile.exists() && !destinationFile.isDirectory()) {
                            XLS_XMtoXLS convert = new XLS_XMtoXLS(sourceFile);//fungerar för xlsm också
                            pref.put("DEFAULT_PATH", sourceFilePath);
                            try {
                                dispose();
                                convert.xlsx2xls_progress(selectedFilePath);

                            } catch (InvalidFormatException | GeneralSecurityException ex) {
                                Logger.getLogger(ImportExcelFile.class.getName()).log(Level.SEVERE, null, ex);
                            }

                        } else {
                            int reply = JOptionPane.showConfirmDialog(null, "Samma dokument finns redan, vill du ersätta det?\n\n"
                                    + "Tips: Innan du ersätter dokumentet är det bra att spara\n"
                                    + "          det först ('Spara och nollställ Excelfil'- knappen).\n\n", "Säkerhetsfråga", JOptionPane.YES_NO_OPTION, JOptionPane.PLAIN_MESSAGE);
                            if (reply == JOptionPane.YES_OPTION) {
                                pref.put("DEFAULT_PATH", sourceFilePath);
                                destinationFile.delete();
                                XLS_XMtoXLS convert = new XLS_XMtoXLS(sourceFile);
                                try {
                                    dispose();
                                    convert.xlsx2xls_progress(selectedFilePath);

                                } catch (InvalidFormatException | GeneralSecurityException ex) {
                                    JOptionPane.showMessageDialog(null, "Det gick inte att importera filen\n"
                                            + "Kontakta ansvarig", "Meddelande", JOptionPane.DEFAULT_OPTION);
                                }
                                try {
                                    UIManager.setLookAndFeel(previousLF);
                                    UIManager.getLookAndFeelDefaults().put("TableHeader.font", new Font("Arial", Font.PLAIN, 14));// ändra text i kategorin
                                } catch (UnsupportedLookAndFeelException ex) {
                                    JOptionPane.showMessageDialog(null, "Fel: LookAndFeel\n"
                                            + "Kontakta ansvarig", "Meddelande", JOptionPane.DEFAULT_OPTION);
                                }
                            } else {
                                try {
                                    UIManager.setLookAndFeel(previousLF);
                                    UIManager.getLookAndFeelDefaults().put("TableHeader.font", new Font("Arial", Font.PLAIN, 14));// ändra text i kategorin
                                } catch (UnsupportedLookAndFeelException ex) {
                                    JOptionPane.showMessageDialog(null, "Fel: LookAndFeel\n"
                                            + "Kontakta ansvarig", "Meddelande", JOptionPane.DEFAULT_OPTION);
                                }
                                dispose();
                            }
                        }
                    } catch (IOException ex) {
                        JOptionPane.showMessageDialog(null, "Det gick inte att importera filen\n\n"
                                + "Följande kan ha inträffat:\n\n"
                                + "1. Excelfilen som du försöker importera används av någon annan.\n\n"
                                + "2. Din nuvarande excelfil är öppen.\n\n"
                                + "3. Du saknar internetuppkoppling.\n\n"
                                + "4. Om ingen av ovanstående stämmer, logga in på servern och försök\n"
                                + "     att importera excelfilen igen.\n\n.", "Meddelande", JOptionPane.DEFAULT_OPTION);
                    }
                } else if (FilenameUtils.getExtension(selectedFilePath).equals("xls")) {
                    File sourceFile = new File(selectedFilePath);
                    File destinationFolder = new File((System.getProperty("user.home") + "\\Desktop\\Attendance Tool\\Dokument"));
                    String absoluteSourcePath = sourceFile.getAbsolutePath();
                    String sourceFilePath = absoluteSourcePath.substring(0, absoluteSourcePath.lastIndexOf(File.separator));
                    try {
                        File f = new File((System.getProperty("user.home") + "\\Desktop\\Attendance Tool\\Dokument\\deltagare.xls"));
                        File fToRename = new File((System.getProperty("user.home") + "\\Desktop\\Attendance Tool\\Dokument\\" + seletedFileName));
                        if (!f.exists() && !f.isDirectory()) {
                            //System.out.println("dokument finns INTE");
                            if (!sourceFilePath.equals(destinationFolder.getAbsolutePath())) {
                                FileUtils.copyFileToDirectory(sourceFile, destinationFolder, false);
                            }
                            if (!fToRename.getName().equals("deltagare.xls")) {
                                fToRename.renameTo(new File((System.getProperty("user.home") + "\\Desktop\\Attendance Tool\\Dokument\\deltagare.xls")));
                            }
                            Desktop dt = Desktop.getDesktop();
                            pref.put("DEFAULT_PATH", sourceFilePath);
                            dispose();
                            showMessage();
                        } else {
                            int reply;
                            //System.out.println("dokument finns");
                            reply = JOptionPane.showConfirmDialog(null, "Samma dokument finns redan, vill du ersätta det?\n\n"
                                    + "Tips: Innan du ersätter dokumentet är det bra att spara\n"
                                    + "          det först ('Spara och nollställ Excelfil'- knappen).\n\n", "Säkerhetsfråga", JOptionPane.YES_NO_OPTION, JOptionPane.PLAIN_MESSAGE);

                            if (reply == JOptionPane.YES_OPTION) {
                                if (!sourceFilePath.equals(destinationFolder.getAbsolutePath())) {
                                    f.delete();
                                    FileUtils.copyFileToDirectory(sourceFile, destinationFolder, false);
                                    fToRename.renameTo(new File((System.getProperty("user.home") + "\\Desktop\\Attendance Tool\\Dokument\\deltagare.xls")));
                                }
                                if (!fToRename.getAbsolutePath().equals(f.getAbsolutePath()) && sourceFilePath.equals(destinationFolder.getAbsolutePath())) {
                                    f.delete();
                                    fToRename.renameTo(new File((System.getProperty("user.home") + "\\Desktop\\Attendance Tool\\Dokument\\deltagare.xls")));
                                }
                                Desktop dt = Desktop.getDesktop();
                                pref.put("DEFAULT_PATH", sourceFilePath);
                                dispose();
                                showMessage();
                                try {
                                    UIManager.setLookAndFeel(previousLF);
                                    UIManager.getLookAndFeelDefaults().put("TableHeader.font", new Font("Arial", Font.PLAIN, 14));// ändra text i kategorin
                                } catch (UnsupportedLookAndFeelException ex) {
                                    JOptionPane.showMessageDialog(null, "Fel: LookAndFeel\n"
                                            + "Kontakta ansvarig", "Meddelande", JOptionPane.DEFAULT_OPTION);
                                }
                            } else {
                                try {
                                    UIManager.setLookAndFeel(previousLF);
                                    UIManager.getLookAndFeelDefaults().put("TableHeader.font", new Font("Arial", Font.PLAIN, 14));// ändra text i kategorin
                                } catch (UnsupportedLookAndFeelException ex) {
                                    JOptionPane.showMessageDialog(null, "Fel: LookAndFeel\n"
                                            + "Kontakta ansvarig", "Meddelande", JOptionPane.DEFAULT_OPTION);
                                }
                                dispose();
                            }
                        }
                    } catch (IOException ex) {
                        JOptionPane.showMessageDialog(null, "Det gick inte att kopiera filen. Den kan vara öppen", "Meddelande", JOptionPane.DEFAULT_OPTION);
                    }
                } else if (!sourceFileTextField.getText().isEmpty()) {
                    JOptionPane.showMessageDialog(null, "Se till att filen är en excelfil och att den är sparad i ett av dessa filformat:\n\n"
                            + "Excel 97-2003 Workbook(*.xls)\n"
                            + "Excel-arbetsbok (*.xlsx)\n"
                            + "Excel Macro-Enabled Workbook(*.xlsm)\n\n"
                            + "Försök sedan att importera filen igen.\n\n", "Meddelande", JOptionPane.DEFAULT_OPTION);
                    File file = new File(System.getProperty("user.home") + "\\Desktop\\Attendance Tool\\Dokument\\deltagare.xls");
                    if (file.exists()) {
                        file.delete();// ta bort deltagare.xls då den inte kan öppnas av programmet.
                    }
                } else {
                    JOptionPane.showMessageDialog(null, "Du måste välja fil först", "Meddelande", JOptionPane.DEFAULT_OPTION);
                }
            }
        });

        setLayout(new FlowLayout(FlowLayout.TRAILING));
        add(sourceFileJLabel);
        add(sourceFileTextField);
        add(sourceFileChoserButton);
        add(copyButton);
    }

    public void showMessage() {
        //JOptionPane.showMessageDialog(null, "Starta om programmet", "Meddelande", JOptionPane.DEFAULT_OPTION);
        GUI.destroyGUI();//stäng GUI
        GUI.createAndShowUI();//starta om UI
        //System.exit(0);
    }
}
