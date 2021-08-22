package Word_classes;

import Excel_classes.*;

import SQL.*;
import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.sql.SQLException;
import java.time.LocalDateTime;
import java.util.Calendar;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;
import javax.swing.JOptionPane;

import OtherClasses.OpenFoldersAndDoc;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.VerticalAlign;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;

/**
 * @author Burim Sadriu Copyright© 2020, All Rights Reserved
 */
public class AttendanceWordList {

    private XWPFDocument document;
    private String existingFile;
    private String fileToBeCreated;
    private List<XWPFParagraph> paragraphs;
    private InputStream inputStream;
    private OpenFoldersAndDoc openDoc = new OpenFoldersAndDoc();
    private LocalDateTime now = LocalDateTime.now();
    private int year = now.getYear();
    private int month = now.getMonthValue();
    private int day = now.getDayOfMonth();
    private Map<String, List<String>> hashMap;
    private static ScheduleDB db;
    private Calendar cal;
    private String todaysDay;
    private int counter = 0;
    private int rader = 0;
    private String dataBaseValue;
    private static String documentFolder = System.getProperty("user.home") + "/Skrivbord/Attendance Tool/Dokument";
    private ExcelHandler excel = new ExcelHandler();

    public AttendanceWordList(String group, String tid, boolean fileExists) {
        try {
            try {
                db = new ScheduleDB();
            } catch (ClassNotFoundException ex) {
                JOptionPane.showMessageDialog(null, "Ett fel har inträffat (databas)", "Meddelande", JOptionPane.DEFAULT_OPTION);
            }
            cal = Calendar.getInstance();
            //todaysDay = cal.getDisplayName(Calendar.DAY_OF_WEEK, Calendar.LONG, Locale.ENGLISH);

            //Om filen inte finns, skapa den först.
            if (!fileExists) {
                existingFile = System.getProperty("user.home") + "/Skrivbord/Attendance Tool/Grupplistor/" + tid + ".docx";
                fileToBeCreated = System.getProperty("user.home") + "/Skrivbord/Attendance Tool/Grupplistor/" + group + ".docx";
                inputStream = new FileInputStream(existingFile);
                try {
                    document = new XWPFDocument(inputStream);
                    paragraphs = document.getParagraphs();
                    CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();
                    XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(document, sectPr);

                    //skapa header
                    CTP ctpHeader = CTP.Factory.newInstance();
                    //CTR ctrHeader = ctpHeader.addNewR();

                    XWPFParagraph headerParagraph = new XWPFParagraph(ctpHeader, document);

                    headerParagraph.setAlignment(ParagraphAlignment.CENTER);
                    setRun(headerParagraph.createRun(), "Calibri", 22, "000F00", "", false, false);
                    //headerParagraph.setVerticalAlignment(TextAlignment.CENTER);
                    XWPFParagraph[] parsHeader = new XWPFParagraph[1];
                    parsHeader[0] = headerParagraph;
                    policy.createHeader(XWPFHeaderFooterPolicy.FIRST, parsHeader);

                    OutputStream out2 = new FileOutputStream(fileToBeCreated);
                    document.write(out2);
                    out2.close();
                    JOptionPane.showMessageDialog(null, "En ny närvarolista, " + group + " (" + tid + "), har skapats.\n\n"
                            + "Vill du ta bort, " + group + ", hittar man den i mappen\n"
                            + "'Attendance Tool' → 'Grupplistor'.\n\n\n", "Meddelande", JOptionPane.DEFAULT_OPTION);
                } catch (IOException ex) {
                    JOptionPane.showMessageDialog(null, "Kontrollera sökvägen till Word filen", "Meddelande", JOptionPane.DEFAULT_OPTION);

                } catch (XmlException e) {
                    e.printStackTrace();
                }
            }

            existingFile = System.getProperty("user.home") + "/Skrivbord/Attendance Tool/Grupplistor/" + group + ".docx";
            fileToBeCreated = System.getProperty("user.home") + "/Skrivbord/Attendance Tool/Grupplistor/" + group + ".docx";
            inputStream = new FileInputStream(existingFile);
            try {
                document = new XWPFDocument(inputStream);
                paragraphs = document.getParagraphs();
                CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();
                XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(document, sectPr);

                //skapa header
                CTP ctpHeader = CTP.Factory.newInstance();
                //CTR ctrHeader = ctpHeader.addNewR();

                XWPFParagraph headerParagraph = new XWPFParagraph(ctpHeader, document);

                headerParagraph.setAlignment(ParagraphAlignment.CENTER);
                setRun(headerParagraph.createRun(), "Calibri", 22, "000F00", "Grupp: " + group, false, false);
                //headerParagraph.setVerticalAlignment(TextAlignment.CENTER);
                XWPFParagraph[] parsHeader = new XWPFParagraph[1];
                parsHeader[0] = headerParagraph;
                policy.createHeader(XWPFHeaderFooterPolicy.FIRST, parsHeader);

                OutputStream out2 = new FileOutputStream(fileToBeCreated);
                document.write(out2);
                out2.close();

            } catch (IOException ex) {
                JOptionPane.showMessageDialog(null, "Kontrollera sökvägen till Word filen", "Meddelande", JOptionPane.DEFAULT_OPTION);

            } catch (XmlException e) {
                e.printStackTrace();
            }
        } catch (FileNotFoundException ex) {
            try {
                if (group.equals("Excelfil saknas")) {
                    JOptionPane.showMessageDialog(null, "Excelfil saknas!", "Meddelande", JOptionPane.DEFAULT_OPTION);
                } else {
                    JOptionPane.showMessageDialog(null, "1. Kontrollera att Wordfilen '" + tid + "' finns i Attendance Tool --> Dokument.\n", "Meddelande", JOptionPane.DEFAULT_OPTION);
                }

                Desktop.getDesktop().open(new File(documentFolder));

                //Logger.getLogger(WordModifier.class.getName()).log(Level.SEVERE, null, ex);
            } catch (IOException ex1) {
                JOptionPane.showMessageDialog(null, "Kontrollera att mappen 'Grupplistor' finns i Attendance Tool mappen.", "Meddelande", JOptionPane.DEFAULT_OPTION);
            }
        }
    }

    //Ändra font, färg, storlek etc.
    private static void setRun(XWPFRun run, String fontFamily, int fontSize, String colorRGB, String text, boolean bold, boolean addBreak) {
        run.setFontFamily(fontFamily);
        run.setFontSize(fontSize);
        run.setColor(colorRGB);

        run.setUnderline(UnderlinePatterns.SINGLE);
        run.setBold(bold);
        run.setSubscript(VerticalAlign.SUBSCRIPT);
        run.setText(text);

        if (addBreak) {
            run.addBreak();
        }
    }

    public boolean addScheduleToWord(String group, Map<String, List<String>> hmap, Map<String, List<String>> endTimePartic, String todayDay) {
        hashMap = new TreeMap<>(hmap);
        todaysDay = todayDay;
        //System.out.println(todaysDay);
        try {
            String fm = " ";
            String em = " ";
            //String text = "test";//används inte för tillfället.

            inputStream = new FileInputStream(fileToBeCreated);
            document = new XWPFDocument(inputStream);

            XWPFTableCell tableCell;
            XWPFTableRow tableRow;
            XWPFTableRow tableRowOne;
            List<XWPFTableCell> tableCells;
            List<XWPFTableRow> tableRows;
            XWPFTable table;
            String cellText;
            String name;
            String caseNbr;
            String activityTime;
            String note = " ";
            List<String> values;
            List<XWPFTable> tables = document.getTables();
            tableRowOne = tables.get(0).getRow(0);//rad 1 och kolumn 1
            String addComment;

            XWPFParagraph addParagraph;
            XWPFRun run;

            tableCells = tableRowOne.getTableCells();

            activityTime = tableCells.get(0).getText().toLowerCase().trim();// em, fm eller heltids lista

            //for (XWPFTable table1 : tables) {
            table = tables.get(0);
            tableRows = table.getRows();
            rader = tableRows.size() - 1;
            //lägg till rader om det behövs
            if (hashMap.size() > rader) {
                for (int i = 0; i < hashMap.size() - rader; i++) {
                    table.addRow(table.getRow(table.getRows().size() - 1));//lägg till lakadana rader som den senaste raden
                }
                try (OutputStream out = new FileOutputStream(fileToBeCreated)) {
                    document.write(out);
                    out.close();
                }
                inputStream = new FileInputStream(fileToBeCreated);
                document = new XWPFDocument(inputStream);
                tables = document.getTables();
                table = tables.get(0);

                tableRows = table.getRows();
                rader = tableRows.size() - 1;
            }

            int i = 1;
            for (Map.Entry<String, List<String>> entry : hashMap.entrySet()) {
                name = entry.getKey();
                values = entry.getValue();
                caseNbr = values.get(0).trim();
                tableRow = tableRows.get(i);
                tableCells = tableRow.getTableCells();
                tableCell = tableCells.get(1);

                addComment = "";

                for (Map.Entry<String, List<String>> entry2 : endTimePartic.entrySet()) {
                    if (name.trim().equals(entry2.getKey().trim())) {
                        switch (entry2.getValue().get(0)) {
                            case "soon":
                                tableCell.setColor("e5e0e0");
                                addComment = "(Slutar snart)";
                                break;
                            case "today":
                                tableCell.setColor("e5e0e0");
                                addComment = "(Sista dagen idag)";
                                break;
                            case "end":
                                tableCell.setColor("e5e0e0");
                                addComment = "(Har slutat!)";
                                break;
                            case "future":
                                tableCell.setColor("e5e0e0");
                                addComment = "(Startar: " + excel.getStartDateByName(name.trim()) + ")";
                                break;
                        }
                        break;
                    }
                }
                //tableCell.setText(name.trim());

                //Raden ovanför kan användas för att sätta namnet, men
                //vi kan inte ändra textstorlek, färg etc.
                //-----------------------
                tableCell.removeParagraph(0);
                addParagraph = tableCell.addParagraph();
                run = addParagraph.createRun();
                run.setFontFamily("Calibri");
                run.setFontSize(14);
                run.setText(name.trim());
                //-----------------------
                tableCell = tableCells.get(2);
                tableCell.setText(caseNbr + "\n" + excel.getEndDateByName(name).trim());
                //todaysDay = "Monday";//OBS! Bara för test, så markera bort när jar-filen ska fixas!
                dataBaseValue = " ";
                if (AttendanceWordList.db.tableExist(caseNbr)) {
                    dataBaseValue = AttendanceWordList.db.getDaySchedule(todaysDay, caseNbr);
                    note = AttendanceWordList.db.getNote(caseNbr);

                    switch (dataBaseValue) {

                        case "SFI-EM":
                            fm = "SFI";
                            em = " ";
                            break;

                        case "FM-SFI":
                            fm = " ";
                            em = "SFI";
                            break;

                        case "SO-EM":
                            fm = "SO";
                            em = " ";
                            break;

                        case "FM-SO":
                            fm = " ";
                            em = "SO";
                            break;

                        case "P-EM":
                            fm = "P";
                            em = " ";
                            break;

                        case "FM-P":
                            fm = " ";
                            em = "P";
                            break;

                        case "LEDIG-EM":
                            fm = "LEDIG";
                            em = " ";
                            break;

                        case "FM-LEDIG":
                            fm = " ";
                            em = "LEDIG";
                            break;

                        case "SFI":
                            fm = "SFI";
                            em = "SFI";
                            break;

                        case "SO"://DELTID
                            fm = "SO";
                            em = "SO";
                            break;

                        case "SO-SO"://HELTID
                            fm = "SO";
                            em = "SO";
                            break;

                        case "P":
                            fm = "P";
                            em = "P";
                            break;

                        case "SFI-SO":
                            fm = "SFI";
                            em = "SO";
                            break;

                        case "SO-SFI":
                            fm = "SO";
                            em = "SFI";
                            break;

                        case "LEDIG":
                            fm = "LEDIG";
                            em = "LEDIG";
                            break;

                        case "SO-LEDIG":
                            fm = "SO";
                            em = "LEDIG";
                            break;

                        case "LEDIG-SO":
                            fm = "LEDIG";
                            em = "SO";
                            break;

                        case "SFI-LEDIG":
                            fm = "SFI";
                            em = "LEDIG";
                            break;

                        case "LEDIG-SFI":
                            fm = "LEDIG";
                            em = "SFI";
                            break;

                        case "P-SO":
                            fm = "P";
                            em = "SO";
                            break;

                        case "SO-P":
                            fm = "SO";
                            em = "P";
                            break;

                        case "P-SFI":
                            fm = "P";
                            em = "SFI";
                            break;

                        case "SFI-P":
                            fm = "SFI";
                            em = "P";
                            break;

                        case "HELTID":
                            fm = " ";
                            em = " ";
                            break;

                        case "EM":
                            fm = " ";
                            em = " ";
                            break;

                        case "FM":
                            fm = " ";
                            em = " ";
                            break;

                        case "SJUK"://Deltid
                            fm = "SJUK";
                            em = "SJUK";
                            break;

                        case "SJUK-SJUK"://Heltid
                            fm = "SJUK";
                            em = "SJUK";
                            break;

                        case "FL"://Deltid
                            fm = "FL";
                            em = "FL";
                            break;

                        case "FL-FL"://heltid
                            fm = "FL";
                            em = "FL";
                            break;

                        case "SFI-FL"://heltid
                            fm = "SFI";
                            em = "FL";
                            break;

                        case "FL-SFI"://heltid
                            fm = "FL";
                            em = "SFI";
                            break;

                        case "BLANDAT":
                            fm = " ";
                            em = " ";
                            break;

                        case "JOBB-JOBB":
                            fm = "JOBB";
                            em = "JOBB";
                            break;

                        case "JOBB":
                            fm = "JOBB";
                            em = "JOBB";
                            break;

                        case "FM-JOBB":
                            fm = " ";
                            em = "JOBB";
                            break;

                        case "JOBB-EM":
                            fm = "JOBB";
                            em = " ";
                            break;

                        case "STOM":
                            fm = "STOM";
                            em = "STOM";
                            break;

                        case "FM-STOM":
                            fm = " ";
                            em = "STOM";
                            break;

                        case "STOM-EM":
                            fm = "STOM";
                            em = " ";
                            break;

                        case "UF":
                            fm = "UF";
                            em = "UF";
                            break;

                        case "FM-UF":
                            fm = " ";
                            em = "UF";
                            break;

                        case "UF-EM":
                            fm = "UF";
                            em = " ";
                            break;

                        case "VAB-VAB":
                            fm = "VAB";
                            em = "VAB";
                            break;

                        case "VAB":
                            fm = "VAB";
                            em = "VAB";
                            break;

                        case "FM-AKT_100%":
                            fm = " ";
                            em = "AKT";

                            break;
                        case "AKT-EM_100%":
                            fm = "AKT";
                            em = " ";
                            break;

                        case "FM-AKT_50%":
                            fm = " ";
                            em = "AKT";
                            break;

                        case "AKT-EM_50%":
                            fm = "AKT";
                            em = " ";
                            break;

                        case "AKT":
                            fm = "AKT";
                            em = "AKT";
                            break;

                        case "AKT-AKT_100%":
                            fm = "AKT";
                            em = "AKT";
                            break;

                        case "AKT-AKT_50%":
                            fm = "AKT";
                            em = "AKT";
                            break;

                        case "FM-KOMV_100%":
                            fm = " ";
                            em = "KOMV";

                            break;

                        case "KOMV-EM_100%":
                            fm = "KOMV";
                            em = " ";
                            break;

                        case "FM-KOMV_50%":
                            fm = " ";
                            em = "KOMV";
                            break;

                        case "KOMV-EM_50%":
                            fm = "KOMV";
                            em = " ";
                            break;

                        case "KOMV":
                            fm = "KOMV";
                            em = "KOMV";
                            break;

                        case "KOMV-KOMV_100%":
                            fm = "KOMV";
                            em = "KOMV";
                            break;

                        case "KOMV-KOMV_50%":
                            fm = "KOMV";
                            em = "KOMV";
                            break;

                        default:
                            fm = " ";
                            em = " ";
                            break;
                    }

                } else {
                    fm = " ";
                    em = " ";
                }

                counter++;

                switch (activityTime) {
                    case "em":
                        tableCell = tableCells.get(3);
                        if (!em.equals(" ")) {
                            tableCell.setColor("e5e0e0");
                        }
                        tableCell.setText(em);

                        if (!note.isEmpty() && !note.equals(" ") && !note.equals("")) {
                            tableCell = tableCells.get(4);//anteckning
                            if (!addComment.isEmpty() && !addComment.equals("")) {
                                if (dataBaseValue.equals("BLANDAT")) {
                                    tableCell.setText("Har ett blandat schema idag. " + note + " | " + addComment);
                                } else {
                                    tableCell.setText(note + " | " + addComment);
                                }
                            } else {
                                if (dataBaseValue.equals("BLANDAT")) {
                                    tableCell.setText("Har ett blandat schema idag. " + note);
                                } else {
                                    tableCell.setText(note);
                                }
                            }
                        } else if (note.isEmpty() || note.equals(" ") || note.equals("")) {
                            tableCell = tableCells.get(4);
                            if (!addComment.isEmpty() && !addComment.equals("")) {
                                //anteckning
                                if (dataBaseValue.equals("BLANDAT")) {
                                    tableCell.setText("Har ett blandat schema idag. " + addComment);
                                } else {
                                    tableCell.setText(addComment);
                                }
                            } else {
                                if (dataBaseValue.equals("BLANDAT")) {
                                    tableCell.setText("Har ett blandat schema idag.");
                                }
                            }
                        }
                        break;
                    //heltidslistan har normalt 10 tabeller
                    case "fm":
                        tableCell = tableCells.get(3);
                        if (!fm.equals(" ")) {
                            tableCell.setColor("e5e0e0");
                        }
                        tableCell.setText(fm);

                        if (!note.isEmpty() && !note.equals(" ") && !note.equals("")) {
                            tableCell = tableCells.get(4);//anteckning
                            if (!addComment.isEmpty() && !addComment.equals("")) {
                                if (dataBaseValue.equals("BLANDAT")) {
                                    tableCell.setText("Har ett blandat schema idag. " + note + " | " + addComment);
                                } else {
                                    tableCell.setText(note + " | " + addComment);
                                }
                            } else {
                                if (dataBaseValue.equals("BLANDAT")) {
                                    tableCell.setText("Har ett blandat schema idag. " + note);
                                } else {
                                    tableCell.setText(note);
                                }
                            }
                        } else if (note.isEmpty() || note.equals(" ") || note.equals("")) {
                            tableCell = tableCells.get(4);
                            if (!addComment.isEmpty() && !addComment.equals("")) {
                                //anteckning
                                if (dataBaseValue.equals("BLANDAT")) {
                                    tableCell.setText("Har ett blandat schema idag. " + addComment);
                                } else {
                                    tableCell.setText(addComment);
                                }
                            } else {
                                if (dataBaseValue.equals("BLANDAT")) {
                                    tableCell.setText("Har ett blandat schema idag.");
                                }
                            }
                        }
                        break;
                    case "h":
                        tableCell = tableCells.get(3);
                        if (!fm.equals(" ")) {
                            tableCell.setColor("e5e0e0");
                        } //else if (dataBaseValue.equals("BLANDAT")) {
                        //tableCell.setColor("F8E0F7");
                        //}
                        tableCell.setText(fm);

                        tableCell = tableCells.get(4);
                        if (!em.equals(" ")) {
                            tableCell.setColor("e5e0e0");
                        } //else if (dataBaseValue.equals("BLANDAT")) {
                        //tableCell.setColor("e5e0e0");
                        //}
                        tableCell.setText(em);

                        if (!note.isEmpty() && !note.equals(" ") && !note.equals("")) {
                            tableCell = tableCells.get(5);//anteckning
                            if (!addComment.isEmpty() && !addComment.equals("")) {
                                if (dataBaseValue.equals("BLANDAT")) {
                                    tableCell.setText("Har ett blandat schema idag. " + note + " | " + addComment);
                                } else {
                                    tableCell.setText(note + " | " + addComment);
                                }
                            } else {
                                if (dataBaseValue.equals("BLANDAT")) {
                                    tableCell.setText("Har ett blandat schema idag. " + note);
                                } else {
                                    tableCell.setText(note);
                                }
                            }
                        } else if (note.isEmpty() || note.equals(" ") || note.equals("")) {
                            tableCell = tableCells.get(5);
                            if (!addComment.isEmpty() && !addComment.equals("")) {
                                //anteckning
                                if (dataBaseValue.equals("BLANDAT")) {
                                    tableCell.setText("Har ett blandat schema idag. " + addComment);
                                } else {
                                    tableCell.setText(addComment);
                                }
                            } else {
                                if (dataBaseValue.equals("BLANDAT")) {
                                    tableCell.setText("Har ett blandat schema idag.");
                                }
                            }
                        }
                        break;
                    default:
                        JOptionPane.showMessageDialog(null, "Ett fel inträffade när grupplistan skulle skapas!\n"
                                + "Kontakta ansvarig\n", "Meddelande", JOptionPane.DEFAULT_OPTION);
                        break;
                }
                i = i + 1;
                note = " ";
            }

            try (OutputStream out = new FileOutputStream(fileToBeCreated)) {
                document.write(out);
                out.close();
            }
        } catch (IOException | SQLException ex) {
            //System.out.println(ex);
            JOptionPane.showMessageDialog(null, "Närvarolistan " + group + " kan vara öppen, eller så "
                    + "finns den inte", "Meddelande", JOptionPane.DEFAULT_OPTION);
            return true;
        }
        openDoc.openAbsenceListFolder();
        return false;
    }

    /*
    //används inte just nu. Kan användas senare
    private static void removeParagraphs(XWPFTableCell tableCell) {
        int count = tableCell.getParagraphs().size();
        for (int i = 0; i < count; i++) {
            tableCell.removeParagraph(i);
        }
    }
     */
}
