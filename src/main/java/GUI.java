import java.awt.*;
import java.awt.event.*;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.SQLException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.time.temporal.ChronoUnit;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.GregorianCalendar;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.prefs.Preferences;
import javax.swing.*;
import javax.swing.border.Border;
import javax.swing.plaf.basic.BasicComboBoxRenderer;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableCellEditor;
import javax.swing.table.TableColumnModel;
import javax.swing.table.TableModel;
import javax.swing.table.TableRowSorter;
import javax.swing.text.AbstractDocument;
import javax.swing.text.AttributeSet;
import javax.swing.text.BadLocationException;
import javax.swing.text.DocumentFilter;

/**
 *
 * @author Burim Sadriu Copyright© 2021, All Rights Reserved
 */
@SuppressWarnings("serial")
public final class GUI extends JPanel {

    public static final String[] COLUMNS = {"Närvaro", "Deltagare", "Ärendenummer", "Avvikelse", "Anteckningar", "Dagar kvar"};
    private CheckBoxDefaultTableModel model = new CheckBoxDefaultTableModel(COLUMNS, 0);
    private final JTable table = new JTable(model);
    static Preferences prefs = Preferences.userRoot();
    //private static String anord = prefs.get("anordString", "");
    //static String samord = prefs.get("samordString", "");
    //static String tel2 = prefs.get("telString", "");
    //static String mail2 = prefs.get("mailString", "");
    //static String aktivitet = prefs.get("aktivitetString", "");
    //static String samord = prefs.get("samordString", "");
    static String chef = "x y";

    static public ImageIcon imageIcon;
    final public String img = "/Images/workfast4.png";
    //ImageIcon iconImg = new ImageIcon(getClass().getResource("/Images/icon2.png"));
    static public ExcelHandler excel;
    Calendar now = GregorianCalendar.getInstance();
    static String groupNow = "";
    String[] absence;
    String[] presence;
    String[] grupper;
    String actualValue = "";
    JComboBox måndag;
    JComboBox tisdag;
    JComboBox onsdag;
    JComboBox torsdag;
    JComboBox fredag;
    String[] grupperInnan;

    String[] day = {"1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16",
        "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31"};

    String[] GUIabsenceOpt = {" ", "S", "A", "V", "Ö", "P", "AF", "0,15", "0,30", "0,45", "1", "1,15", "1,30", "1,45", "2", "2,15", "2,30", "2,45",
        "3", "3,15", "3,30", "3,45", "4", "4,15", "4,30", "4,45", "5", "5,15", "5,30", "5,45", "6", "6,15", "6,30", "6,45",
        "7", "7,15", "7,30", "7,45", "8"};

    String[] changeAbsenceOpt = {"Återställ", "x", "h", "d", "S", "A", "V", "Ö", "P", "AF", "SFI", "SO", "0,15", "0,30", "0,45", "1", "1,15", "1,30",
        "1,45", "2", "2,15", "2,30", "2,45", "3", "3,15", "3,30", "3,45", "4", "4,15", "4,30", "4,45", "5",
        "5,15", "5,30", "5,45", "6", "6,15", "6,30", "6,45", "7", "7,15", "7,30", "7,45", "8"};

    String[] scheduleOpt = {"HELTID", "EM", "FM", "SFI-EM", "FM-SFI", "SFI", "SO-EM", "FM-SO",
        "SO-SO", "SO", "FM-P", "P-EM", "P", "SO-P", "P-SO", "SFI-P", "P-SFI", "KOMV-EM_100%", "FM-KOMV_100%",
        "KOMV-EM_50%", "FM-KOMV_50%", "KOMV", "KOMV-KOMV_100%", "KOMV-KOMV_50%", "SO-SFI", "SFI-SO",
        "SO-LEDIG", "LEDIG-SO", "SFI-LEDIG", "LEDIG-SFI", "FM-LEDIG", "LEDIG-EM", "LEDIG",
        "SJUK-SJUK", "SJUK", "FL-FL", "FL", "FL-SFI", "SFI-FL", "BLANDAT", "JOBB-JOBB", "JOBB", "FM-JOBB", "JOBB-EM",
        "STOM", "FM-STOM", "STOM-EM", "UF", "FM-UF", "UF-EM", "VAB-VAB", "VAB", "AKT-EM_100%", "FM-AKT_100%",
        "AKT-EM_50%", "FM-AKT_50%", "AKT", "AKT-AKT_100%", "AKT-AKT_50%"};

    String[] ScheduleOptTipText = {"Heltid hos oss (100%)", "Bara på EM hos oss (50%)",
        "Bara på FM hos oss (50%)", "SFI på FM, hos oss på EM (50%)", "Hos oss på FM, SFI på EM (50%)", "SFI (ej frånvaro)",
        "Samhälle på FM, hos oss på EM (100%)", "Hos oss på FM, samhälle på EM (100%)", "Samhälle hela dagen (100%)",
        "Samhälle (50%)", "Hos oss på FM, praktik på EM (50%)", "Praktik på FM, hos oss på EM (50%)", "Praktik (ej frånvaro)",
        "Samhälle på FM, praktik på EM (50%)", "Praktik på FM, samhälle på EM (50%)", "SFI på FM, praktik på EM (ej frånvaro)", "Praktik på FM, SFI på EM (ej frånvaro)",
        "Komvux på FM, hos oss på EM (100%)", "Hos oss på FM, Komvux på EM (100%)", "Komvux på FM, hos oss på EM (50%)", "Hos oss på FM, Komvux på EM (50%)",
        "Komvux (ej frånvaro)", "Komvux (100%)", "Komvux (50%)", "Samhälle på FM, SFI på EM (50%)", "SFI på FM, samhälle på EM (50%)", "Samhälle på FM, ledig på EM (50%)", "Ledig på FM, samhälle på EM (50%)",
        "SFI på FM, ledig på EM (ej frånvaro)", "Ledig på FM, SFI på EM (ej frånvaro)", "FM hos oss, ledig på EM (50%)", "Ledig på FM, hos oss på EM (50%)", "Ledig (ej frånvaro)",
        "Sjuk (100%)", "Sjuk (50%)", "Föräldraledig (100%)",
        "Föräldraledig (50%)", "Föräldraledig på FM, SFI på EM (50%)", "SFI på FM, Föräldraledig på EM (50%)", "Deltagaren kommer ibland på FM och ibland på EM (50%)",
        "Arbete hela dagen (100%)", "Arbete (50%)", "Hos oss på FM, arbete på EM (100%)", "Arbete på FM, hos oss på EM (100%)", "Stöd och matchning (ej frånvaro)",
        "Hos oss på FM, stöd och matchning på EM (50%)", "Stöd och matchning på FM, hos oss på EM (50%)", "Ung framtid (AF) (ej frånvaro)",
        "Hos oss på FM, ung framtid på EM (50%)", "Ung framtid på FM, hos oss på EM (50%)", "Vård av barn (100%)", "Vård av barn (50%)",
        "Aktivitet på FM, hos oss på EM (100%)", "Hos oss på FM, aktivitet på EM (100%)", "Aktivitet på FM, hos oss på EM (50%)", "Hos oss på FM, aktivitet på EM (50%)",
        "Aktivitet (ej frånvaro)", "Aktivitet (100%)", "Aktivitet (50%)"};

    String[] registered = {"Heltid", "FM", "EM", "Deltid"};
    String[] whichTime = {"Heltid", "FM", "EM"};
    final JTextField text;
    static JFrame frame;
    String date = "Dagens datum";
    String saved = "Sparat";
    String raportMonth = " ";
    String raportYear = " ";
    boolean sparat = false;
    boolean nullGroup = false;
    static JLabel save = new JLabel();
    JLabel numberOfPa = new JLabel();
    Certificate cf;
    JButton lastButton = new JButton("test");
    JTextField nameField = new JTextField(19);
    JTextField groupField = new JTextField(19);
    JTextField startDate = new JTextField(19);
    JTextField endDate = new JTextField(19);
    JTextField caseField = new JTextField(19);
    int buttonTextSize = 14;
    String font = "Arial";
    JButton schema, changeButton, saveInfo;
    String monAct;
    String tueAct;
    String wedAct;
    String thuAct;
    String friAct;
    String note;
    Timer indScheduleTimer;
    Timer changeInfoTimer;
    DateFormat dateF;
    JComboBox mainGUIcomboBox;
    Map<String, List<String>> hashMap;
    TableCellEditor editor;
    String priviousDate = " ";
    JTextArea teacherName = new TextFieldLimit(150, 35, 50);
    JTextArea city = new TextFieldLimit(150, 35, 40);
    JTextArea aboutParticipant = new TextFieldLimit(400, 300, 720);
    List<String> grupper3;
    JFrame scheduleFrame = null;
    JFrame infoFrame = null;
    String latestGroupp = " ";
    ImportExcelFile copyFile;
    AttendanceWordList wm;

    Object[] certInfo = {
        "Lärare / Handledare:", teacherName, "Stad:", city,
        "Om deltagaren:", aboutParticipant
    };

    String groupTrimed = "";
    String caseTrimed = "";
    static boolean checked = false;
    static boolean canceled = false;

    Map<String, List<String>> endTimePartic;
    List<String> wrongDateList;

    TableColumnModel tcm;
    DefaultTableCellRenderer centerRenderer;

    DefaultComboBoxModel jcomboModel;
    JComboBox comboTypesList;
    JComboBox monthList;

    static String tom = " ";

    static JButton checkAllButton;

    Calendar cal = Calendar.getInstance();
    int nbrOfP = 0;
    LookAndFeel previousLF;
    String nbrOfPa = "";
    static ScheduleDB db;
    boolean windowClosed = true;
    SQLiteJDBCLoader clean = new SQLiteJDBCLoader();

    public GUI() {

        try {
            clean.initialize();
        } catch (Exception ex) {
            Logger.getLogger(GUI.class.getName()).log(Level.SEVERE, null, ex);
        }

        nameField.setPreferredSize(new Dimension(19, 24));
        groupField.setPreferredSize(new Dimension(19, 24));
        startDate.setPreferredSize(new Dimension(19, 24));
        endDate.setPreferredSize(new Dimension(19, 24));
        caseField.setPreferredSize(new Dimension(19, 24));

        try {
            copyFile = new ImportExcelFile();
        } catch (Exception ex) {
        }

        try {
            db = new ScheduleDB();
        } catch (ClassNotFoundException ex) {
            JOptionPane.showMessageDialog(null, "Ett fel har inträffat (databas)", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
        cf = new Certificate();
        numberOfPa.setForeground(new Color(0x000000));
        aboutParticipant.setTransferHandler(null);// låt inte användaren klistra in något 
        previousLF = UIManager.getLookAndFeel();
        excel = new ExcelHandler();
        //imageIcon = (new ImageIcon(getClass().getResource("/Images/icon2.png")));
        this.hashMap = new HashMap<>();
        table.setFillsViewportHeight(true);
        //table.getTableHeader().setFont(new Font(font, Font.BOLD, 16));
        table.setFont(new Font(font, Font.PLAIN, buttonTextSize));

        nameField.setToolTipText("Deltagarens för- och efternamn");
        groupField.setToolTipText("Befintliga gruppens namn. Du kan också skapa en ny grupp här");
        caseField.setToolTipText("Deltagarens ärendenummer");
        startDate.setToolTipText("Deltagarens startdatum i formatet ÅÅÅÅ-MM-DD");
        endDate.setToolTipText("Deltagarens slutdatum i formatet ÅÅÅÅ-MM-DD");

        nameField.setText("Namn");
        groupField.setText("Grupp");
        caseField.setText("Ärendenummer");
        startDate.setText("Startdatum");
        endDate.setText("Slutdatum");

        nameField.setForeground(Color.gray);
        groupField.setForeground(Color.gray);
        caseField.setForeground(Color.gray);
        startDate.setForeground(Color.gray);
        endDate.setForeground(Color.gray);

        nameField.setFont(new Font(font, Font.PLAIN, buttonTextSize));
        groupField.setFont(new Font(font, Font.PLAIN, buttonTextSize));
        caseField.setFont(new Font(font, Font.PLAIN, buttonTextSize));
        startDate.setFont(new Font(font, Font.PLAIN, buttonTextSize));
        endDate.setFont(new Font(font, Font.PLAIN, buttonTextSize));

        dateF = new SimpleDateFormat("yyyy-MM-dd");

        nameField.addKeyListener(new KeyAdapter() {
            @Override
            public void keyPressed(KeyEvent ke) {
                nameField.setForeground(Color.black);
            }
        });
        groupField.addKeyListener(new KeyAdapter() {
            @Override
            public void keyPressed(KeyEvent ke) {
                groupField.setForeground(Color.black);
            }
        });
        caseField.addKeyListener(new KeyAdapter() {
            @Override
            public void keyPressed(KeyEvent ke) {
                caseField.setForeground(Color.black);
            }
        });
        startDate.addKeyListener(new KeyAdapter() {
            @Override
            public void keyPressed(KeyEvent ke) {
                startDate.setForeground(Color.black);
            }
        });
        endDate.addKeyListener(new KeyAdapter() {
            @Override
            public void keyPressed(KeyEvent ke) {
                endDate.setForeground(Color.black);
            }
        });

        nameField.addFocusListener(new FocusListener() {

            @Override
            public void focusGained(FocusEvent e) {
                //when selected...
                if (nameField.getText().equals("Namn")) {
                    nameField.setText("");
                }
            }

            @Override
            public void focusLost(FocusEvent e) {
                //when not selected..
                if (nameField.getText().isEmpty()) {
                    nameField.setForeground(Color.gray);
                    nameField.setText("Namn");
                }
            }
        });
        groupField.addFocusListener(new FocusListener() {

            @Override
            public void focusGained(FocusEvent e) {
                //when selected...
                if (groupField.getText().equals("Grupp")) {
                    groupField.setText("");
                }
            }

            @Override
            public void focusLost(FocusEvent e) {
                //when not selected..
                if (groupField.getText().isEmpty()) {
                    groupField.setForeground(Color.gray);
                    groupField.setText("Grupp");
                }
            }
        });
        caseField.addFocusListener(new FocusListener() {

            @Override
            public void focusGained(FocusEvent e) {
                //when selected...
                if (caseField.getText().equals("Ärendenummer")) {
                    caseField.setText("");
                }
            }

            @Override
            public void focusLost(FocusEvent e) {
                //when not selected..
                if (caseField.getText().isEmpty()) {
                    caseField.setForeground(Color.gray);
                    caseField.setText("Ärendenummer");
                }
            }
        });
        startDate.addFocusListener(new FocusListener() {

            @Override
            public void focusGained(FocusEvent e) {
                //when selected...
                if (startDate.getText().equals("Startdatum")) {
                    startDate.setText("");
                }
            }

            @Override
            public void focusLost(FocusEvent e) {
                //when not selected..
                if (startDate.getText().isEmpty()) {
                    startDate.setForeground(Color.gray);
                    startDate.setText("Startdatum");
                }
            }
        });
        endDate.addFocusListener(new FocusListener() {

            @Override
            public void focusGained(FocusEvent e) {
                //when selected...
                if (endDate.getText().equals("Slutdatum")) {
                    endDate.setText("");
                }
            }

            @Override
            public void focusLost(FocusEvent e) {
                //when not selected..
                if (endDate.getText().isEmpty()) {
                    endDate.setForeground(Color.gray);
                    endDate.setText("Slutdatum");
                }
            }
        });

        //lägg till comboBox i den fjärde columnen
        mainGUIcomboBox = new JComboBox(GUIabsenceOpt);
        mainGUIcomboBox.setMaximumRowCount(20);
        editor = new DefaultCellEditor(mainGUIcomboBox);
        table.getColumnModel().getColumn(3).setCellEditor(editor);

        table.setPreferredScrollableViewportSize(new Dimension(550, 2000));
        //table.setBorder(BorderFactory.createLineBorder(new Color(0xFFAA00)));
        table.setRowHeight(27);

        //ändra breden på listorna
        tcm = table.getColumnModel();
        tcm.getColumn(0).setMaxWidth(115);
        tcm.getColumn(1).setMaxWidth(385);
        tcm.getColumn(2).setMaxWidth(205);
        tcm.getColumn(3).setMaxWidth(130);
        tcm.getColumn(4).setMaxWidth(220);
        tcm.getColumn(5).setMaxWidth(150);
        tcm.getColumn(0).setPreferredWidth(0);
        tcm.getColumn(1).setPreferredWidth(2);
        tcm.getColumn(2).setPreferredWidth(0);
        tcm.getColumn(3).setPreferredWidth(0);
        tcm.getColumn(4).setPreferredWidth(30);
        tcm.getColumn(5).setPreferredWidth(1);
        tcm.setColumnMargin(2);

        table.removeColumn(table.getColumnModel().getColumn(4));//göm denna kolumn då den inte används för tillfället.

        table.setSelectionBackground(new Color(188, 232, 200));
        centerRenderer = new DefaultTableCellRenderer();
        centerRenderer.setHorizontalAlignment(JLabel.CENTER);

        //table.getColumnModel().getColumn(0).setCellRenderer(centerRenderer);
        table.getColumnModel().getColumn(2).setCellRenderer(centerRenderer);
        table.getColumnModel().getColumn(3).setCellRenderer(centerRenderer);
        table.getColumnModel().getColumn(4).setCellRenderer(centerRenderer);
        table.moveColumn(0, 2); // byt plats på kolumner. Flytta kolumn 0 till plats 2.
        table.moveColumn(0, 1);
        table.moveColumn(4, 2);
        table.moveColumn(3, 4);
        table.moveColumn(4, 3);

        grupper = excel.getGroups();
        grupperInnan = grupper.clone();

        // Create the combo box, and set 2nd item as Default
        comboTypesList = new JComboBox(grupper);
        if (grupper.length <= 0) {
            comboTypesList.setSelectedIndex(-1);
        } else {
            comboTypesList.setSelectedIndex(0);
        }
        comboTypesList.setMaximumRowCount(7);
        Object obj = comboTypesList.getSelectedItem();
        if (grupper.length > 0) {
            groupNow = obj.toString();
        } else {
            JOptionPane.showMessageDialog(null, "Excelfilen är tom", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
        comboTypesList.setPreferredSize(new Dimension(163, 22));
        //comboTypesList.setBackground(Color.LIGHT_GRAY);

        comboTypesList.addActionListener((ActionEvent e) -> {
            JComboBox jcmbType = (JComboBox) e.getSource();
            groupNow = jcmbType.getItemAt(comboTypesList.getSelectedIndex()).toString();

            if (!groupNow.equals("Excelfil saknas") && !groupNow.equals("Excelen är tom")) {
                putNamesOnList(groupNow);
            }

        });

        //table.setRowMargin(1);
        JButton selectDate = new JButton("Välj datum");
        selectDate.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                cal = Calendar.getInstance();
                //text.setBackground(Color.white);
                text.setText(new DatePicker(frame).setPickedDate());
                raportMonth = text.getText();

                if (raportMonth.length() > 0) {
                    raportYear = String.valueOf(text.getText().charAt(0) + "" + text.getText().charAt(1) + "" + text.getText().charAt(2)
                            + "" + text.getText().charAt(3));
                    raportMonth = String.valueOf(text.getText().charAt(5) + "" + text.getText().charAt(6));
                } else {
                    raportYear = Integer.toString(cal.get(Calendar.YEAR));
                    raportMonth = Integer.toString(cal.get(Calendar.MONTH) + 1);
                    if ((cal.get(Calendar.MONTH) + 1) > 0 && (cal.get(Calendar.MONTH) + 1) < 10) {
                        raportMonth = "0" + raportMonth;
                    }
                }
                setButtonColor(selectDate);

                date = text.getText();

                if (date.isEmpty() || date.equals("") || date.equals(" ")) {

                    text.setText(dateF.format(cal.getTime()));

                    date = text.getText();
                    if (!priviousDate.equals(date)) {
                        save.setText("Ej sparat");
                        save.setForeground(new Color(153, 0, 0));
                        sparat = false;
                    }
                } else if (!save.getText().equals("Ej sparat")) {
                    save.setText("Ej sparat");
                    save.setForeground(new Color(153, 0, 0));
                    sparat = false;
                }
                priviousDate = date;
            }
        });

        JPanel btnPanel = new JPanel(new GridLayout(4, 1, 5, 2));
        btnPanel.setPreferredSize(new Dimension(0, 135));
        btnPanel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        JPanel imagePanel = new JPanel() {
            @Override
            public Dimension getPreferredSize() {
                return new Dimension(267, 100);

            }
        };

        JButton searchParticipant = new JButton("Sök deltagare");
        searchParticipant.addActionListener((ActionEvent e) -> {
            try {
                excel.searchParticipant(db);

            } catch (NullPointerException ex) {

            }
            setButtonColor(searchParticipant);
        });

        JButton absenceReport = new JButton("Skapa periodisk rapport");
        absenceReport.addActionListener((ActionEvent e) -> {
            if (!groupNow.equals("Excelen är tom")) {
                excel.createPeriodicReports();
                setButtonColor(absenceReport);
            } else {
                JOptionPane.showMessageDialog(null, "Periodiska rapporter kan inte skapas\n"
                        + "eftersom excelen är tom", "Meddelande", JOptionPane.DEFAULT_OPTION);
            }
        });

        JButton addPart = new JButton("Lägg till");
        addPart.addActionListener((ActionEvent e) -> {
            String nameF = nameField.getText().replace("/", "").replace("\\", "").trim();//i en sträng kan snedstreck uppfattas som fil-delare!
            String groupF = groupField.getText().replace("/", "").replace("\\", "").trim();
            String caseF = caseField.getText().replace("/", "").replace("\\", "").trim();
            String startDateF = startDate.getText().replace("/", "").replace("\\", "").trim();
            String endDateF = endDate.getText().replace("/", "").replace("\\", "").trim();

            if (isNum(date)) {
                int dayOfMonth = Integer.parseInt(date);
            }

            try {
                if (nameF.isEmpty() == false && groupF.isEmpty() == false
                        && caseF.isEmpty() == false && startDateF.isEmpty() == false
                        && endDateF.isEmpty() == false && nameF.length() > 0
                        && groupF.length() > 0 && caseF.length() > 0
                        && startDateF.length() > 0 && endDateF.length() > 0
                        && !nameF.equals("Namn") && !groupF.equals("Grupp")
                        && !caseF.equals("Ärendenummer") && !startDateF.equals("Startdatum")
                        && !endDateF.trim().equals("Slutdatum")) {
                    if ((isValidFormat("yyyy-MM-dd", startDateF) && (isValidFormat("yyyy-MM-dd", endDateF)))) {
                        groupTrimed = groupF;
                        if (!excel.checkDublicates(caseF)) {
                            if (!excel.checkForDublicates(nameF, groupF)) {
                                caseTrimed = caseF;

                                //ärendenumret är vanligtvis åtta siffror långt
                                if (caseTrimed.length() < 8 || caseTrimed.length() > 8) {
                                    int reply = JOptionPane.showConfirmDialog(null, "Ärendenumret brukar vara åtta siffror långt.\n"
                                            + "Skrev du rätt ärendenummer?", "Säkerhetsfråga", JOptionPane.YES_NO_OPTION, JOptionPane.PLAIN_MESSAGE);
                                    if (reply == JOptionPane.YES_OPTION) {

                                        //om excelen är öppen ta inte bort texten från fälten eftersom inget lagts till.
                                        if (nyGrupp(grupper, groupTrimed)) {

                                            String tid = (String) JOptionPane.showInputDialog(null, "Du håller på att skapa en ny grupp, " + groupTrimed + "\n\n"
                                                    + "Välj schema för " + nameF + "\n", "Schema",
                                                    JOptionPane.PLAIN_MESSAGE,
                                                    null,
                                                    registered,
                                                    registered[0]);

                                            if ((tid != null) && (tid.length() > 0)) {
                                                if (!excel.addParticipant(nameF, groupTrimed, caseTrimed, startDateF, endDateF, tid)) {
                                                    if (tid.equals("Deltid") && !GUI.db.tableExist(caseTrimed)) {
                                                        JOptionPane.showMessageDialog(null, "Du valde ett deltidsschema.\n"
                                                                + "Om deltagaren är här på FM, EM och ibland Heltid så ska man skapa\n"
                                                                + "ett individuellt schema istället ('Individuellt schema' - knappen)", "Meddelande", JOptionPane.DEFAULT_OPTION);
                                                    }
                                                    grupper = excel.getGroups();
                                                    grupperInnan = grupper.clone();
                                                    jcomboModel = new DefaultComboBoxModel(grupper);
                                                    comboTypesList.setModel(jcomboModel);
                                                    comboTypesList.setSelectedItem(groupTrimed);
                                                    putNamesOnList(groupTrimed);
                                                    setButtonColor(addPart);
                                                    groupNow = groupTrimed;
                                                    nameField.setText("");
                                                    caseField.setText("");
                                                }
                                            }
                                        } else {
                                            String tid = (String) JOptionPane.showInputDialog(null, "Välj schema för " + nameF + "\n", "Schema",
                                                    JOptionPane.PLAIN_MESSAGE,
                                                    null,
                                                    registered,
                                                    registered[0]);

                                            if ((tid != null) && (tid.length() > 0)) {
                                                if (!excel.addParticipant(nameF, groupTrimed, caseTrimed, startDateF, endDateF, tid)) {
                                                    if (tid.equals("Deltid") && !GUI.db.tableExist(caseTrimed)) {
                                                        JOptionPane.showMessageDialog(null, "Du valde ett deltidsschema.\n"
                                                                + "Om deltagaren är här på FM, EM och ibland Heltid så ska man skapa\n"
                                                                + "ett individuellt schema istället ('Individuellt schema' - knappen)", "Meddelande", JOptionPane.DEFAULT_OPTION);
                                                    }
                                                    putNamesOnList(groupNow);
                                                    setButtonColor(addPart);
                                                    nameField.setText("");
                                                    caseField.setText("");
                                                }
                                            }
                                        }
                                    }
                                } else {

                                    if (nyGrupp(grupper, groupTrimed)) {
                                        String tid = (String) JOptionPane.showInputDialog(null, "Du håller på att skapa en ny grupp, " + groupTrimed + ".\n\n"
                                                + "Välj schema för " + nameF + "\n", "Schema",
                                                JOptionPane.PLAIN_MESSAGE,
                                                null,
                                                registered,
                                                registered[0]);

                                        if ((tid != null) && (tid.length() > 0)) {
                                            if (!excel.addParticipant(nameF, groupTrimed, caseTrimed, startDateF, endDateF, tid)) {
                                                if (tid.equals("Deltid") && !GUI.db.tableExist(caseTrimed)) {
                                                    JOptionPane.showMessageDialog(null, "Du valde ett deltidsschema.\n"
                                                            + "Om deltagaren är här på FM, EM och ibland Heltid så ska man skapa\n"
                                                            + "ett individuellt schema istället ('Individuellt schema' - knappen)", "Meddelande", JOptionPane.DEFAULT_OPTION);
                                                }
                                                grupper = excel.getGroups();
                                                grupperInnan = grupper.clone();
                                                jcomboModel = new DefaultComboBoxModel(grupper);
                                                comboTypesList.setModel(jcomboModel);
                                                comboTypesList.setSelectedItem(groupTrimed);
                                                putNamesOnList(groupTrimed);
                                                setButtonColor(addPart);
                                                groupNow = groupTrimed;
                                                nameField.setText("");
                                                caseField.setText("");
                                            }
                                        }
                                    } else {
                                        String tid = (String) JOptionPane.showInputDialog(null, "Välj schema för " + nameF + "\n", "Schema",
                                                JOptionPane.PLAIN_MESSAGE,
                                                null,
                                                registered,
                                                registered[0]);

                                        if ((tid != null) && (tid.length() > 0)) {
                                            if (!excel.addParticipant(nameF, groupTrimed, caseTrimed, startDateF, endDateF, tid)) {
                                                if (tid.equals("Deltid") && !GUI.db.tableExist(caseTrimed)) {
                                                    JOptionPane.showMessageDialog(null, "Du valde ett deltidsschema.\n"
                                                            + "Om deltagaren är här på FM, EM och ibland Heltid så ska man skapa\n"
                                                            + "ett individuellt schema istället ('Individuellt schema' - knappen)", "Meddelande", JOptionPane.DEFAULT_OPTION);
                                                }
                                                putNamesOnList(groupNow);
                                                setButtonColor(addPart);
                                                nameField.setText("");
                                                caseField.setText("");
                                            }
                                        }
                                    }
                                }
                                //om filen inte är öppen fortsätt
                            } else {
                                JOptionPane.showMessageDialog(null, "Samma namn finns redan i " + groupTrimed + ".\nLägg till ett tecken så att man kan skilja på deltagarna.", "Meddelande", JOptionPane.DEFAULT_OPTION);
                            }
                        } else {
                            JOptionPane.showMessageDialog(null, "Deltagare med samma ärendenummer (" + caseF + ") finns redan i grupp " + excel.getGroupByCase(caseF) + "\n", "Meddelande", JOptionPane.DEFAULT_OPTION);
                        }
                    } else {
                        JOptionPane.showMessageDialog(null, "Ange datumet i formatet ÅÅÅÅ-MM-DD", "Meddelande", JOptionPane.DEFAULT_OPTION);
                    }
                } else {
                    JOptionPane.showMessageDialog(null, "Fyll i alla fält", "Meddelande", JOptionPane.DEFAULT_OPTION);
                }
            } catch (IOException ex) {
                JOptionPane.showMessageDialog(null, "Ett fel har inträffat (fel:lägg till deltagare)", "Meddelande", JOptionPane.DEFAULT_OPTION);
            }
        });

        JButton deleteParticipant = new JButton("Ta bort deltagare");
        deleteParticipant.addActionListener((ActionEvent e) -> {
            String checkString;
            if (hashMap.isEmpty()) {
                JOptionPane.showMessageDialog(null, "Ingen grupp är vald", "Meddelande", JOptionPane.DEFAULT_OPTION);
            } else {
                if (model.getCheckedNames().length == 0) {
                    JOptionPane.showMessageDialog(null, "Välj minst en deltagare", "Meddelande", JOptionPane.DEFAULT_OPTION);
                } else {
                    int reply = JOptionPane.showConfirmDialog(null, "Vill du ta bort deltagare från listan?\n\n"
                            + "(OBS! Tas bort permanent!)\n\n", "Säkerhetsfråga", JOptionPane.YES_NO_OPTION, JOptionPane.PLAIN_MESSAGE);
                    if (reply == JOptionPane.YES_OPTION) {
                        if (excel.deleteParticipant(model.getCheckedNames(), groupNow, db)) {
                            grupper = excel.getGroups();
                            jcomboModel = new DefaultComboBoxModel(grupper);
                            comboTypesList.setModel(jcomboModel);
                            Object obj1 = comboTypesList.getSelectedItem();
                            checkString = obj1.toString();
                            if (grupper.length > 0 && !checkString.equals("Excelen är tom")) {
                                boolean arrayEquality = Arrays.equals(grupperInnan, grupper);
                                if (!arrayEquality) {
                                    groupNow = obj1.toString();
                                    grupperInnan = grupper.clone();
                                }
                                comboTypesList.setSelectedItem(groupNow);
                                putNamesOnList(groupNow);
                                setButtonColor(deleteParticipant);
                            } else {
                                putNamesOnList(" ");
                                numberOfPa.setText(" ");
                                JOptionPane.showMessageDialog(null, "Excelfilen är nu tom", "Meddelande", JOptionPane.DEFAULT_OPTION);
                            }
                        }
                    }
                }
            }
        });

        JButton changeGroup = new JButton("Flytta deltagare");
        changeGroup.addActionListener((ActionEvent e) -> {
            int dayOfMonth = cal.get(Calendar.DAY_OF_MONTH);
            boolean wrongGroup = true;
            nullGroup = false;
            if (hashMap.isEmpty()) {
                JOptionPane.showMessageDialog(null, "Ingen grupp är vald", "Meddelande", JOptionPane.DEFAULT_OPTION);
            } else {
                if (model.getCheckedNames().length == 0) {
                    JOptionPane.showMessageDialog(null, "Börja med att markera en deltagare", "Meddelande", JOptionPane.DEFAULT_OPTION);
                } else if (model.getCheckedNames().length > 1) {
                    JOptionPane.showMessageDialog(null, "Markera en deltagare i taget", "Meddelande", JOptionPane.DEFAULT_OPTION);
                } else {
                    while (wrongGroup && nullGroup == false) {
                        try {
                            if (changeGroup(model.getCheckedNames(), dayOfMonth, groupNow)) {
                                grupper = excel.getGroups();
                                boolean arrayEquality = Arrays.equals(grupperInnan, grupper);
                                jcomboModel = new DefaultComboBoxModel(grupper);
                                comboTypesList.setModel(jcomboModel);
                                //om alla deltagare från gruppen flyttats så försvinner gruppen.
                                if (!arrayEquality) {
                                    comboTypesList.setSelectedItem(groupTrimed);
                                    putNamesOnList(groupTrimed);
                                    groupNow = groupTrimed;
                                    grupperInnan = grupper.clone();
                                } else {
                                    comboTypesList.setSelectedItem(groupNow);
                                    putNamesOnList(groupNow);
                                }

                                setButtonColor(changeGroup);
                                wrongGroup = false;
                            }
                        } catch (IOException ex) {
                            JOptionPane.showMessageDialog(null, "Ett fel har inträffat (fel:comboModel)", "Meddelande", JOptionPane.DEFAULT_OPTION);
                        }
                    }
                }
            }
        });

        JButton changeAbsence = new JButton("Ändra / Återställ närvaro");
        changeAbsence.setToolTipText("Man kan markera flera deltagare");
        changeAbsence.addActionListener(new ActionListener() {
            String choosed = "";
            String newAbcence = " ";
            String showConf = "Ändrat";

            @Override
            public void actionPerformed(ActionEvent e) {

                if (hashMap.isEmpty()) {
                    JOptionPane.showMessageDialog(null, "Ingen grupp är vald", "Meddelande", JOptionPane.DEFAULT_OPTION);
                } else {
                    if (model.getCheckedNames().length == 0) {
                        if(nbrOfP > 1) {
                            JOptionPane.showMessageDialog(null, "\nBörja med att markera en eller flera deltagare\n\n"
                                    + "(Man kan markera flera deltagare för att återställa närvaro)\n\n", "Meddelande", JOptionPane.DEFAULT_OPTION);
                        }
                        else if(nbrOfP == 1) {
                            JOptionPane.showMessageDialog(null, "\nBörja med att markera en deltagare\n\n", "Meddelande", JOptionPane.DEFAULT_OPTION);
                        }
                    } else if (model.getCheckedNames().length > 1) {
                        int reply = JOptionPane.showConfirmDialog(null, "Vill du återställa närvaro för de markerade deltagarna?\n"
                                + "(OBS! Går ej att ångra!)", "Fråga", JOptionPane.YES_NO_OPTION, JOptionPane.PLAIN_MESSAGE);

                        if (reply == JOptionPane.YES_OPTION) {
                            String dag = (String) JOptionPane.showInputDialog(null, "Välj dag", "Återställ närvaro", JOptionPane.PLAIN_MESSAGE,
                                    null,
                                    day,
                                    day[0]);
                            if ((dag != null) && (dag.length() > 0)) {
                                showConf = "Närvaro återställd";
                                excel.changeAbsence(model.getCheckedNames(), dag, " ");
                                JOptionPane.showMessageDialog(null, showConf, "Meddelande", JOptionPane.DEFAULT_OPTION);
                                setButtonColor(changeAbsence);
                                model.uncheckAllNames();
                            }
                        }
                    } else {
                        String dag = (String) JOptionPane.showInputDialog(null, "Välj dag", model.getCheckedNames()[0],
                                JOptionPane.PLAIN_MESSAGE,
                                null,
                                day,
                                day[0]);
                        if ((dag != null) && (dag.length() > 0)) {
                            String absence = (String) JOptionPane.showInputDialog(null, "Välj närvaro / frånvaro", model.getCheckedNames()[0] + ", " + "dag" + " " + dag,
                                    JOptionPane.PLAIN_MESSAGE,
                                    null,
                                    changeAbsenceOpt,
                                    changeAbsenceOpt[0]);
                            if (absence != null && dag.length() > 0 && absence.length() > 0) {

                                if (absence.equals("0.15") || absence.equals("0.30")
                                        || absence.equals("0.45") || absence.equals("1") || absence.equals("1.15")
                                        || absence.equals("1.30") || absence.equals("1.45") || absence.equals("2")
                                        || absence.equals("2.15") || absence.equals("2.30") || absence.equals("2.45")
                                        || absence.equals("3") || absence.equals("3.15") || absence.equals("3.30")
                                        || absence.equals("3.45") || absence.equals("4") || absence.equals("4.15")
                                        || absence.equals("4.30") || absence.equals("4.45") || absence.equals("5")
                                        || absence.equals("5.15") || absence.equals("5.30") || absence.equals("5.45")
                                        || absence.equals("6") || absence.equals("6.15") || absence.equals("6.30")
                                        || absence.equals("6.45") || absence.equals("7") || absence.equals("7.15")
                                        || absence.equals("7.30") || absence.equals("7.45") || absence.equals("8")) {
                                    newAbcence = "x-" + absence;
                                    showConf = "Närvaro ändrad";
                                    
                                } else if (absence.equals("Återställ")) {
                                    newAbcence = " ";
                                    showConf = "Närvaro återställd";
                                } else if (!absence.equals("Återställ")){
                                    newAbcence = absence;
                                    showConf = "Närvaro ändrad";
                                }
                                excel.changeAbsence(model.getCheckedNames(), dag, newAbcence.toLowerCase());
                                JOptionPane.showMessageDialog(null, showConf, "Meddelande", JOptionPane.DEFAULT_OPTION);
                                setButtonColor(changeAbsence);
                                model.uncheckAllNames();
                            }
                        } else {

                        }
                    }
                }
            }
        }
        );

        /*
        //används inte för tillfället
        JButton helpButton = new JButton("Hjälp");

        helpButton.addActionListener((ActionEvent e) -> {
            //contactInfo();
            OpenFoldersDocuments m = new OpenFoldersDocuments();
            m.help();
            setButtonColor(helpButton);
        });
         */
        JButton createAttendanceList = new JButton("Skapa dagens närvarolista");
        createAttendanceList.addActionListener((ActionEvent e) -> {
            if (!groupNow.equals("Excelen är tom")) {
                String todaysDay = cal.getDisplayName(Calendar.DAY_OF_WEEK, Calendar.LONG, Locale.ENGLISH).toLowerCase();
                String tid = "";
                Boolean exists = true;
                //todaysDay = "friday";//markera bort när jar-filen ska fixas
                String fileName = System.getProperty("user.home") + "\\Desktop\\Attendance Tool\\Grupplistor\\" + groupNow + ".docx";
                File f = new File(fileName);
                if (!todaysDay.equals("saturday") && !todaysDay.equals("sunday")) {
                    if (!f.exists() && !f.isDirectory()) {
                        exists = false;
                        tid = (String) JOptionPane.showInputDialog(null, "Ange om gruppen " + groupNow + " är en Heltids, FM eller EM grupp. \n"
                                + "(Om gruppen är blandad, välj Heltid)", "Skapa närvarolista",
                                JOptionPane.PLAIN_MESSAGE,
                                null,
                                whichTime,
                                whichTime[0]);
                    }
                    if (tid != null) {
                        wm = new AttendanceWordList(groupNow, tid, exists);
                        wm.addScheduleToWord(groupNow, hashMap, endTimePartic, todaysDay);//hashMap fylls med namn när de ska läggas i GUI. Se metoden putNamesOnList.
                        setButtonColor(createAttendanceList);
                    }
                } else {
                    JOptionPane.showMessageDialog(null, "Det går inte att skapa närvarolista på en helgdag", "Meddelande", JOptionPane.DEFAULT_OPTION);
                }

            } else {
                JOptionPane.showMessageDialog(null, "Närvarolista kan inte skapas\n"
                        + "eftersom excelen är tom", "Meddelande", JOptionPane.DEFAULT_OPTION);
            }
        });

        JButton journal = new JButton("Importera excelfil");
        journal.addActionListener((ActionEvent e) -> {

            if (!copyFile.isVisible()) {
                try {
                    Toolkit tk = Toolkit.getDefaultToolkit();
                    Dimension screenSize = tk.getScreenSize();
                    final int width = screenSize.width;
                    final int height = screenSize.height;

                    copyFile.setSize(290, 120);
                    copyFile.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
                    copyFile.setResizable(false);
                    copyFile.setTitle("Importera excelfil");
                    //copyFile.setIconImage(iconImg.getImage());
                    copyFile.setLocation(width / 3, height / 2);
                    copyFile.setVisible(true);
                    UIManager.setLookAndFeel(previousLF);
                    UIManager.getLookAndFeelDefaults().put("TableHeader.font", new Font("Arial", Font.PLAIN, 14));// ändra text i kategorin
                    if (UIManager.get("Table.alternateRowColor") == null) {
                        UIManager.put("Table.alternateRowColor", new Color(240, 240, 240));
                    }
                } catch (UnsupportedLookAndFeelException ex) {
                    JOptionPane.showMessageDialog(null, "Ett fel har inträffat (fel: fil import)", "Meddelande", JOptionPane.DEFAULT_OPTION);
                }
            } else if (copyFile != null) {
                copyFile.setExtendedState(JFrame.NORMAL);
                copyFile.setAlwaysOnTop(true);
                copyFile.requestFocus();
                copyFile.setAlwaysOnTop(false);
            }
        });

        JButton createReports = new JButton("Skapa närvarorapport");
        createReports.setToolTipText("Har du valt rätt månad?");
        createReports.addActionListener((ActionEvent e) -> {

            /* Det här är för kontaktuppgifter fönstret. Den behövs inte just nu, men kan behövas senare.
             if (!anord.isEmpty() && !samord.isEmpty() && !aktivitet.isEmpty()
             && !mail2.isEmpty() && !tel2.isEmpty()
             && anord.trim().length() > 0 && samord.trim().length() > 0 && aktivitet.trim().length() > 0
             && mail2.trim().length() > 0 && tel2.trim().length() > 0) {
             */
            if (!hashMap.isEmpty()) {
                if (model.getCheckedNames().length != 0) {
                    excel.getValuesFromExcel(model.getCheckedNames(), groupNow, raportMonth, raportYear);
                    setButtonColor(createReports);

                } else {
                    JOptionPane.showMessageDialog(null, "Välj minst en deltagare", "Meddelande", JOptionPane.DEFAULT_OPTION);
                }
            } else {
                JOptionPane.showMessageDialog(null, "Välj grupp först", "Meddelande", JOptionPane.DEFAULT_OPTION);
            }
            // }
            /*
             else {
             JOptionPane.showMessageDialog(null, "Du måste först ange kontaktuppgifter (Använd knappen 'Kontaktuppgifter')", "Meddelande", JOptionPane.DEFAULT_OPTION);
             }
             */
        });

        JButton saveAttendance = new JButton("Spara närvaro");
        saveAttendance.addActionListener((ActionEvent e) -> {
            Map<String, List<String>> hmap = new HashMap<>();
            String dag = "";
            try {
                String todaysDate;
                if (date.equals("Dagens datum") || date.isEmpty() || date.equals("") || date.equals(" ")) {
                    DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
                    cal = Calendar.getInstance();
                    todaysDate = dateFormat.format(cal.getTime());
                    dag = new SimpleDateFormat("EEEE", Locale.ENGLISH).format(new SimpleDateFormat("yyyy-MM-dd").parse(todaysDate));
                } else {
                    dag = new SimpleDateFormat("EEEE", Locale.ENGLISH).format(new SimpleDateFormat("yyyy-MM-dd").parse(date));
                }

            } catch (ParseException ex) {
                JOptionPane.showMessageDialog(null, "Ett fel har inträffat (fel:Datumformat)", "Meddelande", JOptionPane.DEFAULT_OPTION);
            }
            if (!groupNow.isEmpty()) {
                if (!groupNow.equals("Färre kolumner i excelen")
                        && !groupNow.equals("Fler kolumner i excelen")) {
                    if (!groupNow.equals("Excelfil saknas") && !groupNow.equals("Kolumn fel") && !groupNow.equals("Fel i excelen") && !groupNow.equals("Excelen är tom")) {
                        if (!dag.equals("Saturday") && !dag.equals("Sunday")) {

                            try {
                                canceled = false;
                                hmap = model.getInfoFromCheckBoxes(groupNow, dag);//hämta en map med närvaro/frånvaro
                            } catch (IOException ex) {
                                JOptionPane.showMessageDialog(null, "Ett fel har inträffat (fel:getAllInfo)", "Meddelande", JOptionPane.DEFAULT_OPTION);
                            }

                            sparat = false;
                            if (canceled == false) {
                                try {
                                    try {
                                        if (!save.getText().equals("Ej sparat")) {
                                            int reply = JOptionPane.showConfirmDialog(null, "Du har redan sparat (" + date + ").\nVill du spara igen?", "Fråga", JOptionPane.YES_NO_OPTION, JOptionPane.PLAIN_MESSAGE);
                                            if (reply == JOptionPane.YES_OPTION) {
                                                sparat = excel.addValuesToExcel(hmap, date, groupNow);
                                                //JOptionPane.showMessageDialog(null, "Närvaron sparad för, " + date, "Meddelande", JOptionPane.DEFAULT_OPTION);
                                                save.setText("Ej sparat");
                                            } else {
                                                sparat = true;
                                            }
                                        } else {
                                            sparat = excel.addValuesToExcel(hmap, date, groupNow);
                                            //JOptionPane.showMessageDialog(null, "Närvaron sparad för, " + date, "Meddelande", JOptionPane.DEFAULT_OPTION);
                                        }
                                    } catch (FileNotFoundException | ParseException ex) {
                                        JOptionPane.showMessageDialog(null, "Ett fel har inträffat (fel:null text)", "Meddelande", JOptionPane.DEFAULT_OPTION);
                                    }
                                } catch (IOException ex) {
                                    JOptionPane.showMessageDialog(null, "Ett fel har inträffat (fel:add value to excel)", "Meddelande", JOptionPane.DEFAULT_OPTION);
                                }

                                if (sparat == true) {
                                    if (save.getText().equals("Ej sparat") && !tom.equals("tom")) {
                                        if (date.isEmpty()) {
                                            date = "Dagens datum";
                                        }
                                        save.setText("(" + date + ")" + " sparat");
                                        save.setForeground(new Color(0, 153, 76));

                                        setButtonColor(saveAttendance);
                                        clearEverything(groupNow);
                                    }
                                    tom = " ";
                                }
                            }
                        } else {
                            JOptionPane.showMessageDialog(null, "Det går inte att spara närvaro på en helgdag", "Meddelande", JOptionPane.DEFAULT_OPTION);
                        }
                    } else {
                        switch (groupNow) {
                            case "Excelfil saknas":
                                JOptionPane.showMessageDialog(null, "Importera en excelfil först", "Meddelande", JOptionPane.DEFAULT_OPTION);
                                break;
                            case "Kolumn fel":
                                JOptionPane.showMessageDialog(null, "Rätta till kolumnerna först", "Meddelande", JOptionPane.DEFAULT_OPTION);
                                break;
                            case "Fel i excelen":
                                JOptionPane.showMessageDialog(null, "Rätta till excelen först", "Meddelande", JOptionPane.DEFAULT_OPTION);
                                break;
                            case "Excelen är tom":
                                JOptionPane.showMessageDialog(null, "Excelfilen är tom.\n"
                                        + "Börja med att först lägga till deltagare", "Meddelande", JOptionPane.DEFAULT_OPTION);
                                break;
                        }
                    }
                } else {
                    JOptionPane.showMessageDialog(null, "Det är fel med kolumnerna i excelfilen", "Meddelande", JOptionPane.DEFAULT_OPTION);
                }
            } else {
                JOptionPane.showMessageDialog(null, "Ingen grupp är vald", "Meddelande", JOptionPane.DEFAULT_OPTION);
            }
        });

        checkAllButton = new JButton("Markera alla");
        checkAllButton.addActionListener((ActionEvent e) -> {
            checkUncheck();
        });

        JButton resetExcel = new JButton("Spara och nollställ excelfil");
        resetExcel.addActionListener((ActionEvent e) -> {
            if (!groupNow.equals("Excelen är tom")) {
                int reply = JOptionPane.showConfirmDialog(null, "En närvaro-excel (närvarofil) gäller för en månad, och när den sparas\n"
                        + "döps den till, 'aktuella månaden'-deltagare.xls (t.ex. juni-deltagare.xls)\n\n"
                        + "Har du angett rätt månad?\n\n"
                        + " För att ändra månad:\n"
                        + "  1. Tryck på 'Välj datum'.\n"
                        + "  2. Bläddra till rätt månad.\n"
                        + "  3. Välj ett godtyckligt datum (vilket som).\n\n\n", "Spara och nollställ närvaro-excel", JOptionPane.YES_NO_OPTION, JOptionPane.PLAIN_MESSAGE);
                if (reply == JOptionPane.YES_OPTION) {
                    setButtonColor(resetExcel);
                    CopyFiles cpf = new CopyFiles(raportMonth);
                }
            } else {
                JOptionPane.showMessageDialog(null, "Excelen behöver inte nollställas\n"
                        + "eftersom den redan är tom", "Meddelande", JOptionPane.DEFAULT_OPTION);
            }
        });

        /*
        JButton statistics = new JButton("Deltagarstatistik");
        statistics.addActionListener((ActionEvent e) -> {
            setButtonColor(statistics);
            try {
                excel.getTheStatistic();
            } catch (DateTimeParseException dte) {
                JOptionPane.showMessageDialog(null, "Statistik kunde inte hämtas.\n"
                        + "1. Se till att datum finns i excelen.\n"
                        + "2. Se till att datumen är skrivna i formatet ÅÅÅÅ-MM-DD.\n\n"
                        + "Använd knappen 'Deltagarinfo / Ändra info' för att ändra datum.\n\n\n", "Meddelande", JOptionPane.DEFAULT_OPTION);
            }
        });
         */
        int delay = 2000; //milliseconds
        indScheduleTimer = new Timer(delay, new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent evt) {
                schema.setEnabled(true);
                schema.setText("Individuellt schema");
            }
        });
        indScheduleTimer.setRepeats(false);

        int delay2 = 2000; //milliseconds
        changeInfoTimer = new Timer(delay2, new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent evt) {
                changeButton.setEnabled(true);
                changeButton.setText("Deltagarinfo / Ändra info");
                putNamesOnList(groupNow);
            }
        });
        changeInfoTimer.setRepeats(false);

        changeButton = new JButton("Deltagarinfo / Ändra info");
        changeButton.addActionListener((ActionEvent e) -> {
            if (windowClosed == true) {
                if (!hashMap.isEmpty()) {

                    if (model.getCheckedNames().length != 0) {
                        if (model.getCheckedNames().length > 1) {
                            JOptionPane.showMessageDialog(null, "Välj en deltagare i taget", "Meddelande", JOptionPane.DEFAULT_OPTION);
                        } else {
                            windowClosed = false;
                            createInfoPanel(model.getCheckedNames()[0]);
                            setButtonColor(changeButton);
                        }
                    } else {
                        JOptionPane.showMessageDialog(null, "Börja med att markera en deltagare", "Meddelande", JOptionPane.DEFAULT_OPTION);
                    }
                } else {
                    JOptionPane.showMessageDialog(null, "Ingen grupp är vald", "Meddelande", JOptionPane.DEFAULT_OPTION);
                }
            } else if (infoFrame != null) {
                infoFrame.setExtendedState(JFrame.NORMAL);
                infoFrame.setAlwaysOnTop(true);
                infoFrame.requestFocus();
                infoFrame.setAlwaysOnTop(false);
            }
        });

        schema = new JButton("Individuellt schema");
        schema.addActionListener((ActionEvent e) -> {
            if (windowClosed == true) {
                String[] s = new String[1];
                s[0] = "Skapa schema";
                if (!hashMap.isEmpty()) {
                    if (model.getCheckedNames().length != 0) {
                        if (model.getCheckedNames().length > 1) {
                            JOptionPane.showMessageDialog(null, "Välj en deltagare i taget", "Meddelande", JOptionPane.DEFAULT_OPTION);
                        } else {
                            if (db.tableExist(model.getCheckedCaseNumber()[0])) {
                                windowClosed = false;
                                createSchedulePanel(model.getCheckedNames()[0], db.getWeekSchedule(model.getCheckedCaseNumber()[0]), model.getCheckedCaseNumber()[0]);
                                setButtonColor(schema);
                            } else {
                                //int reply = JOptionPane.showConfirmDialog(null, "Vill du skapa ett individuellt schema för " + model.getCheckedNames()[0] + "?\n",
                                //  "Individuellt schema", JOptionPane.YES_NO_OPTION, JOptionPane.PLAIN_MESSAGE);
                                //if (reply == JOptionPane.YES_OPTION) {
                                windowClosed = false;
                                createSchedulePanel(model.getCheckedNames()[0], s, model.getCheckedCaseNumber()[0]);
                                setButtonColor(schema);
                                // }
                            }
                        }
                    } else {
                        JOptionPane.showMessageDialog(null, "Börja med att markera en deltagare", "Meddelande", JOptionPane.DEFAULT_OPTION);
                    }
                } else {
                    JOptionPane.showMessageDialog(null, "Ingen grupp är vald", "Meddelande", JOptionPane.DEFAULT_OPTION);
                }
            } else if (scheduleFrame != null) {
                scheduleFrame.setExtendedState(JFrame.NORMAL);
                scheduleFrame.setAlwaysOnTop(true);
                scheduleFrame.requestFocus();
                scheduleFrame.setAlwaysOnTop(false);
            }
        });
        /* //Används inte för tillfället då den inte är uppdaterad
        JButton certificate = new JButton("Skapa kursintyg");
        certificate.addActionListener((ActionEvent e) -> {
            if (!chef.isEmpty() && chef.trim().length() > 0) {
                int ok = -1;
                try {
                    if (!hashMap.isEmpty()) {

                        if (model.getCheckedNames().length != 0) {
                            if (model.getCheckedNames().length > 1) {
                                JOptionPane.showMessageDialog(null, "Välj en deltagare i taget", "Meddelande", JOptionPane.DEFAULT_OPTION);
                            } else {
                                while (ok == -1) {
                                    int option = JOptionPane.showConfirmDialog(null, certInfo, "Skapa intyg för " + model.getCheckedNames()[0], JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);
                                    if (option == JOptionPane.OK_OPTION) {
                                        if (teacherName.getText().trim().length() > 0 && aboutParticipant.getText().trim().length() > 0 && city.getText().trim().length() > 0) {
                                            ok = 2;
                                            setButtonColor(certificate);

                                            for (int i = 0; i < model.getCheckedNames().length; i++) {
                                                Certificate.createCertificate(model.getCheckedNames()[i], teacherName.getText(), groupNow, aboutParticipant.getText(), city.getText());
                                            }
                                        } else {
                                            JOptionPane.showMessageDialog(null, "Fyll i alla fält", "Meddelande", JOptionPane.DEFAULT_OPTION);
                                        }
                                    } else {
                                        ok = 2;
                                    }
                                }
                            }
                        } else {
                            JOptionPane.showMessageDialog(null, "Börja med att markera en deltagare", "Meddelande", JOptionPane.DEFAULT_OPTION);
                        }

                    } else {
                        JOptionPane.showMessageDialog(null, "Välj grupp först", "Meddelande", JOptionPane.DEFAULT_OPTION);
                    }
                } catch (IOException ex) {
                    JOptionPane.showMessageDialog(null, "PDF-dokumentet kunde inte hittas\n"
                            + "Säkerställ att den är placerad i 'Attendance Tool' mappen", "Meddelande", JOptionPane.DEFAULT_OPTION);
                }
            } else {
                JOptionPane.showMessageDialog(null, "Du måste först ange vem som är regionchef", "Meddelande", JOptionPane.DEFAULT_OPTION);
            }
        });
        certificate.setEnabled(false);//avaktivera. aktuell för gammalt intyg
         */
        addPart.setFont(new Font(font, Font.PLAIN, buttonTextSize));
        save.setFont(new Font(font, Font.PLAIN, buttonTextSize));
        saveAttendance.setFont(new Font(font, Font.PLAIN, buttonTextSize));
        numberOfPa.setFont(new Font(font, Font.PLAIN, buttonTextSize));
        selectDate.setFont(new Font(font, Font.PLAIN, buttonTextSize));
        deleteParticipant.setFont(new Font(font, Font.PLAIN, buttonTextSize));
        changeGroup.setFont(new Font(font, Font.PLAIN, buttonTextSize));
        checkAllButton.setFont(new Font(font, Font.PLAIN, buttonTextSize));
        searchParticipant.setFont(new Font(font, Font.PLAIN, buttonTextSize));
        absenceReport.setFont(new Font(font, Font.PLAIN, buttonTextSize));
        createReports.setFont(new Font(font, Font.PLAIN, buttonTextSize));

        //helpButton.setFont(new Font(font, Font.PLAIN, buttonTextSize));
        //certificate.setFont(new Font(font, Font.PLAIN, buttonTextSize));
        resetExcel.setFont(new Font(font, Font.PLAIN, buttonTextSize));
        //statistics.setFont(new Font(font, Font.PLAIN, buttonTextSize));
        changeAbsence.setFont(new Font(font, Font.PLAIN, buttonTextSize));
        createAttendanceList.setFont(new Font(font, Font.PLAIN, buttonTextSize));
        journal.setFont(new Font(font, Font.PLAIN, buttonTextSize));
        schema.setFont(new Font(font, Font.PLAIN, buttonTextSize));
        changeButton.setFont(new Font(font, Font.PLAIN, buttonTextSize));

        addPart.setPreferredSize(new Dimension(110, 25));
        saveAttendance.setPreferredSize(new Dimension(130, 25));

        imagePanel.setBackground(new Color(0xEDFFE1));
        btnPanel.setBackground(new Color(0xEDFFE1));
        selectDate.setPreferredSize(new Dimension(120, 23));

        btnPanel.add(journal);//importera excelfil
        btnPanel.add(checkAllButton);
        btnPanel.add(changeAbsence);
        btnPanel.add(resetExcel);
        //btnPanel.add(statistics);
        btnPanel.add(changeGroup);
        //btnPanel.add(certificate);
        btnPanel.add(createReports);
        btnPanel.add(schema);//individuellt schema
        btnPanel.add(searchParticipant);
        btnPanel.add(createAttendanceList);
        btnPanel.add(changeButton);//ändra info 
        btnPanel.add(deleteParticipant);
        btnPanel.add(absenceReport);
        //btnPanel.add(helpButton);

        //spara och nollställ excel
        //btnPanel.add(dagJournal);
        text = new JTextField(8);
        JLabel text1 = new JLabel();
        JLabel text2 = new JLabel();
        JLabel text3 = new JLabel();
        JLabel text4 = new JLabel();
        JLabel text5 = new JLabel();
        JLabel text6 = new JLabel();
        JLabel text7 = new JLabel();
        JLabel text8 = new JLabel();
        JLabel text9 = new JLabel();
        JLabel text11 = new JLabel();

        text1.setText("__________________________________");
        text2.setText("__________________________________");
        text3.setText("__________________________________");
        text4.setText("_____________________________");
        text5.setText("__________________________________");
        text6.setText("______________");
        text7.setText("_____________________________");
        text8.setText("_____________________________");
        text9.setText("_____________________________");
        text11.setText("_____________________________");
        save.setText(" ");
        text4.setText("_____________________________");
        text1.setForeground(new Color(0xEA565A));
        text2.setForeground(new Color(0xFFAA00));
        text5.setForeground(new Color(0xFFAA00));
        text6.setForeground(new Color(0xEDFFE1));
        save.setForeground(Color.red);
        text3.setForeground(new Color(0xFFAA00));
        text4.setForeground(new Color(0xEDFFE1));
        text7.setForeground(new Color(0xEDFFE1));
        text8.setForeground(new Color(0xEDFFE1));
        text9.setForeground(new Color(0xEDFFE1));
        text11.setForeground(new Color(0xEDFFE1));
        JLabel label = new JLabel();

        //label.setIcon(new ImageIcon(getClass().getResource(img)));// bild

        imagePanel.add(label);
        imagePanel.add(text1);
        imagePanel.add(selectDate);
        imagePanel.add(text);
        imagePanel.add(text2);
        imagePanel.add(text9);

        imagePanel.add(comboTypesList);
        imagePanel.add(numberOfPa);
        imagePanel.add(text3);

        imagePanel.add(text11);
        imagePanel.add(save);
        imagePanel.add(text4);
        imagePanel.add(saveAttendance);
        imagePanel.add(text5);

        imagePanel.add(nameField);
        imagePanel.add(groupField);
        imagePanel.add(caseField);
        imagePanel.add(startDate);
        imagePanel.add(endDate);

        imagePanel.add(text6);
        imagePanel.add(addPart);

        setLayout(new BorderLayout(1, 2));

        add(new JScrollPane(table), BorderLayout.WEST);
        add(btnPanel, BorderLayout.SOUTH);
        add(imagePanel, BorderLayout.EAST);

        if (ExcelHandler.groups.length > 0 && !groupNow.equals("Excelfil saknas")) {
            putNamesOnList(obj.toString());
        }

        cal = Calendar.getInstance();
        text.setText(dateF.format(cal.getTime()));
        date = text.getText();
        priviousDate = date;
    }

    public static boolean nyGrupp(String grupper[], String groupTrimmed) {
        for (int i = 0; i < grupper.length; i++) {
            if (grupper[i].equals(groupTrimmed)) {
                return false;
            }
        }
        return true;
    }

    /*
    public void contactInfo() {
        JTextField anordnare = new JTextField();
        JTextField samordnare = new JTextField();
        JTextField telefon = new JTextField();
        JTextField mail = new JTextField();
        JTextField aktivitet2 = new JTextField();

        anordnare.setPreferredSize(new Dimension(50, 25));
        samordnare.setPreferredSize(new Dimension(50, 25));
        telefon.setPreferredSize(new Dimension(50, 25));
        mail.setPreferredSize(new Dimension(50, 25));
        aktivitet2.setPreferredSize(new Dimension(50, 25));

        Object[] message = {
            "Informationen nedan används i bl.a.\n"
            + "månadsrapporter.\n\n",
            "Anordnare:", anordnare,
            "Samordnare:", samordnare,
            "Telefonnummer:", telefon,
            "Email:", mail,
            "Aktivitet:", aktivitet2,
            " ", " "
        };
        boolean empty = true;
        int option3;
        anordnare.setText(prefs.get("anordString", ""));
        samordnare.setText(prefs.get("samordString", ""));
        telefon.setText(prefs.get("telString", ""));
        mail.setText(prefs.get("mailString", ""));
        aktivitet2.setText(prefs.get("aktivitetString", ""));

        while (empty == true) {
            option3 = JOptionPane.showConfirmDialog(null, message, "Kontaktuppgifter", JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);

            if (option3 == JOptionPane.OK_OPTION) {
                if (!anordnare.getText().isEmpty() && !samordnare.getText().isEmpty()
                        && !mail.getText().isEmpty() && !telefon.getText().isEmpty() && !aktivitet2.getText().isEmpty()
                        && anordnare.getText().trim().length() > 0 && aktivitet2.getText().trim().length() > 0
                        && samordnare.getText().trim().length() > 0 && mail.getText().trim().length() > 0 && telefon.getText().trim().length() > 0) {

                    prefs.put("anordString", anordnare.getText());
                    prefs.put("samordString", samordnare.getText());
                    prefs.put("mailString", mail.getText());
                    prefs.put("telString", telefon.getText());
                    prefs.put("aktivitetString", aktivitet2.getText());

                    anord = prefs.get("anordString", "");
                    chef = prefs.get("samordString", "");
                    mail2 = prefs.get("mailString", "");
                    tel2 = prefs.get("telString", "");
                    aktivitet = prefs.get("aktivitetString", "");

                    empty = false;
                } else {
                    JOptionPane.showConfirmDialog(null, "Fyll i alla fält", "Meddelande", JOptionPane.DEFAULT_OPTION, JOptionPane.PLAIN_MESSAGE);
                }
            } else {
                empty = false;
            }
        }
    }
     */
    public void createSchedulePanel(String name, String[] weekSchedule, String caseNbr) {

        scheduleFrame = new JFrame(name);
        //scheduleFrame.setBounds(100, 300, 800, 600);
        scheduleFrame.setSize(618, 410);
        scheduleFrame.setResizable(false);
        final JPanel panel = new JPanel();
        final JPanel btnPanel = new JPanel();
        JButton btnOk, btnCancel, btnDelete;

        JTextArea field = new JTextArea(4, 3);
        field.setPreferredSize(new Dimension(300, 10));
        Border border = BorderFactory.createLineBorder(Color.GRAY);
        field.setBorder(BorderFactory.createCompoundBorder(border,
                BorderFactory.createEmptyBorder(10, 10, 10, 10)));
        field.setLineWrap(true);
        field.setWrapStyleWord(true);
        ((AbstractDocument) field.getDocument()).setDocumentFilter(new DocumentFilter() {

            //begränsa antalet bokstäver i kommentarsfältet
            @Override
            public void replace(FilterBypass fb, int offset, int length, String str, AttributeSet attr)
                    throws BadLocationException {
                if (offset < 116) {
                    super.replace(fb, offset, length, str, attr);
                }
            }
        });

        scheduleFrame.addWindowListener(new WindowAdapter() {
            @Override
            public void windowClosing(WindowEvent windowEvent) {
                windowClosed = true;
            }
        });

        btnOk = new JButton("Uppdatera");
        btnOk.setBounds(30, 60, 30, 30);
        btnCancel = new JButton("Avbryt");
        btnCancel.setBounds(70, 60, 80, 30);
        btnDelete = new JButton("Ta bort");
        btnDelete.setBounds(70, 60, 80, 30);

        måndag = new JComboBox<>(scheduleOpt);
        tisdag = new JComboBox<>(scheduleOpt);
        onsdag = new JComboBox<>(scheduleOpt);
        torsdag = new JComboBox<>(scheduleOpt);
        fredag = new JComboBox<>(scheduleOpt);

        måndag.setPreferredSize(new Dimension(117, 25));
        tisdag.setPreferredSize(new Dimension(117, 25));
        onsdag.setPreferredSize(new Dimension(117, 25));
        torsdag.setPreferredSize(new Dimension(117, 25));
        fredag.setPreferredSize(new Dimension(117, 25));

        måndag.setMaximumRowCount(22);
        tisdag.setMaximumRowCount(22);
        onsdag.setMaximumRowCount(22);
        torsdag.setMaximumRowCount(22);
        fredag.setMaximumRowCount(22);

        JLabel title = new JLabel();
        JLabel line = new JLabel();
        JLabel line2 = new JLabel();
        JLabel line3 = new JLabel();

        JLabel line5 = new JLabel();
        JLabel line6 = new JLabel();

        JLabel line8 = new JLabel();
        JLabel textBoxTitle = new JLabel();

        Color color = scheduleFrame.getBackground();
        title.setText("Individuellt schema ( Måndag - Fredag )");

        line.setText("___________________________________________________________________________");
        line2.setText("___________________________________________________________________________");
        line3.setText("___________________________________________________________________________");

        line5.setText("___________________________________________________________________________");
        line6.setText("___________________________________________________________________________");

        line8.setText("___________________________________________________________________________");
        textBoxTitle.setText("   Anteckning:   ");
        line.setForeground(color);
        line2.setForeground(color);
        line3.setForeground(color);

        line5.setForeground(color);
        line6.setForeground(color);

        line8.setForeground(color);

        panel.add(line);
        panel.add(title);
        panel.add(line2);
        panel.add(måndag);
        panel.add(tisdag);
        panel.add(onsdag);
        panel.add(torsdag);
        panel.add(fredag);
        panel.add(line5);
        panel.add(line8);
        panel.add(textBoxTitle);
        panel.add(line3);
        panel.add(field);

        btnPanel.add(btnDelete);
        btnPanel.add(btnCancel);
        btnPanel.add(btnOk);

        //scheduleFrame.getContentPane().add(panel);
        scheduleFrame.getContentPane().add(BorderLayout.SOUTH, btnPanel);
        scheduleFrame.getContentPane().add(BorderLayout.CENTER, panel);
        scheduleFrame.setLocationRelativeTo(null);
        scheduleFrame.setIconImage((imageIcon.getImage()));
        scheduleFrame.setVisible(true);

        if (!db.tableExist(model.getCheckedCaseNumber()[0])) {
            btnDelete.setVisible(false);
        }

        if (!weekSchedule[0].equals("Skapa schema")) {
            måndag.setSelectedItem(weekSchedule[0]);
            tisdag.setSelectedItem(weekSchedule[1]);
            onsdag.setSelectedItem(weekSchedule[2]);
            torsdag.setSelectedItem(weekSchedule[3]);
            fredag.setSelectedItem(weekSchedule[4]);
            field.setText(weekSchedule[5]);

            monAct = weekSchedule[0];
            tueAct = weekSchedule[1];
            wedAct = weekSchedule[2];
            thuAct = weekSchedule[3];
            friAct = weekSchedule[4];
            note = weekSchedule[5];

        } else {
            title.setText("Individuellt schema (Måndag - Fredag)");
            btnOk.setText("Skapa schema");

            monAct = måndag.getItemAt(0).toString();
            tueAct = tisdag.getItemAt(0).toString();
            wedAct = onsdag.getItemAt(0).toString();
            thuAct = torsdag.getItemAt(0).toString();
            friAct = fredag.getItemAt(0).toString();
        }

        måndag.addActionListener((ActionEvent e) -> {
            måndag = (JComboBox) e.getSource();
            monAct = måndag.getItemAt(måndag.getSelectedIndex()).toString();
        }
        );
        tisdag.addActionListener((ActionEvent e) -> {
            tisdag = (JComboBox) e.getSource();
            tueAct = tisdag.getItemAt(tisdag.getSelectedIndex()).toString();
        }
        );
        onsdag.addActionListener((ActionEvent e) -> {
            onsdag = (JComboBox) e.getSource();
            wedAct = onsdag.getItemAt(onsdag.getSelectedIndex()).toString();
        }
        );
        torsdag.addActionListener((ActionEvent e) -> {
            torsdag = (JComboBox) e.getSource();
            thuAct = torsdag.getItemAt(torsdag.getSelectedIndex()).toString();
        }
        );
        fredag.addActionListener((ActionEvent e) -> {
            fredag = (JComboBox) e.getSource();
            friAct = fredag.getItemAt(fredag.getSelectedIndex()).toString();
        }
        );

        /*
         {"HELTID", "EM","FM", "SFI-EM", "FM-SFI", "SFI", "SO-EM", "FM-SO", 
         "SO-SO", "SO", "FM-P", "P-EM", "P", "SO-P", "P-SO", "SFI-P", "P-SFI", "SO-SFI", "SFI-SO",
         "SO-LEDIG", "LEDIG-SO", "SFI-LEDIG", "LEDIG-SFI", "FM-LEDIG", "LEDIG-EM", "LEDIG", "SJUK-SJUK", "SJUK", "FL-FL","FL"};
         */
        måndag.setRenderer(new BasicComboBoxRenderer() {
            @Override
            public Component getListCellRendererComponent(JList list, Object value, int index, boolean isSelected, boolean cellHasFocus) {
                if (isSelected) {
                    setBackground(list.getSelectionBackground());
                    setForeground(list.getSelectionForeground());
                    if (index > -1) {
                        list.setToolTipText(ScheduleOptTipText[index]);
                    }
                } else {
                    setBackground(list.getBackground());
                    setForeground(list.getForeground());
                }
                setFont(list.getFont());
                setText((value == null) ? "" : value.toString());

                return this;
            }
        });
        tisdag.setRenderer(new BasicComboBoxRenderer() {
            @Override
            public Component getListCellRendererComponent(JList list, Object value, int index, boolean isSelected, boolean cellHasFocus) {
                if (isSelected) {
                    setBackground(list.getSelectionBackground());
                    setForeground(list.getSelectionForeground());
                    if (index > -1) {
                        list.setToolTipText(ScheduleOptTipText[index]);
                    }
                } else {
                    setBackground(list.getBackground());
                    setForeground(list.getForeground());
                }
                setFont(list.getFont());
                setText((value == null) ? "" : value.toString());

                return this;
            }
        });
        onsdag.setRenderer(new BasicComboBoxRenderer() {
            @Override
            public Component getListCellRendererComponent(JList list, Object value, int index, boolean isSelected, boolean cellHasFocus) {
                if (isSelected) {
                    setBackground(list.getSelectionBackground());
                    setForeground(list.getSelectionForeground());
                    if (index > -1) {
                        list.setToolTipText(ScheduleOptTipText[index]);
                    }
                } else {
                    setBackground(list.getBackground());
                    setForeground(list.getForeground());
                }
                setFont(list.getFont());
                setText((value == null) ? "" : value.toString());

                return this;
            }
        });
        torsdag.setRenderer(new BasicComboBoxRenderer() {
            @Override
            public Component getListCellRendererComponent(JList list, Object value, int index, boolean isSelected, boolean cellHasFocus) {
                if (isSelected) {
                    setBackground(list.getSelectionBackground());
                    setForeground(list.getSelectionForeground());
                    if (index > -1) {
                        list.setToolTipText(ScheduleOptTipText[index]);
                    }
                } else {
                    setBackground(list.getBackground());
                    setForeground(list.getForeground());
                }
                setFont(list.getFont());
                setText((value == null) ? "" : value.toString());

                return this;
            }
        });
        fredag.setRenderer(new BasicComboBoxRenderer() {
            @Override
            public Component getListCellRendererComponent(JList list, Object value, int index, boolean isSelected, boolean cellHasFocus) {
                if (isSelected) {
                    setBackground(list.getSelectionBackground());
                    setForeground(list.getSelectionForeground());
                    if (index > -1) {
                        list.setToolTipText(ScheduleOptTipText[index]);
                    }
                } else {
                    setBackground(list.getBackground());
                    setForeground(list.getForeground());
                }
                setFont(list.getFont());
                setText((value == null) ? "" : value.toString());

                return this;
            }
        });

        btnOk.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent event) {
                try {
                    if (!excel.checkDubleCaseNbr(name.trim(), caseNbr)) {
                        note = field.getText();
                        schema.setEnabled(false);
                        schema.setText("Vänta...");
                        if (weekSchedule[0].equals("Skapa schema")) {
                            db.createSchedule(model.getCheckedCaseNumber()[0], monAct, tueAct, wedAct, thuAct, friAct, note);
                            try {
                                if (!monAct.equals("HELTID") || !tueAct.equals("HELTID") || !wedAct.equals("HELTID")
                                        || !thuAct.equals("HELTID") || !friAct.equals("HELTID")) {
                                    excel.setActivity(caseNbr, "Deltid");
                                }
                            } catch (FileNotFoundException | ParseException ex) {
                                JOptionPane.showMessageDialog(null, "Kontrollera excelfilen (deltagare.xls) i mappen ’Attendance Tool’ → ’Dokument’", "Meddelande", JOptionPane.DEFAULT_OPTION);
                            }
                            indScheduleTimer.start();
                        } else {
                            db.updateSchedule(model.getCheckedCaseNumber()[0], monAct, tueAct, wedAct, thuAct, friAct, note);
                            indScheduleTimer.start();
                        }
                        windowClosed = true;
                        scheduleFrame.setVisible(false);
                        putNamesOnList(groupNow);
                    }
                } catch (IOException ex) {
                    JOptionPane.showMessageDialog(null, "Excelfilen kunde inte läsas!", "Meddelande", JOptionPane.DEFAULT_OPTION);
                }
            }
        });

        btnCancel.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent event) {
                windowClosed = true;
                scheduleFrame.dispose();
            }
        });

        btnDelete.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent event) {
                String tid;
                int reply = JOptionPane.showConfirmDialog(null, "Vill du ta bort det individuella schemat?\n\n"
                        + "(OBS! Tas bort permanent!)\n\n", name, JOptionPane.YES_NO_OPTION, JOptionPane.PLAIN_MESSAGE);
                if (reply == JOptionPane.YES_OPTION) {

                    //när det individuella schemat har tagits bort så behövs det en allmän schema (Heltid, FM eller EM).
                    tid = (String) JOptionPane.showInputDialog(null, "Ange det nya schemat", "Schema",
                            JOptionPane.PLAIN_MESSAGE,
                            null,
                            registered,
                            registered[0]);
                    if (tid != null) {
                        if (tid.equals("Deltid")) {
                            JOptionPane.showMessageDialog(null, "Du valde ett deltidsschema.\n"
                                    + "Om deltagaren är här på FM, EM och ibland Heltid\n"
                                    + "så bör deltagaren ha ett individuellt schema", "Meddelande", JOptionPane.DEFAULT_OPTION);
                        }
                        try {
                            if (excel.setActivity(model.getCheckedCaseNumber()[0], tid)) {
                                schema.setEnabled(false);
                                schema.setText("Vänta...");
                                db.deleteRow(model.getCheckedCaseNumber()[0]);
                                windowClosed = true;
                                indScheduleTimer.start();
                                scheduleFrame.setVisible(false);
                            }

                        } catch (IOException ex) {
                            JOptionPane.showMessageDialog(null, "Ett fel har inträffat (fel:databas-excel)", "Meddelande", JOptionPane.DEFAULT_OPTION);
                        } catch (ParseException ex) {
                            JOptionPane.showMessageDialog(null, "Ett fel har inträffat (fel:databas-excel)", "Meddelande", JOptionPane.DEFAULT_OPTION);
                        }
                    }
                    putNamesOnList(groupNow);
                }
            }
        });
    }

    public void createInfoPanel(String name) {
        String[] info = excel.getParticipantInfo(name, groupNow, db);

        String[] labels = {"Namn: ", "Startdatum: ", "Slutdatum: ", "Schema: ", "Ärendenummer:", "Språk:", "Grupp:"};
        int numPairs = labels.length;

        //Create and populate the panel.
        JPanel p = new JPanel(new SpringLayout());

        infoFrame = new JFrame("Deltagarinfo / Ändra info");
        infoFrame.setBounds(100, 100, 800, 600);
        infoFrame.setSize(300, 280);
        infoFrame.setResizable(false);
        final JPanel btnPanel = new JPanel();
        JButton btnCancel;

        infoFrame.addWindowListener(new WindowAdapter() {
            @Override
            public void windowClosing(WindowEvent windowEvent) {
                windowClosed = true;
            }
        });

        saveInfo = new JButton("Spara");
        saveInfo.setBounds(30, 60, 30, 30);
        btnCancel = new JButton("Avbryt");
        btnCancel.setBounds(70, 60, 80, 30);
        infoFrame.setVisible(true);

        JLabel nameLabel = new JLabel(labels[0], JLabel.TRAILING);
        JLabel startLabel = new JLabel(labels[1], JLabel.TRAILING);
        JLabel endLabel = new JLabel(labels[2], JLabel.TRAILING);
        JLabel timeLabel = new JLabel(labels[3], JLabel.TRAILING);
        JLabel caseLabel = new JLabel(labels[4], JLabel.TRAILING);
        JLabel languageLabel = new JLabel(labels[5], JLabel.TRAILING);
        JLabel groupLabel = new JLabel(labels[6], JLabel.TRAILING);

        p.add(nameLabel);
        JTextField nameF = new JTextField();
        nameLabel.setLabelFor(nameF);
        p.add(nameF);

        p.add(caseLabel);
        JTextField caseF = new JTextField(10);
        caseLabel.setLabelFor(caseF);
        p.add(caseF);

        p.add(startLabel);
        JTextField startF = new JTextField(10);
        startLabel.setLabelFor(startF);
        p.add(startF);

        p.add(endLabel);
        JTextField endF = new JTextField(10);
        endLabel.setLabelFor(endF);
        p.add(endF);

        p.add(timeLabel);
        JTextField timeF = new JTextField(10);
        timeLabel.setLabelFor(timeF);
        p.add(timeF);

        p.add(groupLabel);
        JTextField groupF = new JTextField(10);
        groupLabel.setLabelFor(groupF);
        p.add(groupF);

        p.add(languageLabel);
        JTextField languageF = new JTextField(10);
        languageLabel.setLabelFor(languageF);
        p.add(languageF);

        SpringUtilities.makeCompactGrid(p,
                numPairs, 2, //rows, cols
                6, 6, //initX, initY
                6, 6);       //xPad, yPad

        JLabel title = new JLabel();
 

        Color color = infoFrame.getBackground();
        title.setText("Redigera information för " + name);

        btnPanel.add(btnCancel);
        btnPanel.add(saveInfo);

        infoFrame.getContentPane().add(p);
        infoFrame.getContentPane().add(BorderLayout.SOUTH, btnPanel);
        infoFrame.setLocationRelativeTo(null);
        infoFrame.setIconImage((imageIcon.getImage()));

        nameF.setText(info[0]);
        startF.setText(info[1]);
        endF.setText(info[2]);
        timeF.setText(info[3]);
        caseF.setText(info[4]);
        languageF.setText(info[5]);
        groupF.setText(info[6]);

        String prevCase = info[4];

        saveInfo.setEnabled(false);

        nameF.addKeyListener(new KeyAdapter() {
            @Override
            public void keyPressed(KeyEvent ke) {
                if (saveInfo.isEnabled() == false) {
                    saveInfo.setEnabled(true);
                }
            }
        });
        startF.addKeyListener(new KeyAdapter() {
            @Override
            public void keyPressed(KeyEvent ke) {
                if (saveInfo.isEnabled() == false) {
                    saveInfo.setEnabled(true);
                }
            }
        });
        endF.addKeyListener(new KeyAdapter() {
            @Override
            public void keyPressed(KeyEvent ke) {
                if (saveInfo.isEnabled() == false) {
                    saveInfo.setEnabled(true);
                }
            }
        });
        timeF.addKeyListener(new KeyAdapter() {
            @Override
            public void keyPressed(KeyEvent ke) {
                if (saveInfo.isEnabled() == false) {
                    saveInfo.setEnabled(true);
                }
            }
        });
        caseF.addKeyListener(new KeyAdapter() {
            @Override
            public void keyPressed(KeyEvent ke) {
                if (saveInfo.isEnabled() == false) {
                    saveInfo.setEnabled(true);
                }
            }
        });
        languageF.addKeyListener(new KeyAdapter() {
            @Override
            public void keyPressed(KeyEvent ke) {
                if (saveInfo.isEnabled() == false) {
                    saveInfo.setEnabled(true);
                }
            }
        });

        groupF.addKeyListener(new KeyAdapter() {
            @Override
            public void keyPressed(KeyEvent ke) {
                if (saveInfo.isEnabled() == false) {
                    saveInfo.setEnabled(true);
                }
            }
        });

        saveInfo.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent event) {
                Object obj1 = comboTypesList.getSelectedItem();
                boolean arrayEquality;
                if (nameF.getText().isEmpty() == false && timeF.getText().isEmpty() == false
                        && caseF.getText().isEmpty() == false && startF.getText().isEmpty() == false
                        && endF.getText().isEmpty() == false && languageF.getText().isEmpty() == false
                        && groupF.getText().isEmpty() == false && nameF.getText().trim().length() > 0
                        && timeF.getText().trim().length() > 0 && caseF.getText().trim().length() > 0
                        && startF.getText().trim().length() > 0 && endF.getText().trim().length() > 0
                        && languageF.getText().trim().length() > 0 && groupF.getText().trim().length() > 0) {

                    if (timeF.getText().trim().toLowerCase().equals("fm") || timeF.getText().trim().toLowerCase().equals("em")
                            || timeF.getText().trim().toLowerCase().equals("heltid") || timeF.getText().trim().toLowerCase().equals("fm/em")
                            || timeF.getText().trim().toLowerCase().equals("deltid") || timeF.getText().trim().toLowerCase().equals("individuellt schema")) {
                        if ((isValidFormat("yyyy-MM-dd", startF.getText().trim()) && (isValidFormat("yyyy-MM-dd", endF.getText().trim())))) {

                            if (caseF.getText().trim().length() < 8 || caseF.getText().trim().length() > 8) {
                                int reply = JOptionPane.showConfirmDialog(null, "Ärendenumret brukar vara åtta siffror långt.\n"
                                        + "Skrev du rätt ärendenummer?", "Säkerhetsfråga", JOptionPane.YES_NO_OPTION, JOptionPane.PLAIN_MESSAGE);
                                if (reply == JOptionPane.YES_OPTION) {
                                    if (timeF.getText().trim().toLowerCase().equals("deltid") && !GUI.db.tableExist(caseF.getText().trim())) {
                                        JOptionPane.showMessageDialog(null, "Du valde ett deltidsschema.\n"
                                                + "Om deltagaren är här på FM, EM och ibland Heltid så ska man skapa\n"
                                                + "ett individuellt schema istället ('Individuellt schema' - knappen)", "Meddelande", JOptionPane.DEFAULT_OPTION);
                                    }
                                    //uppdatera ärendenumret i databasen också.
                                    if (!prevCase.equals(caseF.getText().trim())) {
                                        if (db.tableExist(prevCase)) {
                                            db.updateCaseNbr(caseF.getText().trim(), prevCase);
                                        }
                                    }
                                    changeInfoTimer.start();
                                    changeButton.setEnabled(false);
                                    changeButton.setText("Vänta...");
                                    groupNow = obj1.toString();
                                    //grupperInnan = excel.getGroups();
                                    excel.changeParticipantInfo(name, nameF.getText().trim(), startF.getText().trim(),
                                            endF.getText().trim(), timeF.getText().trim(), caseF.getText().trim(),
                                            languageF.getText().trim(), groupF.getText().trim(), groupNow);
                                    grupper = excel.getGroups();
                                    arrayEquality = Arrays.equals(grupperInnan, grupper);
                                    jcomboModel = new DefaultComboBoxModel(grupper);
                                    comboTypesList.setModel(jcomboModel);
                                    if (!arrayEquality) {
                                        grupperInnan = grupper.clone();//används i andra classer också
                                    }
                                    comboTypesList.setSelectedItem(groupF.getText().trim());
                                    putNamesOnList(groupF.getText().trim());

                                    windowClosed = true;
                                    infoFrame.setVisible(false);
                                }
                            } else {
                                if ((timeF.getText().trim().toLowerCase().equals("deltid") || timeF.getText().trim().toLowerCase().equals("fm/em")) && !GUI.db.tableExist(caseF.getText().trim())) {
                                    JOptionPane.showMessageDialog(null, "Du valde ett deltidsschema.\n"
                                            + "Om deltagaren är här på FM, EM och ibland Heltid så ska man skapa\n"
                                            + "ett individuellt schema istället ('Individuellt schema' - knappen)", "Meddelande", JOptionPane.DEFAULT_OPTION);
                                }
                                //uppdatera ärendenumret i databasen också.
                                if (!prevCase.equals(caseF.getText().trim())) {
                                    if (db.tableExist(prevCase)) {
                                        db.updateCaseNbr(caseF.getText().trim(), prevCase);
                                    }
                                }
                                changeInfoTimer.start();
                                changeButton.setEnabled(false);
                                changeButton.setText("Vänta...");
                                groupNow = obj1.toString();
                                //grupperInnan = excel.getGroups();
                                excel.changeParticipantInfo(name, nameF.getText().trim(), startF.getText().trim(),
                                        endF.getText().trim(), timeF.getText().trim(), caseF.getText().trim(),
                                        languageF.getText().trim(), groupF.getText().trim(), groupNow);
                                grupper = excel.getGroups();
                                arrayEquality = Arrays.equals(grupperInnan, grupper);
                                jcomboModel = new DefaultComboBoxModel(grupper);
                                comboTypesList.setModel(jcomboModel);
                                if (!arrayEquality) {
                                    //groupNow = groupF.getText();
                                    grupperInnan = grupper.clone();
                                }
                                comboTypesList.setSelectedItem(groupF.getText().trim());
                                putNamesOnList(groupF.getText().trim());

                                windowClosed = true;
                                infoFrame.setVisible(false);
                            }

                        } else {
                            JOptionPane.showMessageDialog(null, "Ange datumet i formatet ÅÅÅÅ-MM-DD", "Meddelande", JOptionPane.DEFAULT_OPTION);
                        }
                    } else {
                        JOptionPane.showMessageDialog(null, "Schemat måste skrivas FM, EM, FM/EM, Delid, Heltid eller Individuellt schema", "Meddelande", JOptionPane.DEFAULT_OPTION);
                    }
                } else {
                    JOptionPane.showMessageDialog(null, "Fyll i alla fält", "Meddelande", JOptionPane.DEFAULT_OPTION);
                }
            }
        }
        );

        btnCancel.addActionListener(
                new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent event
            ) {
                nameF.setText("");
                startF.setText("");
                endF.setText("");
                timeF.setText("");
                caseF.setText("");
                languageF.setText("");
                windowClosed = true;
                infoFrame.dispose();
            }
        }
        );
    }

    public static boolean isNum(String strNum) {
        boolean ret = true;
        try {
            Double.parseDouble(strNum);

        } catch (NumberFormatException e) {
            ret = false;
        }
        return ret;
    }

    public boolean changeGroup(String[] name, int day, String group3) throws IOException {
        String grupp2 = (String) JOptionPane.showInputDialog(null, "Välj den nya gruppen", "Flytta " + name[0],
                JOptionPane.PLAIN_MESSAGE,
                null,
                ExcelHandler.groups,
                ExcelHandler.groups[0]);

        boolean multiple = excel.checkForDublicates(name[0].trim(), grupp2);
        groupTrimed = grupp2;
        if (grupp2 != null && grupp2.equals(group3) == false) {
            if (multiple) {
                JOptionPane.showMessageDialog(null, "Namnet, " + name[0] + ", finns redan i gruppen " + grupp2 + "\n", "Meddelande", JOptionPane.DEFAULT_OPTION);
                return false;

            } else {
                String tid;
                if (!db.tableExist(model.getCheckedCaseNumber()[0])) {

                    if (!grupp2.toLowerCase().trim().contains("avbrott") && !grupp2.toLowerCase().trim().contains("avslut")
                            && !grupp2.toLowerCase().trim().contains("avslutat") && !grupp2.toLowerCase().trim().contains("avbrutet")
                            && !grupp2.toLowerCase().trim().contains("avbruten") && !grupp2.toLowerCase().trim().contains("avslutad")
                            && !grupp2.toLowerCase().trim().contains("avbr") && !grupp2.toLowerCase().trim().contains("avsl")
                            && !grupp2.toLowerCase().trim().contains("cancel") && !grupp2.toLowerCase().trim().contains("end")) {
                        tid = (String) JOptionPane.showInputDialog(null, "Välj schema för " + name[0] + "\n", "Schema",
                                JOptionPane.PLAIN_MESSAGE,
                                null,
                                registered,
                                registered[0]);
                        if (tid != null) {
                            if (tid.equals("Deltid")) {
                                JOptionPane.showMessageDialog(null, "Du valde ett deltidsschema.\n"
                                        + "Om deltagaren är här på FM, EM och ibland Heltid så ska man skapa\n"
                                        + "ett individuellt schema istället ('Individuellt schema' - knappen)", "Meddelande", JOptionPane.DEFAULT_OPTION);
                            }
                            excel.copyRow(name, grupp2, group3, tid);
                        } else {
                            groupTrimed = group3;
                        }
                    } else {
                        tid = excel.getActivity(name[0]);
                        if (tid.equals("tomt")) {
                            tid = " ";
                        }
                        excel.copyRow(name, grupp2, group3, tid);
                    }
                } //om det finns ett individuellt schema
                else {
                    tid = excel.getActivity(name[0]);
                    if (tid.equals("tomt")) {
                        tid = " ";
                    }
                    if (tid != null) {
                        excel.copyRow(name, grupp2, group3, tid);
                    } else {
                        groupTrimed = group3;
                    }
                }
            }
        } else {
            if (grupp2 != null) {
                JOptionPane.showMessageDialog(null, name[0] + " finns redan i gruppen " + grupp2 + "\n", "Meddelande", JOptionPane.DEFAULT_OPTION);
            } else {
                nullGroup = true;
            }
            return false;
        }
        return true;
    }

    public void checkUncheck() {
        if (hashMap.isEmpty()) {
            JOptionPane.showMessageDialog(null, "Exelfilen saknas! Börja med att importera en excelfil", "Meddelande", JOptionPane.DEFAULT_OPTION);
        } else {
            setButtonColor(checkAllButton);
            if (!checked) {
                model.checkAllNames();
            } else {
                model.uncheckAllNames();
            }
        }
    }

    public void setButtonColor(JButton buttonNow) {
        lastButton.setForeground(Color.BLACK);
        buttonNow.setForeground(new Color(88, 0, 4));
        lastButton = buttonNow;
    }

    public static boolean isValidFormat(String dateFromat, String dateToValidate) {
        if (dateToValidate == null) {
            return false;
        }
        SimpleDateFormat sdf = new SimpleDateFormat(dateFromat);
        sdf.setLenient(false);

        String[] Arr = dateToValidate.split("-");

        //året ska inte ha mer än fyra siffror, månad och dag max. 2 siffror
        try {
            if (Arr[0].length() > 4 || Arr[1].length() > 2 || Arr[2].length() > 2) {
                throw new ParseException("Error in date: " + dateToValidate, 0);
            }
            //om inte giltigt datum kastas det en exception
            //Date date = sdf.parse(dateToValidate);

        } catch (ParseException e) {
            return false;
        }
        return true;
    }

    private String fixCaseNumber(String fixedCaseNbr) {
        int diff = 0;
        String caseNbr = fixedCaseNbr;
        caseNbr = caseNbr.replace(".", "").replaceAll("[^\\x00-\\x7F]", "");
        caseNbr = caseNbr.replace("E7", "");
        caseNbr = caseNbr.trim();

        if (caseNbr.isEmpty()) {
            caseNbr = "0";
        }

        if (caseNbr.length() < 8 && caseNbr.length() > 0) {
            diff = 8 - caseNbr.length();
            for (int h = 0; h < diff; h++) {
                caseNbr = caseNbr + "0";
            }
        }
        return caseNbr;
    }

    //Lägg till namnen i listan beroende på klass (grupp)
    private void putNamesOnList(String grupp) {
        String namn = "";
        String caseNbr = "";
        String endD;
        String startD;
        grupper3 = null;
        nbrOfP = 0;
        wrongDateList = new ArrayList<>();
        String daysLeftString;
        SimpleDateFormat df;
        DateTimeFormatter formatter;
        LocalDate todayDate;
        LocalDate endDate;
        LocalDate startDate;
        //int diff = 0;
        List<String> values;
        long daysBetween = 0;
        boolean felDatum;

        endTimePartic = new HashMap<>();
        //om man byter grupp så ska också "avmarkera" bli "markera"
        if (GUI.checked == true) {
            GUI.checkAllButton.setText("Markera alla");
            GUI.checked = false;
        }

        try {
            hashMap = excel.findNamesByGroup(grupp);
            Set set = hashMap.entrySet();

            //om namnen redan finns på listan, ta bort de först.
            if (model.getRowCount() > 0) {
                for (int i = model.getRowCount() - 1; i > -1; i--) {
                    model.removeRow(i);
                }
            }

            for (Map.Entry<String, List<String>> entry : hashMap.entrySet()) {
                felDatum = false;
                namn = entry.getKey();
                values = entry.getValue();
                caseNbr = values.get(0);
                caseNbr = fixCaseNumber(caseNbr);
                endD = values.get(1);
                startD = excel.getStartDateByName(namn);

                df = new SimpleDateFormat("yyyy-MM-dd");
                formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");
                cal = Calendar.getInstance();
                try {
                    todayDate = LocalDate.parse(df.format(cal.getTime()), formatter);
                    endDate = LocalDate.parse(endD.trim(), formatter);
                    startDate = LocalDate.parse(startD.trim(), formatter);

                    grupper3 = new ArrayList<>();
                    //räkna inte med lör och sön om deltagaren slutar då. Ange fredag som slutdatum i sådana fall.
                    switch (endDate.getDayOfWeek().name()) {
                        case "SUNDAY":
                            endDate = endDate.minusDays(2);
                            break;
                        case "SATURDAY":
                            endDate = endDate.minusDays(1);
                            break;
                    }
                    //kolla om startdatumet är i framtiden och markera deltagaren med blå/lilla färg
                    daysBetween = ChronoUnit.DAYS.between(todayDate, startDate);
                    if (daysBetween > 0) {
                        grupper3.add("future");
                        endTimePartic.put(namn, grupper3);
                        //räkna dagarna tills deltagaren ska starta
                        daysBetween = ChronoUnit.DAYS.between(todayDate, startDate);
                    } else {
                        daysBetween = ChronoUnit.DAYS.between(todayDate, endDate);
                        //Kolla om deltagaren slutar snart / har slutat
                        //markera deltagare med orange färg tre veckor innan han/hon ska sluta
                        if (daysBetween >= 0 && daysBetween < 22) {
                            if (daysBetween == 0) {
                                grupper3.add("today");
                            } else {
                                grupper3.add("soon");
                            }
                            endTimePartic.put(namn, grupper3);
                        } else if (daysBetween < 0) {
                            grupper3.add("end");
                            endTimePartic.put(namn, grupper3);
                        }
                    }
                } catch (DateTimeParseException ex) {
                    felDatum = true;
                    //daysBetween = -5000;
                    if (namn.equals("Kolumnfel")) {
                        //JOptionPane.showMessageDialog(null, "Ta bort kolumner i excelen och försök sedan igen.\n"
                        // + "Det ska vara totalt 38 kolumner (Namn, Grupp, ..., 31)", "Meddelande", JOptionPane.DEFAULT_OPTION);
                    } else {
                        wrongDateList.add(namn);
                        /*
                         JOptionPane.showMessageDialog(null, "Kontrollera datumformatet för " + namn + ".\n"
                         + "1. Se till att datumen är skrivna i formatet ÅÅÅÅ-MM-DD.\n"
                         + "2. Se till att datumet inte innehåller 'space'\n\n"
                         + "Använd knappen 'Deltagarinfo / Ändra info' för att ändra datumformat / ta bort space\n\n", "Meddelande", JOptionPane.DEFAULT_OPTION);
                         */
                    }
                }

                nbrOfP++;

                daysLeftString = Long.toString(daysBetween);
                if (felDatum == true) {
                    daysLeftString = "Fel datum!";
                }
                //lägg till ett, i, om deltagaren har ett individuellt schema
                if (db.tableExist(caseNbr)) {
                    if (daysLeftString.length() == 1) {
                        daysLeftString = "     " + daysLeftString + "     i";
                    } else if (daysLeftString.length() == 2) {
                        daysLeftString = "    " + daysLeftString + "    i";
                    } else if (daysLeftString.length() == 3) {
                        daysLeftString = "   " + daysLeftString + "   i";
                    } else if (daysLeftString.length() == 4) {
                        daysLeftString = "  " + daysLeftString + "  i";
                    } else {
                        daysLeftString = " " + daysLeftString + "  i";
                    }
                }

                Object[] row = {Boolean.FALSE, namn, caseNbr, " ", " ", daysLeftString};
                model.addRow(row);
                nbrOfPa = Integer.toString(nbrOfP);
                numberOfPa.setText(nbrOfPa + " delt.");
            }
            if (wrongDateList.size() > 0 && !grupp.equals(latestGroupp)) {
                JOptionPane.showMessageDialog(null, "Kontrollera datumen för deltagare.\n"
                        + "1. Se till att datum finns.\n"
                        + "2. Se till att datumen är i formatet ÅÅÅÅ-MM-DD.\n\n"
                        + "Använd knappen 'Deltagarinfo / Ändra info' för att ändra datum.\n\n", "Meddelande", JOptionPane.DEFAULT_OPTION);
            }

        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "Ett fel har inträffat (fel: putNamesOnList)", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
        if ((sparat == true || save.getText().equals(" ") || !save.getText().equals("Ej sparat")) && !grupp.equals(latestGroupp)) {
            save.setText("Ej sparat");
            save.setForeground(new Color(153, 0, 0));
            sparat = false;
        }
        RowSorter<TableModel> sorter = new TableRowSorter<>(model);

        table.setRowSorter(sorter);
        table.getRowSorter().toggleSortOrder(1);
        latestGroupp = grupp;
        table.setDefaultRenderer(Object.class, new DefaultTableCellRenderer() {
            int count2 = 0;
            Component c;

            @Override
            public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus, int row, int column) {
                c = super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
                for (Map.Entry<String, List<String>> entry : endTimePartic.entrySet()) {
                    if (value.equals(entry.getKey())) {

                        switch (entry.getValue().get(0)) {
                            case "soon":
                                setForeground(new Color(0xB45F04));
                                setFont(getFont().deriveFont(Font.BOLD));
                                break;
                            case "today":
                                setForeground(new Color(0xFF0080));
                                setFont(getFont().deriveFont(Font.BOLD));
                                break;
                            case "end":
                                setForeground(new Color(0xDF013A));
                                setFont(getFont().deriveFont(Font.BOLD));
                                break;
                            case "future":
                                setForeground(new Color(0x6565bf));
                                setFont(getFont().deriveFont(Font.BOLD));
                                break;
                        }
                        break;

                    } else {
                        //c.setBackground(Color.WHITE);
                        setForeground(Color.black);
                    }
                }
                return c;
            }
        });
    }

    private void clearEverything(String grupp) {
        try {
            hashMap = excel.findNamesByGroup(grupp);

            //om namnen redan finns på listan, ta bort de först.
            if (model.getRowCount() > 0) {
                for (int i = model.getRowCount() - 1; i > -1; i--) {
                    model.removeRow(i);
                }
            }
            putNamesOnList(groupNow);
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "Ett fel har inträffat (fel: clearEverything)", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
        //model.checkAllNames();
    }

    static void destroyGUI() {
        frame.dispose();
    }

    static void createAndShowUI() {
        frame = new JFrame("© Attendance Tool (2.2.1.1)  2020        ( burim.sadriu82@gmail.com )  ");
        frame.getContentPane().add(new GUI());
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(825, 705);
        frame.setResizable(false);

        //frame.pack();
        frame.setLocationRelativeTo(null);
//        frame.setIconImage((imageIcon.getImage()));
        frame.setVisible(true);

        if (groupNow.equals("Excelfil saknas")) {
            JOptionPane.showMessageDialog(null, "Excelfilen kunde inte hittas\n\n"
                    + "1. Säkerställ att mappen ’Attendance Tool’ finns på ’Skrivbordet’.\n"
                    + "2. Börja med att importera excelfilen med knappen ’Importera excelfil’.\n"
                    + "3. Se till att alla kolumner har ett namn innan du importerar filen\n"
                    + "    och att det finns totalt, 38 kolumner.\n\n\n", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
    }
}

@SuppressWarnings("serial")
class CheckBoxDefaultTableModel extends DefaultTableModel {

    private final List<String> removedItemsList = new ArrayList<>();

    public CheckBoxDefaultTableModel(Object[] columnNames, int rowCount) {
        super(columnNames, rowCount);
    }

    //kolla om namnet finns i listan av de checkade namnen. Om ja return true, annars false.
    public boolean isChecked(String name, String[] checkedNames) {
        boolean isCheked = false;
        for (int i = 0; i < checkedNames.length; i++) {
            if ((checkedNames[i].equals(name))) {
                isCheked = true;
                break;
            }
        }
        return isCheked;
    }

    @Override
    public Class<?> getColumnClass(int columnNumber) {
        if (columnNumber == 0) {
            return Boolean.class;
        }
        return super.getColumnClass(columnNumber);
    }

    public void removeCheckedItems() {
        int rowCount = getRowCount();
        for (int row = rowCount - 1; row >= 0; row--) {
            if ((Boolean) getValueAt(row, 0)) {
                removedItemsList.add(getValueAt(row, 1).toString());
                removeRow(row);
            }
        }
    }

    public void checkAllNames() {
        int rowCount = getRowCount();
        for (int row = rowCount - 1; row >= 0; row--) {
            if (!(Boolean) getValueAt(row, 0)) {
                setValueAt(true, row, 0);
            }
        }
        GUI.checkAllButton.setText("Avmarkera alla");
        GUI.checked = true;
    }

    public void uncheckAllNames() {
        int rowCount = getRowCount();
        for (int row = rowCount - 1; row >= 0; row--) {
            if ((Boolean) getValueAt(row, 0)) {
                setValueAt(false, row, 0);
            }
        }
        GUI.checkAllButton.setText("Markera alla");
        GUI.checked = false;
    }

    //Hämta de markerade namnen för att skapa pdf.
    //Metoden är en parameter till excel.getValuesFromExcel()
    public String[] getCheckedNames() {
        String checkedName = "";
        String[] checkedNames;
        int j = 0;
        //räkna antalet markerade namn, för att initialisera arrayen.
        int rowCount = getRowCount();
        for (int row = rowCount - 1; row >= 0; row--) {
            if ((Boolean) getValueAt(row, 0)) {
                j++;
            }
        }
        checkedNames = new String[j];

        //hämta de markerade namnen
        j = 0;
        for (int row = rowCount - 1; row >= 0; row--) {
            if ((Boolean) getValueAt(row, 0)) {
                checkedName = getValueAt(row, 1).toString();
                checkedNames[j] = checkedName;
                j++;
            }
        }
        return checkedNames;
    }

    public String[] getCheckedCaseNumber() {
        String checkedCaseNbr = "";
        String[] checkedCase;
        int j = 0;
        //räkna antalet markerade namn, för att initialisera arrayen.
        int rowCount = getRowCount();
        for (int row = rowCount - 1; row >= 0; row--) {
            if ((Boolean) getValueAt(row, 0)) {
                j++;
            }
        }
        checkedCase = new String[j];

        //hämta de markerade namnen
        j = 0;
        for (int row = rowCount - 1; row >= 0; row--) {
            if ((Boolean) getValueAt(row, 0)) {
                checkedCaseNbr = getValueAt(row, 2).toString();
                checkedCase[j] = checkedCaseNbr;
                j++;
            }
        }
        return checkedCase;
    }

    //Ta info från varje checkbox och spara i en HashMap. 
    //HashMappen skickas till excel.addValuesToExcel() för att lägga in värden i excel filen.
    public Map<String, List<String>> getInfoFromCheckBoxes(String grupp, String day) throws IOException {
        String timeOf = "";
        String note = "error";
        String name = "";
        String caseNbr = "";
        String schedule = "";
        String[] checkedNames = getCheckedNames();
        GUI.canceled = false;
        GUI.tom = " ";

        Map<String, List<String>> hmap = new HashMap<>();

        int rowCount = getRowCount();
        List<String> values;
        String reducedHours = " ";
        //gå igenom alla checkbox
        for (int row = rowCount - 1; row >= 0; row--) {
            values = new ArrayList<>();
            name = getValueAt(row, 1).toString();
            caseNbr = getValueAt(row, 2).toString();
            //om namnet är markerad
            try {
                //om deltagaren har ett individuellt schema, utgå från det.
                if (GUI.db.tableExist(caseNbr)) {
                    schedule = GUI.db.getDaySchedule(day, caseNbr).trim().toLowerCase();

                    //om deltagarens aktivitet inte varierar
                } else {
                    schedule = GUI.excel.getActivity(name).trim().toLowerCase();
                    schedule = schedule.replaceAll("\\s+", "");// ta bort space
                    if (schedule.equals("tomt")) {
                        JOptionPane.showMessageDialog(null, name + " har ett tomt schema.\n"
                                + "Ändra schemat till FM, EM, Heltid eller Deltid genom att markera\n"
                                + "namnet och sedan välja 'Deltagarinfo / Ändra info'- knappen.\n\n"
                                + "Man kan också skapa ett individuellt schema\n"
                                + "med hjälp av 'Individuellt schema'- knappen.\n\n", "Meddelande", JOptionPane.DEFAULT_OPTION);
                        GUI.tom = "tom";
                        GUI.canceled = true;
                        if (!GUI.save.getText().equals("Ej sparat")) {
                            GUI.save.setText("Ej sparat");
                            GUI.save.setForeground(new Color(153, 0, 0));
                        }
                        break;
                    } else if (!schedule.equals("heltid") && !schedule.equals("em") && !schedule.equals("fm")
                            && !schedule.equals("em/fm") && !schedule.equals("fm/em") && !schedule.equals("deltid")) {
                        JOptionPane.showMessageDialog(null, name + " har ett inkorrekt schema.\n\n"
                                + "Ändra schemat till FM, EM, Heltid eller Deltid genom att markera\n"
                                + "namnet och sedan välja 'Deltagarinfo / Ändra info'- knappen.\n\n"
                                + "Man kan också skapa ett individuellt schema\n"
                                + "med hjälp av 'Individuellt schema'- knappen.\n\n", "Meddelande", JOptionPane.DEFAULT_OPTION);
                        GUI.canceled = true;
                        break;
                    }
                }

            } catch (SQLException ex) {
                JOptionPane.showMessageDialog(null, "Ett fel har inträffat (fel: all info method)", "Meddelande", JOptionPane.DEFAULT_OPTION);
            }
            reducedHours = getValueAt(row, 3).toString().trim().toLowerCase();

            if (isChecked(name, checkedNames)) {

                if ((reducedHours.equals("4.15") || reducedHours.equals("4.30")
                        || reducedHours.equals("4.45") || reducedHours.equals("5") || reducedHours.equals("5.15")
                        || reducedHours.equals("5.30") || reducedHours.equals("5.45") || reducedHours.equals("6")
                        || reducedHours.equals("6.15") || reducedHours.equals("6.30") || reducedHours.equals("6.45")
                        || reducedHours.equals("7") || reducedHours.equals("7.15") || reducedHours.equals("7.30")
                        || reducedHours.equals("7.45") || reducedHours.equals("8"))
                        && (schedule.equals("sfi-em") || schedule.equals("fm-sfi")
                        || schedule.equals("ledig-em") || schedule.equals("fm-ledig")
                        || schedule.equals("p-em") || schedule.equals("fm-p")
                        || schedule.equals("so-ledig") || schedule.equals("ledig-so")
                        || schedule.equals("so-sfi") || schedule.equals("sfi-so")
                        || schedule.equals("so-p") || schedule.equals("p-so")
                        || schedule.equals("so") || schedule.equals("sjuk")
                        || schedule.equals("fl") || schedule.equals("fm")
                        || schedule.equals("em") || schedule.equals("blandat")
                        || schedule.equals("jobb") || schedule.equals("em/fm")
                        || schedule.equals("fm/em") || schedule.equals("deltid")
                        || schedule.equals("fm-stom") || schedule.equals("stom-em")
                        || schedule.equals("fm-uf") || schedule.equals("uf-em")
                        || schedule.equals("fl-sfi") || schedule.equals("sfi-fl")
                        || schedule.equals("vab") || schedule.equals("akt-em_50%")
                        || schedule.equals("fm-akt_50%") || schedule.equals("akt-akt_50%")
                        || schedule.equals("komv-em_50%") || schedule.equals("fm-komv_50%")
                        || schedule.equals("komv-komv_50%"))) {

                    JOptionPane.showConfirmDialog(null, "Du har valt mer än 4 timmars frånvaro för " + name + ".\n"
                            + "Deltid motsvarar max. 4 timmar per dag", "Meddelande", JOptionPane.DEFAULT_OPTION, JOptionPane.PLAIN_MESSAGE);
                    GUI.canceled = true;
                    break;
                } else if (schedule.equals("p") || schedule.equals("sfi") || schedule.equals("so")
                        || schedule.equals("so-so") || schedule.equals("jobb") || schedule.equals("jobb-jobb")
                        || schedule.equals("so-p") || schedule.equals("p-so") || schedule.equals("sfi-p")
                        || schedule.equals("p-sfi") || schedule.equals("so-sfi") || schedule.equals("sfi-so")
                        || schedule.equals("so-ledig") || schedule.equals("ledig-so") || schedule.equals("sfi-ledig")
                        || schedule.equals("ledig-sfi") || schedule.equals("ledig") || schedule.equals("sjuk")
                        || schedule.equals("sjuk-sjuk") || schedule.equals("fl") || schedule.equals("fl-fl")
                        || schedule.equals("stom") || schedule.equals("uf") || schedule.equals("fl-sfi") || schedule.equals("sfi-fl")
                        || schedule.equals("vab") || schedule.equals("vab-vab") || schedule.equals("akt")
                        || schedule.equals("akt-akt_100%") || schedule.equals("akt-akt_50%") || schedule.equals("komv-komv_50%")
                        || schedule.equals("komv-komv_100%") || schedule.equals("komv")) {

                    JOptionPane.showConfirmDialog(null, "Enligt det individuella schemat ska " + name + "\n"
                            + "inte vara närvarande den här dagen!", "Meddelande", JOptionPane.DEFAULT_OPTION, JOptionPane.PLAIN_MESSAGE);
                    GUI.canceled = true;
                    break;
                } //om boxen är checkad och det inte finns anteckningar i
                //frånvarotimmarna så sätt x. x betyder full närvaro.
                else if (reducedHours.equals(" ") || reducedHours.isEmpty() || reducedHours.equals("")) {
                    timeOf = "x";//full närvaro
                } //annars ta anteckningarna i frånvarotimmarna
                else {
                    if (reducedHours.equals("0.15") || reducedHours.equals("0.30")
                            || reducedHours.equals("0.45") || reducedHours.equals("1") || reducedHours.equals("1.15")
                            || reducedHours.equals("1.30") || reducedHours.equals("1.45") || reducedHours.equals("2")
                            || reducedHours.equals("2.15") || reducedHours.equals("2.30") || reducedHours.equals("2.45")
                            || reducedHours.equals("3") || reducedHours.equals("3.15") || reducedHours.equals("3.30")
                            || reducedHours.equals("3.45") || reducedHours.equals("4") || reducedHours.equals("4.15")
                            || reducedHours.equals("4.30") || reducedHours.equals("4.45") || reducedHours.equals("5")
                            || reducedHours.equals("5.15") || reducedHours.equals("5.30") || reducedHours.equals("5.45")
                            || reducedHours.equals("6") || reducedHours.equals("6.15") || reducedHours.equals("6.30")
                            || reducedHours.equals("6.45") || reducedHours.equals("7") || reducedHours.equals("7.15")
                            || reducedHours.equals("7.30") || reducedHours.equals("7.45") || reducedHours.equals("8")) {

                        timeOf = "x-" + reducedHours;//ej full närvaro

                    } else {
                        JOptionPane.showConfirmDialog(null, "Du har markerat " + name + ", men inte valt en siffra som avvikelse\n\n"
                                + "( Markering = närvaro )\n", "Meddelande", JOptionPane.DEFAULT_OPTION, JOptionPane.PLAIN_MESSAGE);
                        GUI.canceled = true;
                        break;
                    }
                }
            } //om namnet inte är checkad så betyder det att deltagaren inte
            //är närvarande.
            else {
                if (reducedHours.equals("0.15") || reducedHours.equals("0.30")
                        || reducedHours.equals("0.45") || reducedHours.equals("1") || reducedHours.equals("1.15")
                        || reducedHours.equals("1.30") || reducedHours.equals("1.45") || reducedHours.equals("2")
                        || reducedHours.equals("2.15") || reducedHours.equals("2.30") || reducedHours.equals("2.45")
                        || reducedHours.equals("3") || reducedHours.equals("3.15") || reducedHours.equals("3.30")
                        || reducedHours.equals("3.45") || reducedHours.equals("4") || reducedHours.equals("4.15")
                        || reducedHours.equals("4.30") || reducedHours.equals("4.45") || reducedHours.equals("5")
                        || reducedHours.equals("5.15") || reducedHours.equals("5.30") || reducedHours.equals("5.45")
                        || reducedHours.equals("6") || reducedHours.equals("6.15") || reducedHours.equals("6.30")
                        || reducedHours.equals("6.45") || reducedHours.equals("7") || reducedHours.equals("7.15")
                        || reducedHours.equals("7.30") || reducedHours.equals("7.45") || reducedHours.equals("8")
                        || reducedHours.equals("s") || reducedHours.equals("a") || reducedHours.equals("v")
                        || reducedHours.equals("ö") || reducedHours.equals("p") || reducedHours.equals("af")) {

                    timeOf = reducedHours;

                } else if (schedule.equals("p") || schedule.equals("sfi")) {

                    timeOf = schedule;

                } else if (schedule.equals("p-sfi")
                        || schedule.equals("sfi-p")) {

                    timeOf = "p";

                } else if (schedule.equals("ledig") || schedule.equals("sfi-ledig")
                        || schedule.equals("ledig-sfi") || schedule.equals("stom")
                        || schedule.equals("uf") || schedule.equals("akt")
                        || schedule.equals("komv")) {

                    timeOf = " ";

                } else if (schedule.equals("sfi-em") || schedule.equals("fm-sfi")
                        || schedule.equals("ledig-em") || schedule.equals("fm-ledig")
                        || schedule.equals("p-em") || schedule.equals("fm-p")
                        || schedule.equals("so-ledig") || schedule.equals("ledig-so")
                        || schedule.equals("so-sfi") || schedule.equals("sfi-so")
                        || schedule.equals("so-p") || schedule.equals("p-so")
                        || schedule.equals("so") || schedule.equals("sjuk")
                        || schedule.equals("fl") || schedule.equals("fm")
                        || schedule.equals("em") || schedule.equals("blandat")
                        || schedule.equals("jobb") || schedule.equals("em/fm")
                        || schedule.equals("fm/em") || schedule.equals("deltid")
                        || schedule.equals("fm-stom") || schedule.equals("stom-em")
                        || schedule.equals("fm-uf") || schedule.equals("uf-em")
                        || schedule.equals("fl-sfi") || schedule.equals("sfi-fl")
                        || schedule.equals("vab") || schedule.equals("fm-akt_50%")
                        || schedule.equals("akt-em_50%") || schedule.equals("akt-akt_50%")
                        || schedule.equals("komv-em_50%") || schedule.equals("fm-komv_50%")
                        || schedule.equals("komv-komv_50%")) {

                    timeOf = "d";//halv frånvaro för deltid

                } else if (schedule.equals("heltid") || schedule.equals("so-em")
                        || schedule.equals("fm-so") || schedule.equals("so-so")
                        || schedule.equals("sjuk-sjuk") || schedule.equals("fl-fl")
                        || schedule.equals("jobb-jobb") || schedule.equals("fm-jobb")
                        || schedule.equals("jobb-em") || schedule.equals("vab-vab")
                        || schedule.equals("fm-akt_100%") || schedule.equals("akt-em_100%")
                        || schedule.equals("akt-akt_100%") || schedule.equals("fm-komv_100%")
                        || schedule.equals("komv-em_100%") || schedule.equals("komv-komv_100%")) {

                    timeOf = "h";//full frånvaro för heltid
                }
            }
            note = getValueAt(row, 4).toString();
            values.add(timeOf);
            values.add(note);
            hmap.put(name, values);
        }
        return hmap;
    }
}
