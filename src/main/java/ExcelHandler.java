import com.itextpdf.text.DocumentException;
import java.awt.Desktop;
import java.awt.Frame;
import java.io.FileInputStream;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.Statement;
import java.io.*;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.GregorianCalendar;
import java.util.HashMap;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.*;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JDialog;
import javax.swing.JOptionPane;
import javax.swing.WindowConstants;
import org.apache.poi.hssf.OldExcelFormatException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 *
 * @author Burim Sadriu Copyright© 2020, All Rights Reserved
 */
public class ExcelHandler {

    ResultSet resultSet = null;
    ResultSet resultSet2 = null;
    Statement statement = null;
    Statement statement2 = null;
    FileInputStream fsIP;
    HSSFWorkbook wb;
    String[] tempDeltagare;
    String grupp = "";
    int row;
    int index;
    Connection con;
    static String filePath = System.getProperty("user.home") + "\\Desktop\\Attendance Tool\\Dokument\\deltagare.xls";
    int[] rows = new int[50];
    Cell cell = null;
    HSSFSheet worksheet;
    int dayNumber;
    Calendar now = GregorianCalendar.getInstance();
    PresenceReportPDF presRepPDF = new PresenceReportPDF();
    HashMap<String, String> hmap;
    boolean wordOpen = false;
    public static String[] groups;
    PeriodicReportPDF abPdf;
    AttendanceWordList tt;
    OpenFoldersAndDoc openFolder = new OpenFoldersAndDoc();

    public int getRowIndexByName(String name) {
        int rowIndex = 0;
        worksheet = wb.getSheetAt(0); //Access the worksheet, so that we can update / modify it.
        cell = null;
        for (int i = 1; i <= worksheet.getLastRowNum(); i++) {
            try {
                cell = worksheet.getRow(i).getCell(0);
            } catch (NullPointerException NPE) {
                continue;
            }
            if (cell.toString().equals(name)) {
                rowIndex = cell.getRowIndex();
                break;
            }
        }
        return rowIndex;
    }

    public int getRowIndexByNameAndGroup(String name, String group) {
        int rowIndex = 0;
        worksheet = wb.getSheetAt(0); //Access the worksheet, so that we can update / modify it.
        cell = null;
        Cell grupp = null;
        for (int i = 1; i <= worksheet.getLastRowNum(); i++) {
            try {
                grupp = worksheet.getRow(i).getCell(1);
                cell = worksheet.getRow(i).getCell(0);
            } catch (NullPointerException e) {
                continue;
            }

            if (cell.toString().equals(name) && grupp.toString().equals(group)) {
                rowIndex = cell.getRowIndex();
                break;
            }
        }
        return rowIndex;
    }

    public void createPeriodicReports() {
        abPdf = new PeriodicReportPDF();
        String previousDate = " ";
        String dateFromExcel;
        File file = new File(filePath);
        boolean isClosed = true;
        try {
            fsIP = new FileInputStream(file); //Read the spreadsheet that needs to be updated
        } catch (FileNotFoundException ex) {
            JOptionPane.showMessageDialog(null, "Kontrollera Excelfilen (deltagare.xls) i mappen 'Attendance Tool' → Dokument'", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
        try {
            wb = new HSSFWorkbook(fsIP); //Access the workbook
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "Excelfilen kunde inte hittas.\n"
                    + "Börja med att importera en Excelfil", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
        worksheet = wb.getSheetAt(0); //Access the worksheet, so that we can update / modify it.
        cell = null;
        for (int i = 1; i <= worksheet.getLastRowNum(); i++) {
            try {
                cell = worksheet.getRow(i).getCell(2);// startdatum

            } catch (NullPointerException e) {
                continue;
            }
            if (cell == null) {
                continue;
            }
            dateFromExcel = cell.toString().trim();
            if (dateFromExcel.isEmpty()) {
                continue;
            }
            if (!dateFromExcel.equals(previousDate)) {

                previousDate = dateFromExcel;
                try {
                    isClosed = getAbsenceFromExcel(getNameByStartDate(previousDate), previousDate);
                    if (isClosed == false) {
                        break;
                    }
                } catch (DocumentException ex) {
                    Logger.getLogger(ExcelHandler.class.getName()).log(Level.SEVERE, null, ex);
                }
            }
        }
        try {
            if (isClosed == true) {
                openFolder.openPeriodReportsFolder();
            }
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(null, "Kontrollera mappen 'Attendance Tool' på skrivbordet.\n", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
        try {
            fsIP.close();
        } catch (IOException ex) {
            Logger.getLogger(ExcelHandler.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    public int getRowIndexByNameAndStartDate(String name, String startDate) {
        int rowIndex = 0;
        //worksheet = wb.getSheetAt(0); //Access the worksheet, so that we can update / modify it.
        cell = null;
        Cell startD = null;
        for (int i = 1; i <= worksheet.getLastRowNum(); i++) {
            try {
                startD = worksheet.getRow(i).getCell(2);
                cell = worksheet.getRow(i).getCell(0);
            } catch (NullPointerException e) {
                continue;
            }

            if (cell.toString().trim().equals(name) && startD.toString().trim().equals(startDate)) {
                rowIndex = cell.getRowIndex();
                break;
            }
        }
        return rowIndex;
    }

    public int getRowIndexByCaseNumber(String caseNbr) {
        int rowIndex = 0;
        int diff = 0;
        String caseFromExcel;
        worksheet = wb.getSheetAt(0); //Access the worksheet, so that we can update / modify it.
        cell = null;
        for (int i = 1; i <= worksheet.getLastRowNum(); i++) {
            try {
                cell = worksheet.getRow(i).getCell(6);// ärendenummer
            } catch (NullPointerException e) {
                continue;
            }
            caseFromExcel = cell.toString();
            caseFromExcel = fixCaseNumber(caseFromExcel);
            if (caseFromExcel.equals(caseNbr)) {
                rowIndex = cell.getRowIndex();
                break;
            }
        }
        return rowIndex;
    }

    public String[] getNameByStartDate(String startDate) {
        int rowIndex = 0;
        int diff = 0;
        String dateFromExcel;
        int arraySize = 0;
        int count = 0;
        String[] names = null;
        //File file = new File(filePath);

        //worksheet = wb.getSheetAt(0); //Access the worksheet, so that we can update / modify it.
        cell = null;
        Cell nameCell = null;
        for (int i = 1; i <= worksheet.getLastRowNum(); i++) {
            try {
                cell = worksheet.getRow(i).getCell(2);// startdatum
                nameCell = worksheet.getRow(i).getCell(0);

            } catch (NullPointerException e) {
                continue;
            }
            if (cell == null || nameCell == null) {
                continue;
            }
            dateFromExcel = cell.toString();
            if (dateFromExcel.trim().isEmpty() || nameCell.toString().trim().isEmpty()) {
                continue;
            }
            if (dateFromExcel.trim().equals(startDate.trim())) {
                arraySize++;
            }
        }
        if (arraySize > 0) {
            names = new String[arraySize];
            for (int i = 1; i <= worksheet.getLastRowNum(); i++) {
                try {
                    cell = worksheet.getRow(i).getCell(2);// startdatum
                    nameCell = worksheet.getRow(i).getCell(0);
                } catch (NullPointerException e) {
                    continue;
                }
                if (cell == null || nameCell == null) {
                    continue;
                }
                dateFromExcel = cell.toString();

                if (dateFromExcel.trim().isEmpty() || nameCell.toString().trim().isEmpty()) {
                    continue;
                }
                if (dateFromExcel.trim().equals(startDate.trim())) {
                    names[count] = nameCell.toString().trim();
                    count++;
                }
            }
        } else {

            return names;
        }
        return names;
    }

    public int getOldPartIndex(String name, String oldGroup) {
        int rowIndex = 0;
        worksheet = wb.getSheetAt(0); //Access the worksheet, so that we can update / modify it.
        Cell gruppCell = null;
        for (int i = 1; i <= worksheet.getLastRowNum(); i++) {
            try {
                cell = worksheet.getRow(i).getCell(0);
                gruppCell = worksheet.getRow(i).getCell(1);
            } catch (NullPointerException e) {
                continue;
            }
            if (cell.toString().equals(name) && gruppCell.toString().equals(oldGroup)) {
                rowIndex = cell.getRowIndex();
                break;
            }
        }
        return rowIndex;
    }

    private String fixCaseNumber(String fixedCaseNbr) {
        int diff = 0;
        String caseNbr = fixedCaseNbr;
        //replace all non ASCII letters
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

    public static boolean isRowEmpty(Row row) {
        if (row != null) {
            for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
                Cell cell = row.getCell(c);
                if (cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK) {
                    return false;
                }
            }
        }
        return true;
    }

    public boolean changeAbsence(String[] name, String day, String absence) {
        boolean isOpen = false;
        File file = new File(filePath);
        try {
            fsIP = new FileInputStream(file); //Read the spreadsheet that needs to be updated
        } catch (FileNotFoundException ex) {
            JOptionPane.showMessageDialog(null, "Kontrollera Excelfilen (deltagare.xls) i mappen 'Attendance Tool' → Dokument'", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
        try {
            wb = new HSSFWorkbook(fsIP); //Access the workbook
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "Excelfilen kunde inte hittas.\n"
                    + "Börja med att importera en Excelfil", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
        worksheet = wb.getSheetAt(0);
        Cell nameCell = null;
        int dayInt = Integer.parseInt(day);
        for (String name1 : name) {
            row = getRowIndexByName(name1);
            wb.setMissingCellPolicy(Row.CREATE_NULL_AS_BLANK);
            nameCell = worksheet.getRow(row).getCell(dayInt + 6);
            nameCell.setCellValue(absence);
        }
        try (FileOutputStream output_file = new FileOutputStream(new File(filePath))) //Open FileOutputStream to write updates
        {
            wb.write(output_file); //write changes
            fsIP.close(); //Close the InputStream
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "Excelfilen är öppen. Stäng den och försök sedan igen", "Meddelande", JOptionPane.DEFAULT_OPTION);
            isOpen = true;
            return isOpen;
        }
        return isOpen;
    }

    public String[] getGroups() {
        String s = "";
        String ss = "";
        String mOrl = "";
        File file = new File(filePath);
        ArrayList<String> grupper = new ArrayList<>();
        try {
            fsIP = new FileInputStream(file);
        } catch (FileNotFoundException ex) {

            grupper.add("Excelfil saknas");
            groups = new String[grupper.size()];
            for (int t = 0; t < grupper.size(); t++) {
                groups[t] = grupper.get(t);
            }
            return groups;
        }
        try {
            wb = new HSSFWorkbook(fsIP); //Access the workbook
        } catch (OldExcelFormatException c) {
            JOptionPane.showMessageDialog(null, "1. Se till att den originala excelfilen är sparad i filformatet\n"
                    + "     Excel-arbetsbok (*.xlsx).\n\n"
                    + "2. Gå till mappen 'Attendance Tool' → ’Dokument’ och ta bort filen 'deltagare.xls'\n\n"
                    + "3. Försök att importera den originala excelfilen igen.\n\n\n", "Meddelande", JOptionPane.DEFAULT_OPTION);

        } catch (IOException ex) {

            JOptionPane.showMessageDialog(null, "Kontrollera Excelfilen (deltagare.xls) i mappen 'Attendance Tool' → ’Dokument’", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }

        worksheet = wb.getSheetAt(0); //Access the worksheet, so that we can update / modify it.
        wb.setMissingCellPolicy(Row.CREATE_NULL_AS_BLANK);
        int noOfColumns = worksheet.getRow(0).getPhysicalNumberOfCells();
        Cell wrongColumn1 = worksheet.getRow(0).getCell(0);
        Cell wrongColumn2 = worksheet.getRow(0).getCell(1);
        if (noOfColumns > 38 || noOfColumns < 38 || wrongColumn1 == null && wrongColumn1.getCellType() == Cell.CELL_TYPE_BLANK || wrongColumn1.toString().isEmpty()
                || wrongColumn2 == null && wrongColumn2.getCellType() == Cell.CELL_TYPE_BLANK || wrongColumn2.toString().isEmpty()) {

            JOptionPane.showMessageDialog(null, "Det finns fel i kolumnerna i excelfilen.\n\n"
                    + "1. Kolumnerna ska vara: Namn, Grupp, Startdatum, Slutdatum,\n"
                    + "    Schema, Språk, Ärendenummer. Utöver de här 7 kolumnerna\n"
                    + "    ska också 31 dagar finnas med, en kolumn för varje dag.\n"
                    + "    Alltså 7 + 31 kolumner, totalt.\n\n "
                    + "2. Om kolumner saknar kategorinamn räknas de\n"
                    + "     inte som kolumner, så se till att kategorinamn\n"
                    + "     finns (se kategorinamnen ovan).\n\n"
                    + "3. Redigera din Excelfil och starta om programmet.\n\n"
                    + "För mer info tryck på Hjälp-knappen och läs 'Attendance Tool - hur fungerar det.pdf'.\n\n\n", "Meddelande", JOptionPane.DEFAULT_OPTION);
            Desktop dt = Desktop.getDesktop();
            try {
                dt.open(file);
                System.exit(0);

            } catch (IOException ex) {
                grupper.add("Fel i Excelen");
                groups = new String[grupper.size()];
                for (int t = 0; t < grupper.size(); t++) {
                    groups[t] = grupper.get(t);
                }
                return groups;
            }
            grupper.add("Kolumn fel");
            groups = new String[grupper.size()];
            for (int t = 0; t < grupper.size(); t++) {
                groups[t] = grupper.get(t);
            }
            return groups;
        }
        Cell gruppCell = null;
        Cell nCell = null;
        boolean existAlready;
        int noGroups = 0;
        for (int i = 1; i <= worksheet.getLastRowNum(); i++) {
            existAlready = false;
            if (!isRowEmpty(worksheet.getRow(i))) {
                try {
                    gruppCell = worksheet.getRow(i).getCell(1);
                    nCell = worksheet.getRow(i).getCell(0);
                } catch (NullPointerException e) {
                    continue;
                }

                if (gruppCell.toString().isEmpty() || gruppCell.toString().equals(" ") || nCell.toString().isEmpty() || nCell.toString().equals(" ")) {
                    continue;
                }
                for (int t = 0; t < grupper.size(); t++) {
                    if (gruppCell.toString().equals(grupper.get(t))) {
                        existAlready = true;
                        break;
                    }
                }
                if (existAlready == false) {
                    grupper.add(gruppCell.toString());
                }
            }
        }
        if (grupper.isEmpty()) {
            grupper.add("Excelen är tom");
            groups = new String[grupper.size()];
            for (int t = 0; t < grupper.size(); t++) {
                groups[t] = grupper.get(t);
            }
        } else {
            groups = new String[grupper.size()];
            for (int t = 0; t < grupper.size(); t++) {
                groups[t] = grupper.get(t);
            }
        }
        try {
            fsIP.close();
        } catch (IOException ex) {
            grupper.add("Fel i Excelen");
            groups = new String[grupper.size()];
            for (int t = 0; t < grupper.size(); t++) {
                groups[t] = grupper.get(t);
            }
            return groups;
        }
        return groups;
    }

    public void changeParticipantInfo(String name, String newName, String start, String end, String time, String caseNbr, String language, String group, String groupNow) {
        File file = new File(filePath);
        if(time.toLowerCase().trim().equals("individuellt schema")) {
            time = "Deltid";
        }
        try {
            fsIP = new FileInputStream(file); //Read the spreadsheet that needs to be updated
        } catch (FileNotFoundException ex) {
            Logger.getLogger(ExcelHandler.class.getName()).log(Level.SEVERE, null, ex);
        }
        try {
            wb = new HSSFWorkbook(fsIP); //Access the workbook
        } catch (IOException ex) {
            Logger.getLogger(ExcelHandler.class.getName()).log(Level.SEVERE, null, ex);
        }
        worksheet = wb.getSheetAt(0); //Access the worksheet, so that we can update / modify it.
        wb.setMissingCellPolicy(Row.CREATE_NULL_AS_BLANK);

        int infoIndex = getRowIndexByNameAndGroup(name, groupNow);

        worksheet.getRow(infoIndex).getCell(0).setCellValue(newName);
        //System.out.println(groupNow);
        //System.out.println(group);
        if (!groupNow.trim().equals(group.trim())) {
            worksheet.getRow(infoIndex).getCell(1).setCellValue(group);
        }
        worksheet.getRow(infoIndex).getCell(2).setCellValue(start);
        worksheet.getRow(infoIndex).getCell(3).setCellValue(end);
        worksheet.getRow(infoIndex).getCell(4).setCellValue(time);
        worksheet.getRow(infoIndex).getCell(5).setCellValue(language);
        worksheet.getRow(infoIndex).getCell(6).setCellValue(caseNbr);

        try (FileOutputStream output_file = new FileOutputStream(new File(filePath))) //Open FileOutputStream to write updates
        {
            wb.write(output_file); //write changes
            fsIP.close(); //Close the InputStream
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "Excelfilen är öppen. Stäng den och försök sedan igen", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
    }

    public String[] getParticipantInfo(String name, String groupNow, ScheduleDB db) {
        File file = new File(filePath);
        try {
            fsIP = new FileInputStream(file); //Read the spreadsheet that needs to be updated
        } catch (FileNotFoundException ex) {
            Logger.getLogger(ExcelHandler.class.getName()).log(Level.SEVERE, null, ex);
        }
        try {
            wb = new HSSFWorkbook(fsIP); //Access the workbook
        } catch (IOException ex) {
            Logger.getLogger(ExcelHandler.class.getName()).log(Level.SEVERE, null, ex);
        }
        worksheet = wb.getSheetAt(0); //Access the worksheet, so that we can update / modify it.
        wb.setMissingCellPolicy(Row.CREATE_NULL_AS_BLANK);

        Cell timeCell;
        Cell nameCell;
        Cell caseCell;
        Cell startCell;
        Cell endCell;
        Cell languageCell;
        Cell groupCell;

        String[] info = new String[7];

        int infoIndex = getRowIndexByNameAndGroup(name, groupNow);

        nameCell = worksheet.getRow(infoIndex).getCell(0);
        groupCell = worksheet.getRow(infoIndex).getCell(1);
        startCell = worksheet.getRow(infoIndex).getCell(2);
        endCell = worksheet.getRow(infoIndex).getCell(3);
        timeCell = worksheet.getRow(infoIndex).getCell(4);
        languageCell = worksheet.getRow(infoIndex).getCell(5);
        caseCell = worksheet.getRow(infoIndex).getCell(6);

        int diff;
        String caseNbr;

        if (nameCell == null || nameCell.toString().isEmpty()) {
            info[0] = "saknas";
        } else {
            info[0] = nameCell.toString().replace(String.valueOf((char) 160), " ").trim();
        }
        if (startCell == null || startCell.toString().isEmpty()) {
            info[1] = "saknas";
        } else {
            info[1] = startCell.toString().replace(String.valueOf((char) 160), " ").trim();
        }
        if (endCell == null || endCell.toString().isEmpty()) {
            info[2] = "saknas";
        } else {
            info[2] = endCell.toString().replace(String.valueOf((char) 160), " ").trim();
        }

        if (timeCell == null || timeCell.toString().isEmpty()) {
            info[3] = "saknas";
        } else {
            //om individuellt schema finns
            if (db.tableExist(caseCell.toString().trim())) {
                info[3] = "Individuellt schema";
            } else {
                info[3] = timeCell.toString().replace(String.valueOf((char) 160), " ").trim();
            }
        }

        if (caseCell == null || caseCell.toString().isEmpty()) {
            info[4] = "saknas";
        } else {
            caseNbr = caseCell.toString().replace(".", "");
            caseNbr = caseNbr.replace("E7", "").replace(String.valueOf((char) 160), " ").trim();
            //lägg till nollorna som saknas

            if (caseNbr.isEmpty()) {
                caseNbr = "0";
            }

            if (caseNbr.length() < 8 && caseNbr.length() > 0) {
                diff = 8 - caseNbr.length();
                for (int h = 0; h < diff; h++) {
                    caseNbr = caseNbr + "0";
                }
            }

            info[4] = caseNbr;
        }

        if (languageCell == null || languageCell.toString().isEmpty()) {
            info[5] = "saknas";
        } else {
            info[5] = languageCell.toString().replace(String.valueOf((char) 160), " ").trim();
        }

        if (groupCell == null || groupCell.toString().isEmpty()) {
            info[6] = "saknas";
        } else {
            info[6] = groupCell.toString().replace(String.valueOf((char) 160), " ").trim();
        }

        try {
            fsIP.close();
        } catch (IOException ex) {
            Logger.getLogger(ExcelHandler.class.getName()).log(Level.SEVERE, null, ex);
        }
        return info;
    }

    public void searchParticipant(ScheduleDB db) {
        boolean empty = true;
        while (empty == true) {
            String caseNumber = JOptionPane.showInputDialog(null, "Ärendenummer: ", "Sök deltagare", JOptionPane.PLAIN_MESSAGE);
            caseNumber = caseNumber.trim();
            if (caseNumber != null) {

                if (!caseNumber.isEmpty()) {

                    File file = new File(filePath);
                    try {
                        fsIP = new FileInputStream(file); //Read the spreadsheet that needs to be updated
                    } catch (FileNotFoundException ex) {
                        Logger.getLogger(ExcelHandler.class.getName()).log(Level.SEVERE, null, ex);
                    }
                    try {
                        wb = new HSSFWorkbook(fsIP); //Access the workbook
                    } catch (IOException ex) {
                        Logger.getLogger(ExcelHandler.class.getName()).log(Level.SEVERE, null, ex);
                    }
                    worksheet = wb.getSheetAt(0); //Access the worksheet, so that we can update / modify it.
                    wb.setMissingCellPolicy(Row.CREATE_NULL_AS_BLANK);

                    Cell nameCell;
                    Cell startCell;
                    Cell endCell;
                    Cell scheduleCell;
                    Cell grpCell;
                    String schedule;

                    int infoIndex = getRowIndexByCaseNumber(caseNumber);
                    if (infoIndex == 0) {
                        JOptionPane.showMessageDialog(null, "Deltagare med ärendenummer, " + caseNumber + ", finns inte", "Meddelande", JOptionPane.DEFAULT_OPTION);
                    } else {
                        nameCell = worksheet.getRow(infoIndex).getCell(0);
                        grpCell = worksheet.getRow(infoIndex).getCell(1);
                        startCell = worksheet.getRow(infoIndex).getCell(2);
                        endCell = worksheet.getRow(infoIndex).getCell(3);
                        scheduleCell = worksheet.getRow(infoIndex).getCell(4);

                        //om individuellt schema finns
                        if (db.tableExist(caseNumber)) {
                            schedule = "Individuellt schema";
                        } else {
                            scheduleCell = worksheet.getRow(infoIndex).getCell(4);
                            schedule = scheduleCell.toString().trim();
                        }

                        JOptionPane.showMessageDialog(null, "----------------------------------------------------\n"
                                + "Namn: " + nameCell.toString().trim() + "\n"
                                + "----------------------------------------------------\n"
                                + "Grupp: " + grpCell.toString().trim() + "\n"
                                + "----------------------------------------------------\n"
                                + "Startdatum: " + startCell.toString().trim() + "\n"
                                + "----------------------------------------------------\n"
                                + "Slutdatum: " + endCell.toString().trim() + "\n"
                                + "----------------------------------------------------\n"
                                + "Schema: " + schedule + "\n"
                                + "----------------------------------------------------\n"
                                + "\n", "Ärende: " + caseNumber, JOptionPane.DEFAULT_OPTION);
                    }

                    try {
                        fsIP.close();
                    } catch (IOException ex) {
                        Logger.getLogger(ExcelHandler.class.getName()).log(Level.SEVERE, null, ex);
                    }
                } else {
                    JOptionPane.showMessageDialog(null, "Tomt fält!", "Meddelande", JOptionPane.DEFAULT_OPTION);
                }
            } else {
                empty = false;
            }
        }
    }

    public String getEndDateByName(String name) {
        String date = "Datum-fel";
        File file = new File(filePath);
        try {
            fsIP = new FileInputStream(file); //Read the spreadsheet that needs to be updated
        } catch (FileNotFoundException ex) {
            JOptionPane.showMessageDialog(null, "Kontrollera Excelfilen (deltagare.xls) i mappen 'Attendance Tool' → ’Dokument’", "Meddelande", JOptionPane.DEFAULT_OPTION);
            System.exit(1);
        }
        try {
            wb = new HSSFWorkbook(fsIP); //Access the workbook
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "Excelfilen kunde inte hittas.\n"
                    + "Börja med att importera en Excelfil", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }

        worksheet = wb.getSheetAt(0); //Access the worksheet, so that we can update / modify it.
        Cell endDate = null;
        Cell nameCell = null;
        for (int i = 1; i <= worksheet.getLastRowNum(); i++) {
            if (!isRowEmpty(worksheet.getRow(i))) {
                try {
                    nameCell = worksheet.getRow(i).getCell(0);
                } catch (NullPointerException ee) {
                    continue;

                }
                if (nameCell == null || nameCell.toString().isEmpty() || nameCell.toString().equals(" ") || nameCell.toString().equals("")) {
                    continue;
                }
                if (nameCell.toString().trim().equals(name.trim())) {
                    endDate = worksheet.getRow(i).getCell(3);
                    if (endDate == null || endDate.toString().isEmpty() || endDate.toString().equals(" ")) {
                        //System.out.println(endDate);
                        return "Datum-fel";
                    }
                    date = endDate.toString();
                    break;
                }
            }
        }
        try {
            fsIP.close();
        } catch (IOException ex) {
            Logger.getLogger(ExcelHandler.class.getName()).log(Level.SEVERE, null, ex);
        }
        return date;
    }

    public String getStartDateByName(String name) {
        String date = "Datum-fel";
        File file = new File(filePath);
        try {
            fsIP = new FileInputStream(file); //Read the spreadsheet that needs to be updated
        } catch (FileNotFoundException ex) {
            JOptionPane.showMessageDialog(null, "Kontrollera Excelfilen (deltagare.xls) i mappen 'Attendance Tool' → ’Dokument’", "Meddelande", JOptionPane.DEFAULT_OPTION);
            System.exit(1);
        }
        try {
            wb = new HSSFWorkbook(fsIP); //Access the workbook
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "Excelfilen kunde inte hittas.\n"
                    + "Börja med att importera en Excelfil", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }

        worksheet = wb.getSheetAt(0); //Access the worksheet, so that we can update / modify it.
        Cell startDate = null;
        Cell nameCell = null;
        for (int i = 1; i <= worksheet.getLastRowNum(); i++) {
            if (!isRowEmpty(worksheet.getRow(i))) {
                try {
                    nameCell = worksheet.getRow(i).getCell(0);
                } catch (NullPointerException ee) {
                    continue;

                }
                if (nameCell == null || nameCell.toString().isEmpty() || nameCell.toString().equals(" ") || nameCell.toString().equals("")) {
                    continue;
                }
                if (nameCell.toString().trim().equals(name.trim())) {
                    startDate = worksheet.getRow(i).getCell(2); //tredje kolumnen
                    if (startDate == null || startDate.toString().isEmpty() || startDate.toString().equals(" ")) {
                        return "Datum-fel";
                    }
                    date = startDate.toString();
                    break;
                }
            }
        }
        try {
            fsIP.close();
        } catch (IOException ex) {
            Logger.getLogger(ExcelHandler.class.getName()).log(Level.SEVERE, null, ex);
        }
        return date;
    }

    public String getActivity(String name) {
        String time = "fel";
        File file = new File(filePath);
        try {
            fsIP = new FileInputStream(file); //Read the spreadsheet that needs to be updated
        } catch (FileNotFoundException ex) {
            JOptionPane.showMessageDialog(null, "Kontrollera Excelfilen (deltagare.xls) i mappen 'Attendance Tool' → ’Dokument’", "Meddelande", JOptionPane.DEFAULT_OPTION);
            System.exit(1);
        }
        try {
            wb = new HSSFWorkbook(fsIP); //Access the workbook
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "Excelfilen kunde inte hittas.\n"
                    + "Börja med att importera en Excelfil", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }

        worksheet = wb.getSheetAt(0); //Access the worksheet, so that we can update / modify it.
        Cell timeCell = null;
        Cell nameCell = null;
        for (int i = 1; i <= worksheet.getLastRowNum(); i++) {
            if (!isRowEmpty(worksheet.getRow(i))) {
                try {
                    nameCell = worksheet.getRow(i).getCell(0);
                } catch (NullPointerException ee) {
                    continue;

                }
                if (nameCell == null || nameCell.toString().isEmpty() || nameCell.toString().equals(" ") || nameCell.toString().equals("")) {
                    continue;
                }
                if (nameCell.toString().equals(name)) {
                    timeCell = worksheet.getRow(i).getCell(4);
                    if (timeCell == null || timeCell.toString().isEmpty() || timeCell.toString().equals(" ")) {
                        return "tomt";
                    }
                    time = timeCell.toString();
                    break;
                }
            }
        }
        try {
            fsIP.close();
        } catch (IOException ex) {
            Logger.getLogger(ExcelHandler.class.getName()).log(Level.SEVERE, null, ex);
        }
        return time;
    }

    /*
    public void getTheStatistic() {
        //boolean open = false;
        File file = new File(filePath);
        try {
            fsIP = new FileInputStream(file); //Read the spreadsheet that needs to be updated
        } catch (FileNotFoundException ex) {
            JOptionPane.showMessageDialog(null, "Excelfilen kunde inte hittas.\n"
                    + "Börja med att importera en excelfil", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
        try {
            wb = new HSSFWorkbook(fsIP); // Access the workbook
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "Excelfilen är öppen. Stäng den och försök sedan igen", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
        worksheet = wb.getSheetAt(0); //Access the worksheet, so that we can update / modify it.
        cell = null;
        wb.setMissingCellPolicy(Row.CREATE_NULL_AS_BLANK);
        int result = JOptionPane.showConfirmDialog(null, "Tomma celler och felstavningar i Excelen påverkar statistiken.\n"
                + "Den här statistiken är beroende av följande kategorier:\n"
                + "* Namn\n"
                + "* Grupp\n"
                + "* Slutdatum\n"
                + "* Schema\n\n",
                "Meddelande", JOptionPane.DEFAULT_OPTION, JOptionPane.PLAIN_MESSAGE);
        if (result == JOptionPane.YES_OPTION) {
            Cell endDate;
            Cell activity;
            Cell gruppNamn;
            Cell name;
            int dateCounter = 0;
            int presenceCounter = 0;
            int absenceCounter = 0;
            int totalParticipants = 0;
            int countNoStart = 0;
            int totalCount = 0;
            int countFM = 0;
            int countEM = 0;
            int countHeltid = 0;
            int countPraktik = 0;
            int countCancel = 0;
            int countEnd = 0;
            int mixedTime = 0;
            SimpleDateFormat df;
            DateTimeFormatter formatter;
            LocalDate firstDate;
            LocalDate secondDate;
            Calendar cal;
            df = new SimpleDateFormat("yyyy-MM-dd");
            formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");
            cal = Calendar.getInstance();
            String secondLetter;
            firstDate = LocalDate.parse(df.format(cal.getTime()), formatter);
            long daysBetween;
            LocalDate day;
            for (int j = 1; j <= worksheet.getLastRowNum(); j++) {
                endDate = worksheet.getRow(j).getCell(3);
                activity = worksheet.getRow(j).getCell(4);
                gruppNamn = worksheet.getRow(j).getCell(1);
                name = worksheet.getRow(j).getCell(0);

                //räkna inte med deltagare som redan slutat
                if (gruppNamn != null && !gruppNamn.toString().equals(" ") && !gruppNamn.toString().isEmpty() && !gruppNamn.toString().toLowerCase().trim().contains("avbrott")
                        && !gruppNamn.toString().toLowerCase().trim().contains("avslut") && !gruppNamn.toString().toLowerCase().trim().contains("avslutat")
                        && !gruppNamn.toString().toLowerCase().trim().contains("avbrutet") && !gruppNamn.toString().toLowerCase().trim().contains("avbruten")
                        && !gruppNamn.toString().toLowerCase().trim().contains("avslutad") && !gruppNamn.toString().toLowerCase().trim().contains("avbr")
                        && !gruppNamn.toString().toLowerCase().trim().contains("avsl") && !gruppNamn.toString().toLowerCase().trim().contains("cancel")
                        && !gruppNamn.toString().toLowerCase().trim().contains("end")) {
                    totalParticipants++;
                }
                if ((gruppNamn == null || gruppNamn.toString().isEmpty() || gruppNamn.toString().equals(" ")) && (name != null && !name.toString().isEmpty() && !name.toString().equals(" "))) {
                    totalParticipants++;
                }

                if (!gruppNamn.toString().equals(" ") && !gruppNamn.toString().isEmpty() && !gruppNamn.toString().trim().toLowerCase().equals("avbrott")
                        && !gruppNamn.toString().trim().toLowerCase().equals("avslut") && !gruppNamn.toString().trim().toLowerCase().equals("ej start")
                        && !gruppNamn.toString().substring(1).trim().toLowerCase().equals("ej start") && !gruppNamn.toString().trim().toLowerCase().equals("ej startat")
                        && !gruppNamn.toString().substring(1).trim().toLowerCase().equals("ej startat") && !gruppNamn.toString().substring(1).trim().toLowerCase().equals("avbrott")
                        && !gruppNamn.toString().substring(1).trim().toLowerCase().equals("avslut") && !gruppNamn.toString().trim().toLowerCase().equals("avbrutit")
                        && !gruppNamn.toString().substring(1).trim().toLowerCase().equals("avbrutit") && !gruppNamn.toString().trim().toLowerCase().equals("avslutad")
                        && !gruppNamn.toString().substring(1).trim().toLowerCase().equals("avslutad")) {

                    if (activity != null && activity.getCellType() != Cell.CELL_TYPE_BLANK && !activity.toString().equals(" ")
                            && !activity.toString().isEmpty() && !gruppNamn.toString().trim().toLowerCase().equals("praktik")
                            && !gruppNamn.toString().trim().toLowerCase().equals("p")) {

                        switch (activity.toString().replaceAll("\\s+", "").toLowerCase()) {
                            case "fm":
                                countFM++;
                                break;
                            case "em":
                                countEM++;
                                break;
                            case "deltid":
                                mixedTime++;
                                break;
                            case "fm/em":
                                mixedTime++;
                                break;
                            case "heltid":
                                countHeltid++;
                                break;
                            case "em/fm":
                                mixedTime++;
                                break;
                        }

                    }
                    if (gruppNamn.toString().toLowerCase().trim().equals("praktik") || gruppNamn.toString().toLowerCase().trim().equals("p")) {
                        countPraktik++;
                    }
                } else if (!gruppNamn.toString().isEmpty() && !gruppNamn.toString().equals(" ")) {
                    if ((gruppNamn.toString().trim().toLowerCase().equals("avbrott")
                            || gruppNamn.toString().substring(1).trim().toLowerCase().equals("avbrott")
                            || gruppNamn.toString().trim().toLowerCase().equals("avbrutit")
                            || gruppNamn.toString().substring(1).trim().toLowerCase().equals("avbrutit"))) {
                        countCancel++;
                    } else if ((gruppNamn.toString().trim().toLowerCase().equals("avslut")
                            || gruppNamn.toString().substring(1).trim().toLowerCase().equals("avslut")
                            || gruppNamn.toString().trim().toLowerCase().equals("avslutad")
                            || gruppNamn.toString().substring(1).trim().toLowerCase().equals("avslutad"))) {
                        countEnd++;
                    } else if (gruppNamn.toString().trim().toLowerCase().equals("ej start")
                            || gruppNamn.toString().substring(1).trim().toLowerCase().equals("ej start")
                            || gruppNamn.toString().trim().toLowerCase().equals("ej startat")
                            || gruppNamn.toString().substring(1).trim().toLowerCase().equals("ej startat")) {
                        countNoStart++;
                    }
                }
                //räkna inte med lör och sön.
                if (endDate != null && endDate.getCellType() != Cell.CELL_TYPE_BLANK
                        && !endDate.toString().equals(" ") && !endDate.toString().isEmpty()) {

                    secondDate = LocalDate.parse(String.valueOf(endDate.toString()).trim(), formatter);
                    switch (secondDate.getDayOfWeek().name()) {
                        case "SUNDAY":
                            secondDate = secondDate.minusDays(2);
                            totalCount--;
                            presenceCounter--;
                            break;
                        case "SATURDAY":
                            secondDate = secondDate.minusDays(1);
                            totalCount--;
                            presenceCounter--;
                            break;
                    }
                    daysBetween = ChronoUnit.DAYS.between(firstDate, secondDate);

                    if (daysBetween >= 0 && daysBetween < 22) {
                        dateCounter++;
                    }

                }
                for (int i = 1; i < 32; i++) {
                    cell = worksheet.getRow(j).getCell(i + 6);
                    if (cell == null || cell.toString().isEmpty()) {
                        continue;
                    }
                    secondLetter = String.valueOf(cell.toString().replaceAll("\\s+", "").trim());
                    if (secondLetter.toLowerCase().trim().equals("h") || secondLetter.toLowerCase().trim().equals("d")
                            || secondLetter.toLowerCase().trim().equals("so") || secondLetter.toLowerCase().trim().equals("so+sfi")
                            || secondLetter.toLowerCase().trim().equals("sfi+so")) {
                        absenceCounter++;
                        totalCount++;
                    } else {
                        day = LocalDate.parse(String.valueOf(Integer.toString(i)).trim(), formatter);
                        //räkna inte med helgdagar
                        switch (day.getDayOfWeek().name()) {
                            case "SUNDAY":
                                continue;

                            case "SATURDAY":
                                continue;
                        }
                        presenceCounter++;
                        totalCount++;
                    }
                }
            }
            //ta bort lördagar och söndagar, ca 7 dagar
            int presencePercent = Math.round((presenceCounter * 100.0f) / totalCount);
            int absencePercent = Math.round((absenceCounter * 100.0f) / totalCount);

            JOptionPane.showMessageDialog(null, "Antal deltagare: " + totalParticipants + "\n"
                    + "Närvaro: " + presencePercent + " %" + "\n"
                    + "Frånvaro: " + absencePercent + " %" + "\n"
                    + "Slutdatum nära: " + dateCounter + "\n"
                    + "Avslut: " + countEnd + "\n"
                    + "Avbrott: " + countCancel + "\n"
                    + "Ej startat: " + countNoStart + "\n"
                    + "Praktik: " + countPraktik + "\n"
                    + "FM-are: " + countFM + "\n"
                    + "EM-are: " + countEM + "\n"
                    + "Heltid-are: " + countHeltid + "\n"
                    + "Individuellt schema: " + mixedTime + "\n\n", "Deltagarstatistik", JOptionPane.DEFAULT_OPTION);
        }
    }
     */
    public boolean clearExcelData() throws IOException, FileNotFoundException, ParseException {
        boolean open = false;
        File file = new File(filePath);
        try {
            fsIP = new FileInputStream(file); //Read the spreadsheet that needs to be updated
        } catch (FileNotFoundException ex) {

        }
        try {
            wb = new HSSFWorkbook(fsIP); //Access the workbook
        } catch (IOException ex) {

            open = true;
        }
        worksheet = wb.getSheetAt(0); //Access the worksheet, so that we can update / modify it.
        cell = null;
        wb.setMissingCellPolicy(Row.CREATE_NULL_AS_BLANK);
        JOptionPane pane = new JOptionPane("Gammal närvarodata kommer nu att sparas och närvaro-filen kommer att nollställas.\n"
                + "Under den här processen kommer programmet inte att vara tillgängligt (minimeras).\n"
                + "Tidslängden beror på filstorleken, men det brukar ta några sekunder.\n\n");
        JDialog dialog = pane.createDialog(null, "Meddelande");
        dialog.setDefaultCloseOperation(WindowConstants.DO_NOTHING_ON_CLOSE);
        dialog.setVisible(true);
        GUI.frame.setTitle("Nollställer...");
        GUI.frame.setState(Frame.ICONIFIED);

        for (int j = 1; j <= worksheet.getLastRowNum(); j++) {
            for (int i = 1; i < 32; i++) {
                cell = worksheet.getRow(j).getCell(i + 6);

                if (cell.toString().isEmpty() || cell == null || cell.toString().equals(" ")) {
                    continue;
                }
                cell.setCellValue(" ");

                try (FileOutputStream output_file = new FileOutputStream(new File(filePath))) //Open FileOutputStream to write updates
                {
                    wb.write(output_file); //write changes
                    fsIP.close(); //Close the InputStream
                } catch (IOException ex) {
                    JOptionPane.showMessageDialog(null, "Excelfilen är öppen", "Meddelande", JOptionPane.DEFAULT_OPTION);
                    open = true;
                    break;
                }
            }
        }
        pane = new JOptionPane("Filen har nollställts och närvaro-filen har nu sparats \n"
                + "i mappen 'Attendance Tool' → ’Dokument’ → ’Tidigare Excel’ \n\n");
        GUI.frame.setState(Frame.NORMAL);
        GUI.frame.setTitle("© Attendance Tool (1.1.0)  2020    ¤ E ¤    ( burim333@gmail.com )  ");
        dialog = pane.createDialog(null, "Meddelande");
        dialog.setDefaultCloseOperation(WindowConstants.DO_NOTHING_ON_CLOSE);
        dialog.setVisible(true);

        if (open == true) {
            return false;
        } else {
            return true;
        }
    }

    //lägg in värden i excelfilen. Värden som ska läggas in kommer från GUI classen (getAlltInfoFromCheckBoxes).
    public boolean addValuesToExcel(Map<String, List<String>> hmap, String date, String group) throws IOException, FileNotFoundException, ParseException {
        //tt = new WordModifier(group);
        boolean open = false;
        String[] parts = date.split("-");
        String day = parts[2].trim();
        //om datumet skulle bli fel (för att vara säker)
        if (date.equals("Dagens datum") || date.isEmpty() || date.equals("") || date.equals(" ")) {
            dayNumber = now.get(Calendar.DAY_OF_MONTH);//välj dagens datum om användaren inte valt ett datum    
        } else {
            dayNumber = Integer.parseInt(day);//om användaren valt ett datum
        }
        File file = new File(filePath);
        try {
            fsIP = new FileInputStream(file); //Read the spreadsheet that needs to be updated
        } catch (FileNotFoundException ex) {
            JOptionPane.showMessageDialog(null, "Kontrollera Excelfilen (deltagare.xls) i mappen 'Attendance Tool' → ’Dokument’", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
        try {
            wb = new HSSFWorkbook(fsIP); //Access the workbook
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "Excelfilen kunde inte hittas.\n"
                    + "Börja med att importera en Excelfil", "Meddelande", JOptionPane.DEFAULT_OPTION);
            open = true;
        }
        worksheet = wb.getSheetAt(0); //Access the worksheet, so that we can update / modify it.
        cell = null;
        wb.setMissingCellPolicy(Row.CREATE_NULL_AS_BLANK);

        // iterate and display values
        List<String> values;
        for (Map.Entry<String, List<String>> entry : hmap.entrySet()) {
            String key = entry.getKey();
            values = entry.getValue();
            row = getRowIndexByNameAndGroup(key, group);
            cell = worksheet.getRow(row).getCell(dayNumber + 6);
            //wordOpen = tt.addScheduleToWord(key, group, "sfi","sfi", "sfi", "sfi", "har sfi på em");
            if (wordOpen) {
                open = true;
                break;
            }
            if (cell == null) {
                continue;
            }
            if (!cell.toString().equals("-")) {
                cell.setCellValue(values.get(0));
            }
            try (FileOutputStream output_file = new FileOutputStream(new File(filePath))) //Open FileOutputStream to write updates
            {
                wb.write(output_file); //write changes
                fsIP.close(); //Close the InputStream
            } catch (IOException ex) {
                JOptionPane.showMessageDialog(null, "Excelfilen är öppen. Stäng den och försök sedan igen", "Meddelande", JOptionPane.DEFAULT_OPTION);
                open = true;
                break;
            }
        }
        Desktop desktop = Desktop.getDesktop();
        File file2 = new File(filePath);

        if (open == true) {
            return false;
        } else {
            return true;
        }
    }

    public boolean setActivity(String caseNbr, String activity) throws IOException, FileNotFoundException, ParseException {
        boolean open = false;

        File file = new File(filePath);
        try {
            fsIP = new FileInputStream(file); //Read the spreadsheet that needs to be updated
        } catch (FileNotFoundException ex) {
            JOptionPane.showMessageDialog(null, "Kontrollera Excelfilen (deltagare.xls) i mappen 'Attendance Tool' → ’Dokument’", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
        try {
            wb = new HSSFWorkbook(fsIP); //Access the workbook
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "Excelfilen kunde inte hittas.\n"
                    + "Börja med att importera en Excelfil", "Meddelande", JOptionPane.DEFAULT_OPTION);
            open = true;
        }
        worksheet = wb.getSheetAt(0); //Access the worksheet, so that we can update / modify it.
        cell = null;
        wb.setMissingCellPolicy(Row.CREATE_NULL_AS_BLANK);

        row = getRowIndexByCaseNumber(caseNbr);
        cell = worksheet.getRow(row).getCell(4);//ärendenummer-cell

        if (cell != null) {
            cell.setCellValue(activity);
        }
        try (FileOutputStream output_file = new FileOutputStream(new File(filePath))) //Open FileOutputStream to write updates
        {
            wb.write(output_file); //write changes
            fsIP.close(); //Close the InputStream
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "Excelfilen är öppen. Stäng den och försök sedan igen", "Meddelande", JOptionPane.DEFAULT_OPTION);
            open = true;
        }
        if (open == true) {
            return false;
        } else {
            return true;
        }
    }

    //hämta värden från excelet för att lägga in de i pdf:et
    public void getValuesFromExcel(String[] namn, String group, String raportMonth, String raportYear) {
        try {
            File file = new File(filePath);
            fsIP = new FileInputStream(file); //Read the spreadsheet that needs to be updated
            wb = new HSSFWorkbook(fsIP); //Access the workbook
            worksheet = wb.getSheetAt(0); //Access the worksheet, so that we can update / modify it.
            cell = null;
            wb.setMissingCellPolicy(Row.CREATE_NULL_AS_BLANK);
            String numericVal = "";
            int tempNumVal = 0;
            String firstLetter = "";
            String secondLetter = "";
            String thirdLetter = "";
            String schedule;
            boolean firstTime = true;
            for (int ii = 0; ii < namn.length; ii++) {
                schedule = getActivity(namn[ii]).replaceAll("\\s+", "").trim().toLowerCase();
                if ((schedule.equals("tomt") || (!schedule.equals("heltid") && !schedule.equals("fm")
                        && !schedule.equals("em") && !schedule.equals("deltid")
                        && !schedule.equals("fm/em") && !schedule.equals("em/fm") 
                        && !schedule.equals("individuelltschema"))) && firstTime == true) {
                    System.out.println(schedule);
                    JOptionPane.showMessageDialog(null, "1. Du har ändrat deltagarens namn i programmet (tabellen)!\n\n"
                            + "2. I excelen under kategorin, Aktivitet, finns fel i celler.\n\n "
                            + "Ingen månadsrapport har skapats för berörda deltagare.\n\n\n"
                            + "Punkt 2. Ett av följande alternativen ska finnas med under Aktivitetskolumnen:\n"
                            + "* Heltid\n"
                            + "* FM\n"
                            + "* EM\n"
                            + "* FM/EM\n"
                            + "* Deltid\n" 
                            + "* Individuelltschema\n\n", "Meddelande", JOptionPane.DEFAULT_OPTION);
                    Desktop dt = Desktop.getDesktop();
                    dt.open(file);
                    firstTime = false;
                }
                if ((schedule.equals("tomt") || (!schedule.equals("heltid") && !schedule.equals("fm")
                        && !schedule.equals("em") && !schedule.equals("deltid")
                        && !schedule.equals("fm/em") && !schedule.equals("em/fm")
                        && !schedule.equals("individuelltschema")))) {
                    continue;
                }
                PresenceReportPDF.initializePDF(namn[ii], group);
                PresenceReportPDF.setStaticInfo(namn[ii], raportMonth, raportYear, group);
                row = getRowIndexByNameAndGroup(namn[ii], group);
                //Gå igenom alla dagarna i månaden för raden i fråga
                for (int i = 1; i < 32; i++) {
                    cell = worksheet.getRow(row).getCell(i + 6);
                    //skipa toma celler. Vi vill ha bara celler med värde i
                    if (cell == null) {
                        continue;
                    }
                    if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                        tempNumVal = (int) cell.getNumericCellValue();//double to int
                        numericVal = Integer.toString(tempNumVal).replaceAll("\\s+", "").trim();// int to string
                        firstLetter = String.valueOf(numericVal.charAt(0));
                        if (numericVal.length() > 2 && (firstLetter.toLowerCase().equals("x") || firstLetter.toLowerCase().equals("p"))) {
                            secondLetter = String.valueOf(numericVal.charAt(1));
                            if (secondLetter.equals("-")) {
                                if (numericVal.length() > 4) {
                                    thirdLetter = String.valueOf(numericVal.substring(2));
                                    PresenceReportPDF.addAbsenceToPDF(thirdLetter, i, group, getActivity(namn[ii]));
                                } else {
                                    thirdLetter = String.valueOf(numericVal.charAt(2));
                                    PresenceReportPDF.addAbsenceToPDF(thirdLetter, i, group, getActivity(namn[ii]));
                                }
                            } else {
                                PresenceReportPDF.addAbsenceToPDF(numericVal, i, group, getActivity(namn[ii]));
                            }
                        } else {
                            PresenceReportPDF.addAbsenceToPDF(numericVal, i, group, getActivity(namn[ii]));
                        }
                    } else if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
                        if (!cell.getRichStringCellValue().getString().replaceAll("\\s+", "").trim().isEmpty()) {
                            firstLetter = String.valueOf(cell.getRichStringCellValue().getString().replaceAll("\\s+", "").trim().charAt(0));
                            if (cell.getRichStringCellValue().getString().replaceAll("\\s+", "").trim().length() > 2
                                    && (firstLetter.toLowerCase().equals("x") || firstLetter.toLowerCase().equals("p"))) {
                                secondLetter = String.valueOf(cell.getRichStringCellValue().getString().replaceAll("\\s+", "").trim().charAt(1));
                                if (secondLetter.equals("-")) {
                                    //om t.ex x-0.25
                                    if (cell.getRichStringCellValue().getString().replaceAll("\\s+", "").trim().length() > 4) {
                                        thirdLetter = String.valueOf(cell.getRichStringCellValue().getString().replaceAll("\\s+", "").trim().substring(2));
                                        PresenceReportPDF.addAbsenceToPDF(thirdLetter, i, group, getActivity(namn[ii]));
                                    } else {
                                        thirdLetter = String.valueOf(cell.getRichStringCellValue().getString().replaceAll("\\s+", "").trim().charAt(2));
                                        PresenceReportPDF.addAbsenceToPDF(thirdLetter, i, group, getActivity(namn[ii]));
                                    }
                                } else {
                                    PresenceReportPDF.addAbsenceToPDF(cell.getRichStringCellValue().getString().trim(), i, group, getActivity(namn[ii]));
                                }
                            } else {
                                PresenceReportPDF.addAbsenceToPDF(cell.getRichStringCellValue().getString().replaceAll("\\s+", "").trim(), i, group, getActivity(namn[ii]));
                            }
                        }

                    }
                }
                fsIP.close();
                presRepPDF.openPDF();
            }
            openFolder.openMonthReportsFolder();
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "Excelfilen kunde inte hittas.\n"
                    + "Börja med att importera en Excelfil", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
        PresenceReportPDF.writer.close();
        PresenceReportPDF.reader.close();
        PresenceReportPDF.document.close();
    }

    public boolean getAbsenceFromExcel(String[] namn, String startdatum) throws DocumentException {
        try {
            cell = null;
            wb.setMissingCellPolicy(Row.CREATE_NULL_AS_BLANK);
            String numericVal = "";
            int tempNumVal = 0;
            String absenceVal = "";
            int counterX;
            int counterP;
            boolean isClosed;
            isClosed = abPdf.setStaticInfo(startdatum);
            for (int ii = 0; ii < namn.length; ii++) {
                if (isClosed == false) {
                    PeriodicReportPDF.writer.close();
                    PeriodicReportPDF.reader.close();
                    PeriodicReportPDF.document.close();
                    return false;
                }
                row = getRowIndexByNameAndStartDate(namn[ii], startdatum);
                abPdf.addNameToPDF(namn[ii], ii);
                counterX = 1;
                counterP = 1;
                //Gå igenom alla dagarna i månaden för raden i fråga
                for (int i = 1; i < 32; i++) {
                    cell = worksheet.getRow(row).getCell(i + 6);
                    //skipa toma celler
                    if (cell == null) {
                        continue;
                    }
                    if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                        tempNumVal = (int) cell.getNumericCellValue();//double to int
                        numericVal = Integer.toString(tempNumVal).replaceAll("\\s+", "").trim();// int to string
                        absenceVal = String.valueOf(numericVal).trim();
                        if (!absenceVal.isEmpty()) {
                            if (absenceVal.toLowerCase().equals("p") || absenceVal.toLowerCase().equals("d") || absenceVal.toLowerCase().equals("h")) {
                                abPdf.addAbsenceToPDF(absenceVal, i, ii, counterX, counterP);
                                if (absenceVal.toLowerCase().equals("p")) {
                                    counterP++;
                                } else {
                                    counterX++;
                                }
                            }
                        }
                    } else if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
                        if (!cell.getRichStringCellValue().getString().replaceAll("\\s+", "").trim().isEmpty()) {
                            absenceVal = String.valueOf(cell.getRichStringCellValue().getString().replaceAll("\\s+", "").trim());

                            if (absenceVal.toLowerCase().equals("p") || absenceVal.toLowerCase().equals("d") || absenceVal.toLowerCase().equals("h")) {
                                abPdf.addAbsenceToPDF(absenceVal, i, ii, counterX, counterP);
                                if (absenceVal.toLowerCase().equals("p")) {
                                    counterP++;
                                } else {
                                    counterX++;
                                }
                            }
                        }
                    }
                }
            }
            abPdf.closeAbsencePDF();
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "Excelfilen kunde inte hittas.\n"
                    + "Börja med att importera en Excelfil", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
        PeriodicReportPDF.writer.close();
        PeriodicReportPDF.reader.close();
        PeriodicReportPDF.document.close();
        return true;
    }

    public boolean checkDubleCaseNbr(String namn, String caseNbr) throws IOException {
        File file = new File(filePath);
        try {
            fsIP = new FileInputStream(file); //Read the spreadsheet that needs to be updated
        } catch (FileNotFoundException ex) {
            JOptionPane.showMessageDialog(null, "Kontrollera Excelfilen (deltagare.xls) i mappen 'Attendance Tool' → ’Dokument’", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
        try {
            wb = new HSSFWorkbook(fsIP); //Access the workbook
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "Kontrollera Excelfilen (deltagare.xls) i mappen 'Attendance Tool' → ’Dokument’", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
        worksheet = wb.getSheetAt(0); //Access the worksheet, so that we can update / modify it.
        cell = null;
        Cell caseCell = null;
        String fixedCaseNbr;
        wb.setMissingCellPolicy(Row.CREATE_NULL_AS_BLANK);

        for (int i = 1; i <= worksheet.getLastRowNum(); i++) {
            if (worksheet.getRow(i) != null) {
                cell = worksheet.getRow(i).getCell(0);
                caseCell = worksheet.getRow(i).getCell(6);
                fixedCaseNbr = fixCaseNumber(caseCell.toString());
                if (!cell.toString().trim().equals(namn) && fixedCaseNbr.equals(caseNbr)) {
                    fsIP.close();
                    JOptionPane.showMessageDialog(null, "Deltagaren, " + cell.toString().trim() + ", har samma ärendenummer (" + fixedCaseNbr + ") !\n", "Meddelande", JOptionPane.DEFAULT_OPTION);
                    return true;
                }
            }
        }
        fsIP.close();
        return false;
    }

    public boolean checkForDublicates(String namn, String grupp) throws IOException {
        File file = new File(filePath);
        try {
            fsIP = new FileInputStream(file); //Read the spreadsheet that needs to be updated
        } catch (FileNotFoundException ex) {
            JOptionPane.showMessageDialog(null, "Kontrollera Excelfilen (deltagare.xls) i mappen 'Attendance Tool' → ’Dokument’", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
        try {
            wb = new HSSFWorkbook(fsIP); //Access the workbook
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "Kontrollera Excelfilen (deltagare.xls) i mappen 'Attendance Tool' → ’Dokument’", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
        worksheet = wb.getSheetAt(0); //Access the worksheet, so that we can update / modify it.
        cell = null;
        Cell gruppCell = null;
        wb.setMissingCellPolicy(Row.CREATE_NULL_AS_BLANK);

        for (int i = 1; i <= worksheet.getLastRowNum(); i++) {
            if (worksheet.getRow(i) != null) {
                cell = worksheet.getRow(i).getCell(0);
                gruppCell = worksheet.getRow(i).getCell(1);
                if (cell.toString().trim().equals(namn) && gruppCell.toString().trim().equals(grupp)) {
                    fsIP.close();
                    return true;
                }
            }
        }
        fsIP.close();
        return false;
    }

    //kolla om samma ärendenummer finns redan i excelet (det betyder inte att det är samma deltagare)
    public String getGroupByCase(String caseNbr) throws IOException {
        File file = new File(filePath);
        Cell gruppcell;
        try {
            fsIP = new FileInputStream(file); //Read the spreadsheet that needs to be updated.
        } catch (FileNotFoundException ex) {
            JOptionPane.showMessageDialog(null, "Kontrollera Excelfilen (deltagare.xls) i mappen 'Attendance Tool' → ’Dokument’", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
        try {
            wb = new HSSFWorkbook(fsIP); //Access the workbook
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "Kontrollera Excelfilen (deltagare.xls) i mappen 'Attendance Tool' → ’Dokument’", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
        worksheet = wb.getSheetAt(0); //Access the worksheet, so that we can update / modify it.
        cell = null;
        wb.setMissingCellPolicy(Row.CREATE_NULL_AS_BLANK);
        String fixedCaseNbr = " ";
        for (int i = 1; i <= worksheet.getLastRowNum(); i++) {
            if (worksheet.getRow(i) != null) {
                cell = worksheet.getRow(i).getCell(6);
                fixedCaseNbr = cell.toString();
                fixedCaseNbr = fixCaseNumber(fixedCaseNbr);
                if (fixedCaseNbr.equals(caseNbr)) {
                    gruppcell = worksheet.getRow(i).getCell(1);
                    fsIP.close();
                    return gruppcell.toString().trim();
                }
            }
        }
        fsIP.close();
        return "";
    }

    //kolla om samma ärendenummer finns redan i excelet (det betyder inte att det är samma deltagare)
    public boolean checkDublicates(String caseNbr) throws IOException {
        File file = new File(filePath);
        try {
            fsIP = new FileInputStream(file); //Read the spreadsheet that needs to be updated
        } catch (FileNotFoundException ex) {
            JOptionPane.showMessageDialog(null, "Kontrollera Excelfilen (deltagare.xls) i mappen 'Attendance Tool' → ’Dokument’", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
        try {
            wb = new HSSFWorkbook(fsIP); //Access the workbook
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "Kontrollera Excelfilen (deltagare.xls) i mappen 'Attendance Tool' → ’Dokument’", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
        worksheet = wb.getSheetAt(0); //Access the worksheet, so that we can update / modify it.
        cell = null;
        wb.setMissingCellPolicy(Row.CREATE_NULL_AS_BLANK);
        String fixedCaseNbr = " ";
        for (int i = 1; i <= worksheet.getLastRowNum(); i++) {
            if (worksheet.getRow(i) != null) {
                cell = worksheet.getRow(i).getCell(6);
                fixedCaseNbr = cell.toString();
                fixedCaseNbr = fixCaseNumber(fixedCaseNbr);
                if (fixedCaseNbr.equals(caseNbr)) {
                    fsIP.close();
                    return true;
                }
            }
        }
        fsIP.close();
        return false;
    }

    //Namnen och ärendenumren hämtas för att skapa närvarolistor
    public Map findNameAndCaseByGroup(String grupp) throws IOException {
        File file = new File(filePath);
        List<String> values;
        Map<String, List<String>> hashMap = new HashMap<>();
        try {
            fsIP = new FileInputStream(file); //Read the spreadsheet that needs to be updated
            wb = new HSSFWorkbook(fsIP);
        } catch (FileNotFoundException ex) {
            JOptionPane.showMessageDialog(null, "Excelfilen kunde inte hittas.\n"
                    + "Börja med att importera en Excelfil", "Meddelande", JOptionPane.DEFAULT_OPTION);
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "Excelfilen är öppen. Stäng den och försök sedan igen", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
        worksheet = wb.getSheetAt(0); //Access the worksheet, so that we can update / modify it.
        cell = null;
        wb.setMissingCellPolicy(Row.CREATE_NULL_AS_BLANK);

        Cell caseNbr = null;
        String fixedCaseNbr = " ";
        String n = "";
        for (int i = 1; i <= worksheet.getLastRowNum(); i++) {
            values = new ArrayList<>();
            if (!isRowEmpty(worksheet.getRow(i))) {
                try {
                    cell = worksheet.getRow(i).getCell(0);
                    groupCell = worksheet.getRow(i).getCell(1);
                    caseNbr = worksheet.getRow(i).getCell(6);
                } catch (NullPointerException e) {
                    continue;
                }
                fixedCaseNbr = fixCaseNumber(caseNbr.toString());
                values.add(fixedCaseNbr);

                if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                    hashMap.put("Kolumnfel", values);
                    Desktop dt = Desktop.getDesktop();

                    dt.open(file);
                    fsIP.close();
                    return hashMap;
                } else {
                    n = cell.getStringCellValue();
                }
                //om namnet är tommt men gruppen finns, stanna här och fortsätt i for loopen
                if (cell == null || cell.getCellType() == Cell.CELL_TYPE_BLANK || n.isEmpty() || n.equals(" ")) {
                    continue;
                }
                if (groupCell.toString().equals(grupp)) {
                    hashMap.put(cell.toString(), values);
                }
            } else {
                //vad ska hända om gruppen inte finns?
            }

        }
        fsIP.close();
        return hashMap;
    }

    //Namnen som läggs i listan hämtas av den här metoden. Namnen hämtas direkt från Excel
    public Map findNamesByGroup(String grupp) throws IOException {
        File file = new File(filePath);
        List<String> values;
        Map<String, List<String>> hashMap = new HashMap<>();
        try {
            fsIP = new FileInputStream(file); //Read the spreadsheet that needs to be updated
            wb = new HSSFWorkbook(fsIP);
        } catch (FileNotFoundException ex) {
            JOptionPane.showMessageDialog(null, "Excelfilen kunde inte hittas.\n"
                    + "Börja med att importera en Excelfil", "Meddelande", JOptionPane.DEFAULT_OPTION);
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "Excelfilen är öppen. Stäng den och försök sedan igen", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
        worksheet = wb.getSheetAt(0); //Access the worksheet, so that we can update / modify it.
        cell = null;
        wb.setMissingCellPolicy(Row.CREATE_NULL_AS_BLANK);

        Cell caseNbr = null;
        Cell endDate = null;
        String fixedCaseNbr = " ";
        String n = "";
        for (int i = 1; i <= worksheet.getLastRowNum(); i++) {
            values = new ArrayList<>();
            if (!isRowEmpty(worksheet.getRow(i))) {
                try {
                    cell = worksheet.getRow(i).getCell(0);
                    groupCell = worksheet.getRow(i).getCell(1);
                    caseNbr = worksheet.getRow(i).getCell(6);
                    endDate = worksheet.getRow(i).getCell(3);
                } catch (NullPointerException e) {
                    continue;
                }
                fixedCaseNbr = fixCaseNumber(caseNbr.toString());
                values.add(fixedCaseNbr);
                values.add(endDate.toString());

                if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                    hashMap.put("Kolumnfel", values);
                    Desktop dt = Desktop.getDesktop();

                    dt.open(file);
                    fsIP.close();
                    return hashMap;
                } else {
                    n = cell.getStringCellValue();
                }
                //om namnet är tommt men gruppen finns, stanna här och fortsätt i for loopen
                if (cell == null || cell.getCellType() == Cell.CELL_TYPE_BLANK || n.isEmpty() || n.equals(" ")) {
                    continue;
                }
                if (groupCell.toString().equals(grupp)) {
                    hashMap.put(cell.toString(), values);
                }
            } else {
                //vad ska hända om gruppen inte finns?
            }
        }
        fsIP.close();
        return hashMap;
    }

    Cell groupCell = null;
    Cell caseCell = null;
    Cell startDateCell = null;
    Cell endDateCell = null;

    //lägg till deltagare som ska börja/börjat skolan
    public boolean addParticipant(String name, String group, String caseNbr, String startDate, String endDate, String tid) {
        File file = new File(filePath);
        try {
            fsIP = new FileInputStream(file); //Read the spreadsheet that needs to be updated
            wb = new HSSFWorkbook(fsIP);
        } catch (FileNotFoundException ex) {
            JOptionPane.showMessageDialog(null, "Excelfilen kunde inte hittas.\n"
                    + "Börja med att importera en Excelfil", "Meddelande", JOptionPane.DEFAULT_OPTION);
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "Excelfilen är öppen. Stäng den och försök sedan igen", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
        worksheet = wb.getSheetAt(0); //Access the worksheet, so that we can update / modify it.

        wb.setMissingCellPolicy(Row.CREATE_NULL_AS_BLANK);
        Cell cell = null;
        int count = 0;
        for (int i = 1; i <= worksheet.getLastRowNum(); i++) {
            if (worksheet.getRow(i) == null) {
                worksheet.shiftRows(i + 1, worksheet.getLastRowNum(), -1);
                i--;//Adjusts the sweep in accordance to a row removal
            }
            //hämta grupp från Excel
            cell = worksheet.getRow(i).getCell(1);

            if (cell.toString().equals(group)) {
                count = i;
            }
        }
        if (count == 0) {
            int validRows = 0;
            for (int i = 1; i <= worksheet.getLastRowNum(); i++) {
                if (!(isRowEmpty(worksheet.getRow(i)) || (worksheet.getRow(i) != null))) {
                    validRows++;
                } else {
                    break;
                }
            }
            count = validRows;
        }
        // Get the source / new row
        HSSFRow sourceRow = worksheet.getRow(count);
        HSSFRow newRow = worksheet.getRow(count + 1);
        HSSFCell oldCell;
        HSSFCell newCell;
        HSSFCellStyle newCellStyle;
        // If the row exist in destination, push down all rows by 1 else create a new row
        if (newRow != null) {
            worksheet.shiftRows(count + 1, worksheet.getLastRowNum(), 1);
        } else {
            newRow = worksheet.createRow(count + 1);
        }
        // Loop through source columns to add to new row
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            // Grab a copy of the old/new cell
            oldCell = sourceRow.getCell(i);
            newCell = newRow.createCell(i);

            if (oldCell.getCellType() != Cell.CELL_TYPE_STRING) {
                oldCell.setCellType(Cell.CELL_TYPE_STRING);
            }
            // If the old cell is null jump to next cell
            if (oldCell == null) {
                newCell = null;
                continue;
            }
            // Copy style from old cell and apply to new cell
            newCellStyle = oldCell.getCellStyle();
            newCell.setCellStyle(newCellStyle);

            // If there is a cell comment, copy
            if (oldCell.getCellComment() != null) {
                newCell.setCellComment(oldCell.getCellComment());
            }

            // If there is a cell hyperlink, copy
            if (oldCell.getHyperlink() != null) {
                newCell.setHyperlink(oldCell.getHyperlink());
            }

            // Set the cell data type
            newCell.setCellType(oldCell.getCellType());

            // Set the cell data value
            switch (oldCell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                    newCell.setCellValue(oldCell.getRichStringCellValue());
                    cell = worksheet.getRow(count + 1).getCell(0);
                    cell.setCellValue(name);
                    cell = worksheet.getRow(count + 1).getCell(1);
                    cell.setCellValue(group);
                    cell = worksheet.getRow(count + 1).getCell(2);
                    cell.setCellValue(startDate);
                    cell = worksheet.getRow(count + 1).getCell(3);
                    cell.setCellValue(endDate);
                    cell = worksheet.getRow(count + 1).getCell(5);
                    cell.setCellValue(" ");
                    cell = worksheet.getRow(count + 1).getCell(6);
                    cell.setCellValue(caseNbr);

                    if (!tid.equals("old")) {
                        cell = worksheet.getRow(count + 1).getCell(4);
                        cell.setCellValue(tid);
                    }
                    for (int u = 7; u < 32 + 7; u++) {
                        cell = worksheet.getRow(count + 1).getCell(u);
                        cell.setCellValue(" ");
                    }
                    break;
            }
        }

        // If there are any merged regions in the source row, copy to new row
        for (int i = 0; i < worksheet.getNumMergedRegions(); i++) {
            CellRangeAddress cellRangeAddress = worksheet.getMergedRegion(i);
            if (cellRangeAddress.getFirstRow() == sourceRow.getRowNum()) {
                CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.getRowNum(),
                        (newRow.getRowNum()
                        + (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow())),
                        cellRangeAddress.getFirstColumn(),
                        cellRangeAddress.getLastColumn());
                worksheet.addMergedRegion(newCellRangeAddress);
            }
        }

        try {
            FileOutputStream output_file = new FileOutputStream(new File(filePath)); //Open FileOutputStream to write updates
            wb.write(output_file); //write changes
            fsIP.close();
        } catch (FileNotFoundException ex) {
            JOptionPane.showMessageDialog(null, "Excelfilen är öppen. Stäng den och försök sedan igen", "Meddelande", JOptionPane.DEFAULT_OPTION);
            return true;
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "Excelfilen är öppen. Stäng den och försök sedan igen", "Meddelande", JOptionPane.DEFAULT_OPTION);
            return true;
        }
        return false;
    }

    public static void removeRow(HSSFSheet sheet, int rowIndex) {
        int lastRowNum = sheet.getLastRowNum();
        if (rowIndex >= 0 && rowIndex < lastRowNum) {
            sheet.shiftRows(rowIndex + 1, lastRowNum, -1);
        }
        if (rowIndex == lastRowNum) {
            Row removingRow = sheet.getRow(rowIndex);
            if (removingRow != null) {
                sheet.removeRow(removingRow);
            }
        }
    }

    public boolean deleteParticipantByNameAndGroup(String name, String group) {
        File file = new File(filePath);
        try {
            fsIP = new FileInputStream(file); //Read the spreadsheet that needs to be updated
        } catch (FileNotFoundException ex) {
            JOptionPane.showMessageDialog(null, "Kontrollera Excelfilen (deltagare.xls) i mappen 'Attendance Tool' → ’Dokument’", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
        try {
            wb = new HSSFWorkbook(fsIP); //Access the workbook
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "Kontrollera Excelfilen (deltagare.xls) i mappen 'Attendance Tool' → ’Dokument’", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
        worksheet = wb.getSheetAt(0); //Access the worksheet, so that we can update / modify it.
        cell = null;
        Cell groupCell = null;
        wb.setMissingCellPolicy(Row.CREATE_NULL_AS_BLANK);

        for (int i = 1; i <= worksheet.getLastRowNum(); i++) {
            cell = worksheet.getRow(i).getCell(0);
            groupCell = worksheet.getRow(i).getCell(1);
            if (cell.toString().equals(name) && groupCell.toString().equals(group)) {
                removeRow(worksheet, getOldPartIndex(name, group));

                try (FileOutputStream output_file = new FileOutputStream(new File(filePath))) //Open FileOutputStream to write updates
                {
                    wb.write(output_file); //write changes
                    fsIP.close(); //Close the InputStream
                } catch (IOException ex) {
                    JOptionPane.showMessageDialog(null, "Excelfilen är öppen. Stäng den och försök sedan igen", "Meddelande", JOptionPane.DEFAULT_OPTION);
                    return false;

                }
            }
        }
        return true;
    }

    //ta bort deltagare
    public boolean deleteParticipant(String[] name, String grupp, ScheduleDB db) {
        File file = new File(filePath);
        String caseNbr = "";
        int diff = 0;
        try {
            fsIP = new FileInputStream(file); //Read the spreadsheet that needs to be updated
        } catch (FileNotFoundException ex) {
            JOptionPane.showMessageDialog(null, "Kontrollera Excelfilen (deltagare.xls) i mappen 'Attendance Tool' → ’Dokument’", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
        try {
            wb = new HSSFWorkbook(fsIP); //Access the workbook
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "Kontrollera Excelfilen (deltagare.xls) i mappen 'Attendance Tool' → ’Dokument’", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
        worksheet = wb.getSheetAt(0); //Access the worksheet, so that we can update / modify it.
        cell = null;
        wb.setMissingCellPolicy(Row.CREATE_NULL_AS_BLANK);

        int[] rowIndex = new int[name.length];
        for (int i = 0; i < name.length; i++) {
            rowIndex[i] = getRowIndexByNameAndGroup(name[i], grupp);
            cell = worksheet.getRow(rowIndex[i]).getCell(0);
            if (cell.toString().equals(name[i])) {
                //ta bort det individuella schemat med.
                caseNbr = worksheet.getRow(rowIndex[i]).getCell(6).toString();
                caseNbr = fixCaseNumber(caseNbr);
                if (db.tableExist(caseNbr)) {
                    db.deleteRow(caseNbr);
                }
                removeRow(worksheet, rowIndex[i]);

                try (FileOutputStream output_file = new FileOutputStream(new File(filePath))) //Open FileOutputStream to write updates
                {
                    wb.write(output_file); //write changes
                    fsIP.close(); //Close the InputStream
                } catch (IOException ex) {
                    JOptionPane.showMessageDialog(null, "Excelfilen är öppen. Stäng den och försök sedan igen", "Meddelande", JOptionPane.DEFAULT_OPTION);
                    return false;
                }
            }
        }
        return true;
    }

    public boolean copyRow(String[] name, String group, String group2, String tid) {
        worksheet = wb.getSheetAt(0);
        Cell cell = null;
        Cell nameCell = null;
        int count = 0;
        int oldIndex = getRowIndexByName(name[0]);
        wb.setMissingCellPolicy(Row.CREATE_NULL_AS_BLANK);
        for (int i = 1; i <= worksheet.getLastRowNum(); i++) {
            if (worksheet.getRow(i) == null) {
                worksheet.shiftRows(i + 1, worksheet.getLastRowNum(), -1);
                i--;//Adjusts the sweep in accordance to a row removal
            }
            //hämta grupp från Excel
            cell = worksheet.getRow(i).getCell(1);

            if (cell.toString().equals(group)) {
                count = i;
            }
        }

        // Get the source / new row
        HSSFRow sourceRow = worksheet.getRow(count);
        HSSFRow newRow = worksheet.getRow(count + 1);

        // If the row exist in destination, push down all rows by 1 else create a new row
        if (newRow != null) {
            worksheet.shiftRows(count + 1, worksheet.getLastRowNum(), 1);
        } else {
            newRow = worksheet.createRow(count + 1);
        }

        // Copy style from old cell and apply to new cell
        HSSFCellStyle newCellStyle;
        HSSFCell oldCell;
        HSSFCell newCell;

        // Loop through source columns to add to new row
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            // Grab a copy of the old/new cell
            oldCell = sourceRow.getCell(i);
            newCell = newRow.createCell(i);
            if (oldCell.getCellType() != Cell.CELL_TYPE_STRING) {
                oldCell.setCellType(Cell.CELL_TYPE_STRING);
            }
            // If the old cell is null jump to next cell
            if (oldCell == null) {
                newCell = null;
                continue;
            }
            newCellStyle = oldCell.getCellStyle();
            newCell.setCellStyle(newCellStyle);

            // If there is a cell comment, copy
            if (oldCell.getCellComment() != null) {
                newCell.setCellComment(oldCell.getCellComment());
            }

            // If there is a cell hyperlink, copy
            if (oldCell.getHyperlink() != null) {
                newCell.setHyperlink(oldCell.getHyperlink());
            }

            // Set the cell data type
            newCell.setCellType(oldCell.getCellType());
            newCell.setCellValue(oldCell.getRichStringCellValue());
            cell = worksheet.getRow(count + 1).getCell(0);
        }
        // If there are any merged regions in the source row, copy to new row
        for (int i = 0; i < worksheet.getNumMergedRegions(); i++) {
            CellRangeAddress cellRangeAddress = worksheet.getMergedRegion(i);
            if (cellRangeAddress.getFirstRow() == sourceRow.getRowNum()) {
                CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.getRowNum(),
                        (newRow.getRowNum()
                        + (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow())),
                        cellRangeAddress.getFirstColumn(),
                        cellRangeAddress.getLastColumn());
                worksheet.addMergedRegion(newCellRangeAddress);
            }
        }

        if (oldIndex > getRowIndexByName(worksheet.getRow(count + 1).getCell(0).toString())) {
            oldIndex = oldIndex + 1;
        }
        String caseNbr = " ";
        for (int u = 0; u < 32 + 7; u++) {
            nameCell = worksheet.getRow(getOldPartIndex(name[0], group2)).getCell(u);
            cell = worksheet.getRow(count + 1).getCell(u);
            if (u == 1) {
                cell.setCellValue(group);
            } else if (u == 4) {
                cell.setCellValue(tid);
            } else if (u == 6) {
                caseNbr = nameCell.toString();
                caseNbr = fixCaseNumber(caseNbr);
                cell.setCellValue(caseNbr);
            } else {
                if (u != 0 && u != 2 && u != 3 && u != 5) {

                    cell.setCellStyle(nameCell.getCellStyle());
                }
                cell.setCellValue(nameCell.toString());
            }
        }
        removeRow(worksheet, getOldPartIndex(name[0], group2));
        try {
            FileOutputStream output_file = new FileOutputStream(new File(filePath)); //Open FileOutputStream to write updates
            wb.write(output_file); //write changes
            fsIP.close();
        } catch (FileNotFoundException ex) {
            JOptionPane.showMessageDialog(null, "Excelfilen är öppen. Stäng den och försök sedan igen", "Meddelande", JOptionPane.DEFAULT_OPTION);
            return true;
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "Excelfilen är öppen. Stäng den och försök sedan igen", "Meddelande", JOptionPane.DEFAULT_OPTION);
            return true;
        }
        return false;
    }
}
