package Excel_classes;

import Gui.*;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.security.GeneralSecurityException;
import java.text.SimpleDateFormat;
import java.util.Iterator;
import javax.swing.JOptionPane;
import javax.swing.JPasswordField;
import org.apache.poi.POIXMLException;
import org.apache.poi.hssf.OldExcelFormatException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Burim Sadriu Copyright© 2020, All Rights Reserved
 */
public class XLS_XMtoXLS {

    private final String outFn = System.getProperty("user.home") + "/Skrivbord/Attendance Tool/Dokument/deltagare.xls";
    private final File inpFn;
    private InputStream in; 
    private boolean cancel = false;
    private Workbook wbOut = null;
    private Workbook wbIn = null;

    public XLS_XMtoXLS(File inpFn) {
        this.inpFn = inpFn;
    }

    public boolean isEncrypted(String path) {
        try {
            try {
                new POIFSFileSystem(new FileInputStream(path));
            } catch (IOException ex) {

            }
            return true;
        } catch (OfficeXmlFileException e) {
            return false;
        }
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

    public void xlsx2xls_progress(String sourcepath) throws InvalidFormatException, IOException, GeneralSecurityException {
        String pass;
        int ok = -1;
        POIFSFileSystem filesystem;
        EncryptionInfo info;
        Decryptor d;
        SimpleDateFormat sdf;
        String cellValue;
        String fixedCaseNbr;
        String password;
        JPasswordField passwordField = new JPasswordField(10);
        int action;

        if (isEncrypted(sourcepath)) {

            while (ok == -1) {
                action = JOptionPane.showConfirmDialog(null, passwordField, "Ange Excel lösenord", JOptionPane.OK_CANCEL_OPTION, JOptionPane.DEFAULT_OPTION);
                password = new String(passwordField.getPassword());
                //System.out.println(action);
                if (action == 0 && password.length() > 0) {
                    pass = password.trim();

                    filesystem = new POIFSFileSystem(new FileInputStream(inpFn));
                    info = new EncryptionInfo(filesystem);
                    d = Decryptor.getInstance(info);

                    if (!d.verifyPassword(pass)) {
                    } else {
                        ok = 2;
                    }
                    try {
                        in = d.getDataStream(filesystem);
                    } catch (Exception e) {
                        JOptionPane.showMessageDialog(null, "Fel lösenord!", "Meddelande", JOptionPane.DEFAULT_OPTION);
                        passwordField.setText("");
                    }
                } else if (password.isEmpty() && action == 0) {
                    JOptionPane.showMessageDialog(null, "Tomt fält!", "Meddelande", JOptionPane.DEFAULT_OPTION);
                } else if (action > 0 || action == -1) {
                    cancel = true;
                    ok = 2;
                }
            }
        } else {
            in = new FileInputStream(inpFn);
        }
        try {
            if (cancel != true) {
                wbIn = new XSSFWorkbook(in);
                File outF = new File(outFn);
                if (outF.exists()) {
                    outF.delete();
                }

                wbOut = new HSSFWorkbook();
                //int sheetCnt = wbIn.getNumberOfSheets();
                Sheet sOut;
                Sheet sIn;
                sIn = wbIn.getSheetAt(0);
                //int sheetIndex;
                Cell cellOut;
                Cell cellIn;
                CellStyle styleIn;
                CellStyle styleOut;
                Row rowIn;
                Row rowOut;
                Iterator<Cell> cellIt;
                Iterator<Row> rowIt;

                //sheetIndex = wbOut.getSheetIndex(sIn.getSheetName());
                //if (sheetIndex == -1) { ?
                sOut = wbOut.createSheet(sIn.getSheetName());
                //}
                rowIt = sIn.rowIterator();

                while (rowIt.hasNext()) {
                    rowIn = rowIt.next();
                    rowOut = sOut.createRow(rowIn.getRowNum());

                    cellIt = rowIn.cellIterator();
                    while (cellIt.hasNext()) {
                        cellIn = cellIt.next();
                        //max 255 kolumner är tillåtna. Vi behöver 38 kolumner (från 0 till 37).
                        if (cellIn.getColumnIndex() > 37) {
                            break;
                        }
                        cellOut = rowOut.createCell(cellIn.getColumnIndex(), cellIn.getCellType());

                        styleIn = cellIn.getCellStyle();
                        styleOut = cellOut.getCellStyle();

                        styleOut.setDataFormat(styleIn.getDataFormat());
                        styleOut.setAlignment(CellStyle.ALIGN_CENTER);
                        cellOut.setCellComment(cellIn.getCellComment());

                        switch (cellIn.getCellType()) {
                            case Cell.CELL_TYPE_BLANK:
                                cellOut.setCellValue(" ");
                                break;

                            case Cell.CELL_TYPE_BOOLEAN:
                                cellOut.setCellValue(cellIn.getBooleanCellValue());
                                //System.out.println("boolean: "+cellIn.getBooleanCellValue());
                                break;

                            case Cell.CELL_TYPE_ERROR:
                                //cellOut.setCellValue(" ");
                                cellOut.setCellValue(cellIn.getErrorCellValue());
                                //System.out.println("error: "+cellIn.getErrorCellValue());
                                break;

                            case Cell.CELL_TYPE_FORMULA:
                                //System.out.println("formula: "+cellIn.getCellFormula());
                                cellOut.setCellFormula(cellIn.getCellFormula());
                                break;

                            case Cell.CELL_TYPE_NUMERIC:
                                if (DateUtil.isCellDateFormatted(cellIn)) {
                                    sdf = new SimpleDateFormat("yyyy-MM-dd");
                                    cellValue = sdf.format(cellIn.getDateCellValue());
                                    cellOut.setCellValue(cellValue.trim());
                                } else {
                                    if (cellIn.getColumnIndex() == 6 && cellIn.getRowIndex()!= 0) {
                                        cellIn.setCellType(Cell.CELL_TYPE_STRING);
                                        cellOut.setCellType(Cell.CELL_TYPE_STRING);
                                        //System.out.println("double to string: " + cellIn.getStringCellValue().trim());
                                        cellOut.setCellValue(fixCaseNumber(cellIn.getStringCellValue().trim()));
                                    } else {
                                        cellOut.setCellValue(cellIn.getNumericCellValue());
                                    }
                                }
                                break;

                            case Cell.CELL_TYPE_STRING:
                                //ärendenumer kolumnet = 6
                                if (cellIn.getColumnIndex() == 6 && cellIn.getRowIndex()!= 0) {
                                    //System.out.println("här: " + cellIn.getStringCellValue().trim());
                                    cellOut.setCellValue(fixCaseNumber(cellIn.getStringCellValue().trim()));
                                } else {
                                    cellOut.setCellValue(cellIn.getStringCellValue().trim());
                                }
                                break;
                        }
                    }
                }

                try (OutputStream out = new BufferedOutputStream(new FileOutputStream(outF))) {
                    wbOut.write(out);
                    //JOptionPane.showMessageDialog(null, "Importen lyckades!\n\n"
                            //+ "Starta om programmet\n\n\n", "Meddelande", JOptionPane.DEFAULT_OPTION);
                    in.close();
                    out.close();
                    GUI.destroyGUI();//stäng GUI
                    GUI.createAndShowUI();//starta om UI
                    //System.exit(0);
                }
            }
        } catch (POIXMLException | OldExcelFormatException c) {
            JOptionPane.showMessageDialog(null, "Se till att excelfilen är sparad i ett av dessa filformat:\n\n"
                    + "Excel 97-2003 Workbook(*.xls)\n"
                    + "Excel-arbetsbok (*.xlsx)\n"
                    + "Excel Macro-Enabled Workbook(*.xlsm)\n\n"
                    + "Försök sedan att importera filen igen.\n\n", "Meddelande", JOptionPane.DEFAULT_OPTION);
            File file = new File(System.getProperty("user.home") + "\\Desktop\\Attendance Tool\\Dokument\\deltagare.xls");
            if (file.exists()) {
                file.delete();// ta bort deltagare.xls då den inte kan öppnas av programmet.
            }
        } finally {
            if (cancel = false) {
                in.close();
            }
        }
    }
}
