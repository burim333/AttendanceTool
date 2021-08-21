import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JOptionPane;

/**
 *
 * @author Burim Sadriu Copyright© 2019, All Rights Reserved
 */
public class ScheduleDB {

    private String dbPath;
    private static Connection con;
    private ResultSet resultSet = null;
    private static Statement statement = null;
    private PreparedStatement sm;

    public ScheduleDB() throws ClassNotFoundException {
        try {
            Class.forName("org.sqlite.JDBC");
            dbPath = "jdbc:sqlite:" + System.getProperty("user.home") + "/Skrivbord/Attendance Tool/Dokument/schema.sqlite";
            con = DriverManager.getConnection(dbPath);

            String sql = "CREATE TABLE IF NOT EXISTS schedule (\n"
                    + "	Casenumber text,\n"
                    + "	Monday text,\n"
                    + "	Tuesday text,\n"
                    + "	Wednesday text,\n"
                    + "	Thursday text,\n"
                    + "	Friday text,\n"
                    + "	Note text\n"
                    + ");";

            sm = con.prepareStatement(sql);
            sm.executeUpdate();
            sm.close();

        } catch (SQLException ex) {
            Logger.getLogger(ScheduleDB.class.getName()).log(Level.SEVERE, null, ex);
            JOptionPane.showMessageDialog(null, "Databasen kunde inte hittas.\n"
                    + "Kontrollera att den finns i mappen ’Attendance Tool’ → ’Dokument’.", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
    }

    public boolean tableExist(String caseNbr) {
        String dbName;
        try {
            statement = con.createStatement();
            resultSet = statement.executeQuery("SELECT Casenumber FROM schedule where Casenumber = \"" + caseNbr + "\"");
            dbName = resultSet.getString(1);
            statement.close();
            resultSet.close();
            return true;

        } catch (SQLException ex) {
            return false;
        }

    }

    /*
     public String getPercentage(String caseNbr) throws SQLException {
     String percent;
     statement = con.createStatement();
     resultSet = statement.executeQuery("SELECT Percentage FROM schedule where Casenumber = \"" + caseNbr + "\"");
     percent = resultSet.getString(1);
     statement.close();
     resultSet.close();
     return percent;
     }
     */
    public String getDaySchedule(String day, String caseNbr) throws SQLException {
        String dayPlan;
        statement = con.createStatement();
        resultSet = statement.executeQuery("SELECT " + day + " FROM schedule where Casenumber = \"" + caseNbr + "\"");
        dayPlan = resultSet.getString(1);
        statement.close();
        resultSet.close();
        return dayPlan;
    }

    public String getNote(String caseNbr) throws SQLException {
        String dayPlan;
        statement = con.createStatement();
        resultSet = statement.executeQuery("SELECT Note FROM schedule where Casenumber = \"" + caseNbr + "\"");
        dayPlan = resultSet.getString(1);
        statement.close();
        resultSet.close();
        return dayPlan;
    }

    public void updateSchedule(String caseNbr, String mon, String tue, String wed, String thu, String fri, String note) {
        try {
            String query = "update schedule set Monday=?,Tuesday=?,Wednesday=?,Thursday=?,Friday=?,Note=? where Casenumber = ?";
            sm = con.prepareStatement(query);

            sm.setString(1, mon);
            sm.setString(2, tue);
            sm.setString(3, wed);
            sm.setString(4, thu);
            sm.setString(5, fri);
            sm.setString(6, note);
            sm.setString(7, caseNbr);
            sm.executeUpdate();
            sm.close();

        } catch (SQLException ex) {
            JOptionPane.showMessageDialog(null, "Databasen är upptagen.\n\n"
                    + "Möjligtvis har du öppnat fler fönster av programmet.\n"
                    + "Stäng alla fönster och starta om programmet.\n\n", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
    }

    public void updateCaseNbr(String newCaseNbr, String oldCaseNbr) {
        try {
            String query = "update schedule set Casenumber=? where Casenumber = ?";
            sm = con.prepareStatement(query);
            sm.setString(1, newCaseNbr);
            sm.setString(2, oldCaseNbr);
            sm.executeUpdate();
            sm.close();
        } catch (SQLException ex) {
            JOptionPane.showMessageDialog(null, "Databasen är upptagen.\n\n"
                    + "Möjligtvis har du öppnat fler fönster av programmet.\n"
                    + "Stäng alla fönster och starta om programmet.\n\n", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
    }

    public String[] getWeekSchedule(String caseNbr) {
        String[] schedule = new String[6];
        try {
            sm = con.prepareStatement("select Monday, Tuesday, Wednesday, Thursday, Friday, Note from schedule where Casenumber = ?");
            sm.setString(1, caseNbr);
            resultSet = sm.executeQuery();
            schedule[0] = resultSet.getString(1);//måndag
            if (resultSet.wasNull()) {
                schedule[0] = "välj";
            }
            schedule[1] = resultSet.getString(2);
            schedule[2] = resultSet.getString(3);
            schedule[3] = resultSet.getString(4);
            schedule[4] = resultSet.getString(5);//fredag
            schedule[5] = resultSet.getString(6);//anteckning
            sm.close();
            resultSet.close();
        } catch (SQLException ex) {
            JOptionPane.showMessageDialog(null, "Databasen är upptagen.\n\n"
                    + "Möjligtvis har du öppnat fler fönster av programmet.\n"
                    + "Stäng alla fönster och starta om programmet.\n\n", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }

        return schedule;
    }

    public void createSchedule(String caseNbr, String mon, String tue, String wed, String tho, String fri, String note) {
        try {
            sm = con.prepareStatement(
                    "INSERT INTO schedule (Casenumber,Monday,Tuesday,Wednesday,Thursday,Friday,Note) VALUES(?,?,?,?,?,?,?)");
            sm.setString(1, caseNbr);
            sm.setString(2, mon);
            sm.setString(3, tue);
            sm.setString(4, wed);
            sm.setString(5, tho);
            sm.setString(6, fri);
            sm.setString(7, note);
            sm.executeUpdate();
            sm.close();

        } catch (SQLException ex) {
            JOptionPane.showMessageDialog(null, "Databasen är upptagen.\n\n"
                    + "Möjligtvis har du öppnat fler fönster av programmet.\n"
                    + "Stäng alla fönster och starta om programmet.\n\n", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
    }

    public void deleteRow(String caseNbr) {
        try {
            sm = con.prepareStatement("delete from schedule where Casenumber = ?");
            sm.setString(1, caseNbr);
            sm.executeUpdate();
            sm.close();
        } catch (SQLException ex) {
            JOptionPane.showMessageDialog(null, "Databasen är upptagen.\n\n"
                    + "Möjligtvis har du öppnat fler fönster av programmet.\n"
                    + "Stäng alla fönster och starta om programmet.\n\n", "Meddelande", JOptionPane.DEFAULT_OPTION);
        }
    }
}
