import javax.swing.*;
import java.awt.*;
import java.awt.event.*;
import java.util.Arrays;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.temporal.ChronoUnit;
import java.util.Calendar;

/**
 *
 * @author Burim Sadriu Copyright© 2020, All Rights Reserved
 */
public class PasswordGUI extends JPanel implements ActionListener {

    private static final String OK = "ok";
    private static final String HELP = "hjälp";
    private static JFrame frame;
    private static ImageIcon img;

    private static SimpleDateFormat df;
    private static DateTimeFormatter formatter;
    private static Calendar cal;
    private static String newLimit = "2019-12-20";

    private final JFrame controllingFrame; //needed for dialogs
    private final JPasswordField passwordField;

    public PasswordGUI(JFrame f) {
        //Use the default FlowLayout.
        controllingFrame = f;
        img = new ImageIcon(getClass().getResource("/Images/icon2.png"));
        //Create everything.
        passwordField = new JPasswordField(10);
        passwordField.setActionCommand(OK);
        passwordField.addActionListener(this);

        JLabel label = new JLabel("Lösenord: ");
        label.setLabelFor(passwordField);

        JComponent buttonPane = createButtonPanel();

        //Lay out everything.
        JPanel textPane = new JPanel(new FlowLayout(FlowLayout.TRAILING));
        textPane.add(label);
        textPane.add(passwordField);
        textPane.setBackground(new Color(0xEDFFE1));
        add(textPane);
        add(buttonPane);
    }

    protected JComponent createButtonPanel() {
        JPanel jpanel = new JPanel(new GridLayout(0, 1));
        JButton okButton = new JButton("OK");
        JButton helpButton = new JButton("Hjälp");

        okButton.setActionCommand(OK);
        helpButton.setActionCommand(HELP);
        helpButton.addActionListener(this);
        okButton.addActionListener(this);

        jpanel.add(okButton);
        jpanel.add(helpButton);

        return jpanel;
    }

    @Override
    public void actionPerformed(ActionEvent e) {
        String cmd = e.getActionCommand();

        if (OK.equals(cmd)) { //Process the password.
            char[] input = passwordField.getPassword();
            if (isPasswordCorrect(input)) {
                frame.dispose();
                GUI.createAndShowUI();
            } else {
                JOptionPane.showMessageDialog(controllingFrame,
                        "Fel lösenord",
                        "Felmeddelande",
                        JOptionPane.DEFAULT_OPTION);
            }

            //Zero out the possible password, for security.
            Arrays.fill(input, '0');

            passwordField.selectAll();
            resetFocus();
        } else { //När användaren frågar om hjälp
            JOptionPane.showMessageDialog(controllingFrame,
                    "Kontakt:\n"
                    + "Burim Sadriu\n"
                    + "0735711195\n"
                    + "burim.sadriu82@gmail.com"
                    + "",
                    "Hjälp",
                    JOptionPane.DEFAULT_OPTION);
        }
    }

    /**
     * Checks the passed-in array against the correct password. After this
     * method returns, you should invoke eraseArray on the passed-in array.
     */
    private static boolean isPasswordCorrect(char[] input) {
        boolean isCorrect = true;
        LocalDate firstDate;
        LocalDate secondDate;

        df = new SimpleDateFormat("yyyy-MM-dd");
        formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");
        cal = Calendar.getInstance();
        firstDate = LocalDate.parse(df.format(cal.getTime()), formatter);
        secondDate = LocalDate.parse(newLimit, formatter);

        long limit = ChronoUnit.DAYS.between(firstDate, secondDate);

        if (limit > 0) {
            char[] password = {'t', 'o', 'o', 'l'};
            if (input.length != password.length) {
                isCorrect = false;
            } else {
                isCorrect = Arrays.equals(input, password);
            }

            //Zero out the password.
            Arrays.fill(password, '0');
        } else {
            JOptionPane.showMessageDialog(null, "Kontakta Burim (burim.sadriu82@gmail.com)\n"
                    + "Användartiden var begränsad i uppdateringsyfte\n\n", "Meddelande", JOptionPane.DEFAULT_OPTION);
            isCorrect = false;

        }
        return isCorrect;
    }

    //Must be called from the event dispatch thread.
    protected void resetFocus() {
        passwordField.requestFocusInWindow();
    }

    /**
     * Create the GUI and show it. For thread safety, this method should be
     * invoked from the event dispatch thread.
     */
    public static void createAndShowPassGUI() {
        //Create and set up the window.
        frame = new JFrame("Attendance Tool");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        //Create and set up the content pane.
        final PasswordGUI newContentPane = new PasswordGUI(frame);
        newContentPane.setOpaque(true); //content panes must be opaque
        frame.setContentPane(newContentPane);

        //Make sure the focus goes to the right component
        //whenever the frame is initially given the focus.
        frame.addWindowListener(new WindowAdapter() {
            public void windowActivated(WindowEvent e) {
                newContentPane.resetFocus();
            }
        });

        //Display the window.
        frame.setSize(300, 100);
        frame.setLocationRelativeTo(null);
        frame.setResizable(false);
        frame.getContentPane().setBackground(new Color(0xEDFFE1));
        frame.setIconImage(img.getImage());
        //frame.pack();
        frame.setVisible(true);
    }
}
