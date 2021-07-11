import java.awt.*;
import java.awt.event.*;
import java.util.Locale;
import javax.swing.*;

/**
 *
 * @author Burim Sadriu Copyright© 2020, All Rights Reserved
 */
public class DatePicker {

    //define variables
    int month = java.util.Calendar.getInstance().get(java.util.Calendar.MONTH);
    int year = java.util.Calendar.getInstance().get(java.util.Calendar.YEAR);
    java.util.Calendar cal;
    //create object of JLabel with alignment
    JLabel l = new JLabel("", JLabel.CENTER);
    //define variable
    String day = "";
    //declaration
    JDialog d;
    //create object of JButton
    JButton[] button = new JButton[49];

    ImageIcon img = new ImageIcon(getClass().getResource("/Images/icon2.png"));

    public DatePicker(JFrame parent) {
        //create object
        d = new JDialog();
        //set modal true
        d.setModal(true);
        //define string
        String[] header = {"Sön", "Mån", "Tis", "Ons", "Tors", "Fre", "Lör"};
        //create JPanel object and set layout
        JPanel p1 = new JPanel(new GridLayout(7, 7));
        //set size
        p1.setPreferredSize(new Dimension(440, 160));
        //for loop condition
        for (int x = 0; x < button.length; x++) {
            //define variable
            final int selection = x;
            //create object of JButton
            button[x] = new JButton();
            //set focus painted false
            button[x].setFocusPainted(false);
            //set background colour
            button[x].setBackground(Color.white);
            //if loop condition
            if (x > 6) //add action listener
            {
                button[x].addActionListener(new ActionListener() {
                    public void actionPerformed(ActionEvent ae) {
                        day = button[selection].getActionCommand();
                        //call dispose() method
                        d.dispose();
                    }
                });
            }
            if (x < 7)//if loop condition 
            {
                button[x].setText(header[x]);
                //set fore ground colour
                button[x].setForeground(new Color(0xF5231D));
            }
            p1.add(button[x]);//add button
        }
        //create JPanel object with grid layout
        JPanel p2 = new JPanel(new GridLayout(1, 3));

        //create object of button for previous month
        JButton previous = new JButton("<< Tidigare");
        previous.setForeground(new Color(0xF5231D));
        //add action command
        previous.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent ae) {
                //decrement month by 1
                month--;
                //call method
                displayDate();
            }
        });
        previous.setVisible(true);
        p2.add(previous);//add button
        p2.add(l);//add label
        //create object of button for next month
        JButton next = new JButton("Senare >>");
        next.setForeground(new Color(0xF5231D));
        //add action command
        next.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent ae) {
                //increment month by 1
                month++;
                //call method
                displayDate();
            }
        });
        next.setVisible(true);
        p2.add(next);// add next button
        //set border alignment
        d.add(p1, BorderLayout.CENTER);
        d.add(p2, BorderLayout.SOUTH);
        d.setIconImage(img.getImage());
        d.pack();
        //set location
        d.setLocationRelativeTo(parent);
        //call method
        displayDate();
        //set visible true
        d.setVisible(true);
    }

    public void displayDate() {
        for (int x = 7; x < button.length; x++)//for loop
        {
            button[x].setText("");//set text
        }
        //java.text.SimpleDateFormat sdf = new java.text.SimpleDateFormat("yyyy-MM-dd");
        //create object of SimpleDateFormat 
        cal = java.util.Calendar.getInstance();
        //create object of java.util.Calendar 
        cal.set(year, month, 1); //set year, month and date
        //define variables
        int dayOfWeek = cal.get(java.util.Calendar.DAY_OF_WEEK);
        
        //System.out.println(dayOfWeek);
        int daysInMonth = cal.getActualMaximum(java.util.Calendar.DAY_OF_MONTH);
        //condition
        for (int x = 6 + dayOfWeek, day = 1; day <= daysInMonth; x++, day++) //set text
        {
            button[x].setText("" + day);
        }
        String month2 = cal.getDisplayName(cal.MONTH, cal.LONG, Locale.getDefault());
        month2 = month2.substring(0, 1).toUpperCase() + month2.substring(1);
        l.setText(month2 + " " + cal.getWeekYear() );

        d.setTitle("Välj datum");
    }

    public String setPickedDate() {
        if (day.equals("")) {
            return day;
        }
        java.text.SimpleDateFormat sdf = new java.text.SimpleDateFormat("yyyy-MM-dd");
        cal = java.util.Calendar.getInstance();
        cal.set(year, month, Integer.parseInt(day));
        return sdf.format(cal.getTime());
    }
}
