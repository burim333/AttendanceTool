import java.awt.Color;
import java.awt.Font;
import javax.swing.UIDefaults;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;

/**
 *
 * @author Burim Sadriu Copyright© 2020, All Rights Reserved
 */
public class Main {

    public static void main(String[] args) {
        try {
            UIManager.put("OptionPane.noButtonText", "Nej"); //ändra yes till ja
            UIManager.put("OptionPane.yesButtonText", "Ja");
            UIManager.put("OptionPane.cancelButtonText", "Avbryt");
            UIManager.put("OptionPane.okButtonText", "Ok");
            /*
            com.jtattoo.plaf.mint.MintLookAndFeel
            com.jtattoo.plaf.mcwin.McWinLookAndFeel
            com.jtattoo.plaf.smart.SmartLookAndFeel
            com.jtattoo.plaf.texture.TextureLookAndFeel
            com.jtattoo.plaf.acryl.AcrylLookAndFeel
            com.jtattoo.plaf.graphite.GraphiteLookAndFeel
             */
            UIManager.setLookAndFeel("com.jtattoo.plaf.mint.MintLookAndFeel");
            UIManager.getLookAndFeelDefaults().put("TableHeader.font", new Font("Arial", Font.PLAIN, 15));// Ã¤ndra text i kategorin
            UIDefaults defaults = UIManager.getLookAndFeelDefaults();
            if (defaults.get("Table.alternateRowColor") == null) {
                defaults.put("Table.alternateRowColor", new Color(240, 240, 240));
            }
        } catch (ClassNotFoundException | InstantiationException | IllegalAccessException | UnsupportedLookAndFeelException ex) {
        }
        //UIManager.put("swing.boldMetal", Boolean.FALSE);
        //PasswordGUI.createAndShowPassGUI();
        GUI.createAndShowUI();
    }
}
