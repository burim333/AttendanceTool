import java.awt.Color;
import java.awt.Dimension;
import java.awt.event.KeyAdapter;
import java.awt.event.KeyEvent;
import javax.swing.BorderFactory;
import javax.swing.JTextArea;

/**
 *
 * @author Burim Sadriu Copyright© 2019, All Rights Reserved
 */
//Den här classen har hand om intyg.
public class TextFieldLimit extends JTextArea {

    private int limit;

    public TextFieldLimit() {

    }

    public TextFieldLimit(int width, int height, int limit) {
        super.setPreferredSize(new Dimension(width, height));
        this.limit = limit;

        javax.swing.border.Border border = BorderFactory.createLineBorder(Color.GRAY);
        setBorder(BorderFactory.createCompoundBorder(border,
                BorderFactory.createEmptyBorder(10, 10, 10, 10)));

        //behövs för automatisk newline
        setWrapStyleWord(true);
        setLineWrap(true);

        addKeyListener(new KeyAdapter() {

            @Override
            public void keyTyped(KeyEvent k) {
                if (getText().length() >= TextFieldLimit.this.limit) {
                    k.consume();
                }
                if (k.getKeyCode() == KeyEvent.VK_ENTER) {
                    append("\n");
                    //k.notify();
                }
            }
        });
    }
}
