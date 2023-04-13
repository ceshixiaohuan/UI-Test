package web;

import java.awt.*;
import java.awt.event.KeyEvent;


public class RobotUtils {

    Robot rb;

    public RobotUtils(){
        try {
            rb=new Robot();
        } catch (AWTException e) {
            e.printStackTrace();
        }
    }

    public void moveToclick(int x,int y){
        rb.mouseMove(x,y);
        rb.mousePress(KeyEvent.BUTTON1_DOWN_MASK);
        rb.mouseRelease(KeyEvent.BUTTON1_DOWN_MASK);
        rb.delay(500);
    }

}
