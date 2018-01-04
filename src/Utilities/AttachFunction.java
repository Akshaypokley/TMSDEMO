package Utilities;

import org.openqa.selenium.WebDriver;

import java.awt.*;
import java.awt.datatransfer.StringSelection;
import java.awt.event.KeyEvent;

/**
 * Created by laxmikant on 03/05/2017.
 */
public class AttachFunction {

    WebDriver driver;

    public static void AttachFuntn(WebDriver driver, String FilePath) throws AWTException, InterruptedException {
        StringSelection uploadphoto = new StringSelection(FilePath);
        Toolkit.getDefaultToolkit().getSystemClipboard().setContents(uploadphoto, null);

        try {

            Robot robot = new Robot();

            robot.keyPress(KeyEvent.VK_CONTROL);

            robot.keyPress(KeyEvent.VK_V);

            robot.keyRelease(KeyEvent.VK_V);

            robot.keyRelease(KeyEvent.VK_CONTROL);

            robot.keyPress(KeyEvent.VK_ENTER);

            robot.keyRelease(KeyEvent.VK_ENTER);

        } catch (AWTException e) {
            Robot robot = new Robot();

            robot.keyPress(KeyEvent.VK_CONTROL);
            robot.keyPress(KeyEvent.VK_V);
            robot.keyRelease(KeyEvent.VK_V);
            robot.keyRelease(KeyEvent.VK_CONTROL);
            robot.keyPress(KeyEvent.VK_ENTER);
            robot.keyRelease(KeyEvent.VK_ENTER);

        }


        Thread.sleep(4000);
        //WebDriverWait wait = new WebDriverWait(driver,1000);

    }
}