import java.awt.AWTException;
import java.awt.Dimension;
import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;

import javax.imageio.ImageIO;


public class CaptureEcran {

	/**
	 * @param args
	 * @throws AWTException 
	 * @throws IOException 
	 */
	public static void main(String[] args) throws AWTException, IOException {
		Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
  	  Rectangle screenRectangle = new Rectangle(screenSize);
  	  Robot robot = new Robot();
  	  BufferedImage image = robot.createScreenCapture(screenRectangle);
  	  ImageIO.write(image, "png", new File("/screenshot2.png"));
  	  System.out.println("fin");
  	  return;
	}

}
