import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import javax.imageio.ImageIO;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class CopyOfCreerDocWord {

	
	static String docTitre;
	static String docAuteur;
	static String docContexte;
	static String docDate;
	static String docURL;

    
    public static void captureEcran(String name) throws Exception{
    	BufferedImage image;
		image = new Robot().createScreenCapture(new Rectangle(Toolkit.getDefaultToolkit().getScreenSize()));
		ImageIO.write(image, "png", new File("/screenshot2.png"));
    }
    /**
     * 
     * @param name
     * @throws Exception
     */
    public static void InsererCaptureEcran(String imgFile,String nameWordsDocument,XWPFDocument doc,XWPFParagraph title,XWPFRun run  ) throws Exception{
    	  
				run.addBreak(BreakType.PAGE);   //  saut de page 
    		    FileInputStream is = new FileInputStream(imgFile);
    		    
    		    int format=XWPFDocument.PICTURE_TYPE_PNG; // par defaut le format est PNG
                if(imgFile.endsWith(".emf")) format = XWPFDocument.PICTURE_TYPE_EMF;
                else if(imgFile.endsWith(".wmf")) format = XWPFDocument.PICTURE_TYPE_WMF;
                else if(imgFile.endsWith(".pict")) format = XWPFDocument.PICTURE_TYPE_PICT;
                else if(imgFile.endsWith(".jpeg") || imgFile.endsWith(".jpg")) format = XWPFDocument.PICTURE_TYPE_JPEG;
                else if(imgFile.endsWith(".png")) format = XWPFDocument.PICTURE_TYPE_PNG;
                else if(imgFile.endsWith(".dib")) format = XWPFDocument.PICTURE_TYPE_DIB;
                else if(imgFile.endsWith(".gif")) format = XWPFDocument.PICTURE_TYPE_GIF;
                else if(imgFile.endsWith(".tiff")) format = XWPFDocument.PICTURE_TYPE_TIFF;
                else if(imgFile.endsWith(".eps")) format = XWPFDocument.PICTURE_TYPE_EPS;
                else if(imgFile.endsWith(".bmp")) format = XWPFDocument.PICTURE_TYPE_BMP;
                else if(imgFile.endsWith(".wpg")) format = XWPFDocument.PICTURE_TYPE_WPG;
                else {
                    System.err.println("Unsupported picture: " + imgFile +
                            ". Expected emf|wmf|pict|jpeg|png|dib|gif|tiff|eps|bmp|wpg");
                }
                
    		    run.addPicture(is,format, imgFile, Units.toEMU(450), Units.toEMU(450)); // 200x200 pixels
    		    run.addBreak();               // saut de ligne
    		    run.setText("Capture d' écran2");
    		    run.setBold(true);
    		    title.setAlignment(ParagraphAlignment.CENTER);

    		    
    		    is.close();

    		   	  
    			FileOutputStream fos = new FileOutputStream(nameWordsDocument);

    			doc.write(fos);
    			fos.flush();
    			fos.close();       
    }  
    
  
    
    
    public static void InsererTableau(String nameWordsDocument,XWPFDocument document,XWPFParagraph title ) throws Exception{
    	
   	   //Write the Document in file system
   	   FileOutputStream out = new FileOutputStream(nameWordsDocument);
   	        
   	   //create table
   	   XWPFTable table = document.createTable();
   	
   	   //create first row
   	   XWPFTableRow tableRowOne = table.getRow(0);
   	   tableRowOne.getCell(0).setText("Titre");
   	   tableRowOne.addNewTableCell().setText(docTitre);
   
   	   //create second row
   	   XWPFTableRow tableRowTwo = table.createRow();
   	   tableRowTwo.getCell(0).setText("Auteur");
   	   tableRowTwo.getCell(1).setText(docAuteur);

   	   //create third row
   	   XWPFTableRow tableRowThree = table.createRow();
   	   tableRowThree.getCell(0).setText("Contexte");
   	   tableRowThree.getCell(1).setText(docContexte);
   	   
   	   //create third row
   	   XWPFTableRow tableRowFour = table.createRow();
   	   tableRowFour.getCell(0).setText("Date du test :");
   	   tableRowFour.getCell(1).setText(docDate);
		   
   	   document.write(out);
   	
	   	out.flush();
	   	out.close();   
    	
    }
  	  
    

                        
	/**
	 * @param args
	 */
	public static void  main(String[] args)throws Exception {
		// TODO Auto-generated method stub
		System.out.println("***********************************");
		System.out.println("***********************************");
		System.out.println("Ce programme effectue une capture d'ecran et la met dans un document word ");
		System.out.println("Ce programme prend en parametre toutes les informations qui concernent le bug detecte");
		System.out.println("***********************************");
		System.out.println("Attention : il ne peut s'executer que si le document word de bug est fermé");
		System.out.println("***********************************");
		System.out.println("***********************************");
		int i ;
		for( i = 0; i < args.length; i++) {            
            System.out.println("arg["+i+"]="+args[i]);
        }
		if (i>4){
			docTitre=args[0];
			docAuteur=args[1];
			docContexte=args[2];
			docDate=args[3];
			docURL=args[4];
		}else{
			System.out.println("bad arg");
		}		
		
		
		String nameScreenShot = "/logoKarren.png";
		String nameWordsDocument = docURL;
		XWPFDocument doc = new XWPFDocument(); 
		XWPFParagraph title = doc.createParagraph();    
		XWPFRun run = title.createRun();

		
		InsererTableau(nameWordsDocument,doc,title);
		
		captureEcran(nameScreenShot);
		
		InsererCaptureEcran(nameScreenShot,nameWordsDocument,doc,title,run);


		System.out.println("fin");
		return ;
		

		   
            

	}

}
