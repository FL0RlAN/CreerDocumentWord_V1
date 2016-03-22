import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHdrFtr;

public class CreerPageDeGardeDocWord {

	
	static String docTitre;//todel
	static String docAuteur;//todel
	static String docContext;//todel
	static String docDate;//todel
	static String docURL;//todel
	static String docMessagePersonnaliseeHeader;

public static void ajoutnumberpage(XWPFDocument doc ){
    	
    	CTP ctp = CTP.Factory.newInstance();
    	//this add page number incremental
    	ctp.addNewR().addNewPgNum();
 
    	
    	XWPFParagraph codePara = new XWPFParagraph(ctp, doc);
    	XWPFParagraph[] paragraphs = new XWPFParagraph[1];
    	paragraphs[0] = codePara;
    	//position of number
    	codePara.setAlignment(ParagraphAlignment.RIGHT);
    	CTSectPr sectPr = doc.getDocument().getBody().addNewSectPr();
    	try {
    		XWPFHeaderFooterPolicy hfPolicy = null;
    	    XWPFFooter footer = null;
    	    List<XWPFParagraph> footerParas = null;
    	    XWPFParagraph para = null;
    		hfPolicy = doc.getHeaderFooterPolicy();
    		//if i haven't duplication of footer
    		if(hfPolicy!=null){
                footer = hfPolicy.getFooter(1);
                footerParas = footer.getParagraphs();
                para = footerParas.get(0);
                //if i haven't duplication of footer
    			if(para==null || para.getText()=="" || para.getText()==" " )
    			{
    			    XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(doc, sectPr);
    	    	    headerFooterPolicy.createFooter(STHdrFtr.DEFAULT, paragraphs);
    			}
            }else{
            	 XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(doc, sectPr);
    	    	 headerFooterPolicy.createFooter(STHdrFtr.DEFAULT, paragraphs);
            }

    	} catch (Exception e) {
    	    e.printStackTrace();
    	}	
    }
    
    public static void ajoutheader(XWPFDocument doc,String headerText) throws Exception{
 	   
    	CTP ctp = CTP.Factory.newInstance();
    	//this add page number incremental
   
		CTR ctrHeader = ctp.addNewR();
		CTText ctHeader = ctrHeader.addNewT();
		ctHeader.setStringValue(headerText);	
    	
    	XWPFParagraph codePara = new XWPFParagraph(ctp, doc);
    	XWPFParagraph[] paragraphs = new XWPFParagraph[1];
    	paragraphs[0] = codePara;
    	//position of number
    	codePara.setAlignment(ParagraphAlignment.RIGHT);
    	CTSectPr sectPr = doc.getDocument().getBody().addNewSectPr();
    	try {
    		XWPFHeaderFooterPolicy hfPolicy = null;
    	    XWPFHeader header = null;
    	    List<XWPFParagraph> headerParas = null;
    	    XWPFParagraph para = null;
    		hfPolicy = doc.getHeaderFooterPolicy();
    		//if i haven't duplication of header
    		if(hfPolicy!=null){
                header = hfPolicy.getHeader(1);
                headerParas = header.getParagraphs();
                para = headerParas.get(0);
                //if i haven't duplication of header
    			if(para==null || para.getText()=="" || para.getText()==" " )
    			{
    			    XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(doc, sectPr);
    	    	    headerFooterPolicy.createHeader(STHdrFtr.DEFAULT, paragraphs);
    			}
            }else{
            	 XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(doc, sectPr);
    	    	 headerFooterPolicy.createHeader(STHdrFtr.DEFAULT, paragraphs);
            }

    	} catch (Exception e) {
    	    e.printStackTrace();
    	}	
    }
    /**
     * 
     * @param name
     * @throws Exception
     */
    public static void InsererCaptureEcran(String imgFile,String nameWordsDocument,XWPFDocument doc,XWPFParagraph title,XWPFRun run  ) throws Exception{
    	  
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
                
    		    run.addPicture(is,format, imgFile, Units.toEMU(450), Units.toEMU(300)); // 200x200 pixels
    		    run.addBreak();               // saut de ligne
    		    run.setText("DPS, Digital Product Simulation");
    		    run.setBold(true);
    		    title.setAlignment(ParagraphAlignment.CENTER);
    		    is.close();
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
		if (i>2){
			docTitre=args[0];   //todel
			docAuteur=args[1]; //todel
			docContext=args[2]; //todel
			docDate=args[3];//todel
			docURL=args[4];//todel
			docMessagePersonnaliseeHeader=args[5];
		}else{
			System.out.println("bad arg");
		}		
		
		
		String nameLogo = "/TestKarren/logiciel/image/logoKarren.png";
		String nameWordsDocument = docURL;
		XWPFDocument doc = new XWPFDocument(); 
		XWPFParagraph title = doc.createParagraph();    
		XWPFRun run = title.createRun();

	//	InsererTableau(nameWordsDocument,doc,title);
		//captureEcran(nameScreenShot);
		
		InsererCaptureEcran(nameLogo,nameWordsDocument,doc,title,run);
		run.addBreak();
		run.addBreak();
		run.addBreak();
		run.addBreak();
		run.addBreak();
		run.addBreak();
		run.setText("Nom du rappport des TESTS : "+docTitre);
		run.addBreak();		
		run.setText("Auteur : "+ docAuteur);	
		run.addBreak();		
		run.setText("Date : "+docDate);		
		run.addBreak();				
		run.setText("Context des tests : " +docContext);
		run.setFontSize(18);
		run.setShadow(true);
		run.setFontFamily("Time New Roman");
	    run.setBold(true);
	    
	    
		ajoutnumberpage(doc);

		ajoutheader(doc,docMessagePersonnaliseeHeader);
	    
	    
	    FileOutputStream fos = new FileOutputStream(nameWordsDocument);
		doc.write(fos);
		fos.flush();
		fos.close();   
		System.out.println("fin CreerPageDeGardeWord.java");
		return ;
	}
}
