package POIPlus;
use strict; 
use warnings;
use Inline Java => "DATA", SHARED_JVM => 1;
### CONSTRUCTOR

###############################################################
# new()
###############################################################    
sub new {
        my $class    = shift;
        my $mess = shift;
        return POIPlus::POIPlus->new();
    }
 1;

    __DATA__
    __Java__

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xslf.XSLFSlideShow;
import org.apache.poi.POIXMLProperties;
import org.apache.poi.POIXMLProperties.CoreProperties;
import org.apache.poi.POIXMLProperties.CustomProperties;
import org.openxmlformats.schemas.officeDocument.x2006.customProperties.*;

public class POIPlus {
	public POIPlus (){}
	
	// Implementation for docx type
 
 	public String PushProperties_docx(String filename, String docname, String doctitle, String revision, String author) {
	try {
	FileInputStream inputStream = new FileInputStream(filename);
	XWPFDocument document = new XWPFDocument(inputStream);
	inputStream.close();
	POIXMLProperties properties = document.getProperties();	
	POIXMLProperties.CustomProperties customproperties = properties.getCustomProperties(); 	
	customproperties.addProperty("Document name", docname);
	customproperties.addProperty("Document title", doctitle);
	customproperties.addProperty("Revision number", revision);
	customproperties.addProperty("Author", author);		
	FileOutputStream out = new FileOutputStream(filename);
	document.write(out);
	out.close();		   
    }
   catch( Exception e )
    {
     e.printStackTrace();
    }
	return docname;
 } 
 //END of implementation for docx
}