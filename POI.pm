package POI;
use strict; 
use warnings;
use Inline Java => "DATA", SHARED_JVM => 1;

### CONSTRUCTOR
###############################################################
# new()
###############################################################     
sub new {
        my $class    = shift;
        my $proto = shift;        
        return POI::POI->new();
    }
 1;
__DATA__
__Java__

//Import POI classes
import org.apache.poi.hpsf.CustomProperties;
import org.apache.poi.hpsf.DocumentSummaryInformation;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Date;
import org.apache.poi.hpsf.PropertySet;
import org.apache.poi.hpsf.PropertySetFactory;
import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.poifs.filesystem.DirectoryEntry;
import org.apache.poi.poifs.filesystem.DocumentEntry;
import org.apache.poi.poifs.filesystem.DocumentInputStream;
import org.apache.poi.poifs.filesystem.POIFSFileSystem; 


class POI {
    public POI(){
}
/* function for pushing custom variables */

public void PushCustomProperties(String filename, String manname, String docname, String doctitle, String revision, String author) {
	  try {
		File poiFilesystem = new File(filename);
     /* Open the POI filesystem. */
        InputStream is = new FileInputStream(poiFilesystem);
        POIFSFileSystem poifs = new POIFSFileSystem(is);
        is.close();

     /* Read the summary information. */
        DirectoryEntry dir = poifs.getRoot();
		SummaryInformation si;
            
        try {
            DocumentEntry siEntry = (DocumentEntry)
                dir.getEntry(SummaryInformation.DEFAULT_STREAM_NAME);
            DocumentInputStream dis = new DocumentInputStream(siEntry);
            PropertySet ps = new PropertySet(dis);
            dis.close();
            si = new SummaryInformation(ps);
        }
        catch (FileNotFoundException ex) {
            /* There is no summary information yet. We have to create a new
             * one. */
            si = PropertySetFactory.newSummaryInformation();
        }

        /* Change the author to "Pradeep Pant". Any former author value will
         * be lost. If there has been no author yet, it will be created. */
        si.setAuthor("Pradeep Pant");
		si.setTitle(doctitle);
		si.setSubject("Car parts manual");
		si.setComments("Testing manual for making car parts");
		si.setKeywords("automobiles");
		si.setRevNumber("2.0");
		System.out.println("Author changed to " + si.getAuthor() + ".");
		
        /* Read the document summary information. */
        DocumentSummaryInformation dsi;
        try {
            DocumentEntry dsiEntry = (DocumentEntry)
			dir.getEntry(DocumentSummaryInformation.DEFAULT_STREAM_NAME);
            DocumentInputStream dis = new DocumentInputStream(dsiEntry);
            PropertySet ps = new PropertySet(dis);
            dis.close();
            dsi = new DocumentSummaryInformation(ps);
        }
        catch (FileNotFoundException ex) {
            /* There is no document summary information yet. We have to create a
             * new one. */
            dsi = PropertySetFactory.newDocumentSummaryInformation();
        }
        dsi.setCategory("Quality Manual");  
		dsi.setCompany("PRADEEPPANT.COM");
		dsi.setManager("PK PANT");
        CustomProperties cp = dsi.getCustomProperties();
	if (cp == null)
        cp = new CustomProperties();
		cp = new CustomProperties();
		cp.put("Manual Name",manname);
		cp.put("Document Name",docname);
		cp.put ("Document Title",doctitle);
		cp.put("Revision number",revision);
		cp.put("Date", new Date());
    
     /* Write the custom properties back to the document summary information. */
		dsi.setCustomProperties(cp);
		si.write(dir, SummaryInformation.DEFAULT_STREAM_NAME);
		dsi.write(dir, DocumentSummaryInformation.DEFAULT_STREAM_NAME);
	/* Write the POI filesystem back to the original file. Please note that
         * in production code you should never write directly to the origin
         * file! In case of a writing error everything would be lost. */
        OutputStream out = new FileOutputStream(poiFilesystem);
        poifs.writeFilesystem(out);
        out.close();
    
   } // end of try
   catch( Exception e ) {
     e.printStackTrace();
   }   
  } // end of GetCustomInfoForView
} //end of public POI
 
