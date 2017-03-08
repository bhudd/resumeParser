package resumeParser;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;

public class Utils {
	public static void printAll(List<IBodyElement> bodyEl, boolean showMore)
    {
		for(IBodyElement el : bodyEl)
	 	{
			if(el instanceof XWPFTable)
			{
				System.out.println("TABLE PARSED");
				XWPFTable table = (XWPFTable)el;
	      
				for(XWPFTableRow row : table.getRows())
				{
					for(XWPFTableCell cell : row.getTableCells())
					{
						System.out.println(cell.getText());
					}
				}
			}
			else if(el instanceof XWPFParagraph)
			{
				System.out.println("PARAGRAPH PARSED");
				XWPFParagraph para = (XWPFParagraph)el;
				if(showMore)
				{
					System.out.println("NumFmt: " + para.getNumFmt());
					if(para.getRuns().size() > 0)
					{
						System.out.println("first run color: " + para.getRuns().get(0).getColor());
						System.out.println("first run bold: " + para.getRuns().get(0).isBold());
						System.out.println("first run highlighted: " + para.getRuns().get(0).isHighlighted());
					}
				}
				System.out.println(para.getText());
			}
	 	}
    }
    
    public static void cloneParagraph(XWPFParagraph clone, XWPFParagraph source) {
        CTPPr pPr = clone.getCTP().isSetPPr() ? clone.getCTP().getPPr() : clone.getCTP().addNewPPr();
        pPr.set(source.getCTP().getPPr());
        // clear all current runs
        while(clone.getRuns().size() > 0)
        {
        	clone.removeRun(0);
        }
        for (XWPFRun r : source.getRuns()) {
            XWPFRun nr = clone.createRun();
            cloneRun(nr, r);
        }
    }

    public static void cloneRun(XWPFRun clone, XWPFRun source) {
        CTRPr rPr = clone.getCTR().isSetRPr() ? clone.getCTR().getRPr() : clone.getCTR().addNewRPr();
        rPr.set(source.getCTR().getRPr());
    	clone.setText(source.getText(0));
    }
    
    public static void saveToFile(XWPFDocument doc, String filename) throws IOException
    {
    	FileOutputStream out = new FileOutputStream(new File(filename));
    	doc.write(out);
    	out.close();
    	System.out.println("file saved!!");
    }
}
