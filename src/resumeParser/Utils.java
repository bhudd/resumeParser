package resumeParser;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Comparator;
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
	
	public static void replaceAll(String target, String replacement, XWPFDocument doc, boolean keepColor)
	{
		for (XWPFParagraph p : doc.getParagraphs()) {
		    List<XWPFRun> runs = p.getRuns();
		    if (runs != null) {
		        for (XWPFRun r : runs) {
		            String text = r.getText(0);
		            if (text != null && text.contains(target)) {
		                text = text.replace(target, replacement);
		                r.setText(text, 0);
		                
		                if(!keepColor)
		                {
		                	r.setColor("000000");
		                }
		            }
		        }
		    }
		}
	}
	
	public static XWPFParagraph findParagraph(String comparer, XWPFDocument doc, int startingIndex, Comparator<String> comp)
	{
		
		for(int i=startingIndex; i < doc.getBodyElements().size(); i++)
		{
			IBodyElement element = doc.getBodyElements().get(i);
			
			if(element instanceof XWPFParagraph)
			{
				if(comp.compare(((XWPFParagraph) element).getText(), comparer) == 0)
				{
					return (XWPFParagraph) element;
				}
			}
		}
		
		// cannot find paragraph
		return null;
	}
	
	public static XWPFParagraph findTableParagraph(String comparer, XWPFDocument doc, int startingIndex, Comparator<String> comp)
	{
		for(int i=startingIndex; i < doc.getBodyElements().size(); i++)
		{
			IBodyElement element = doc.getBodyElements().get(i);
			
			if(element instanceof XWPFTable)
			{
				for(XWPFTableRow row : ((XWPFTable) element).getRows())
				{
					for(XWPFTableCell cell : row.getTableCells())
					{
						for(XWPFParagraph para : cell.getParagraphs())
						{
							if(comp.compare(para.getText(), comparer) == 0)
							{
								return para;
							}
						}
					}
				}
			}
		}
		
		// cannot find paragraph
		return null;
	}
	
	// Comparator definitions to use for strings
	public Comparator<String> COMPARE_EXACT = (String str1, String str2)->str1.equals(str2) ? 0 : 1;
	public Comparator<String> COMPARE_EXACT_IGNORE_CASE = (String str1, String str2)->str1.equalsIgnoreCase(str2) ? 0 : 1;
	public Comparator<String> COMPARE_CONTAINS = (String str1, String str2)->str1.contains(str2) ? 0 : 1;
	public Comparator<String> COMPARE_STARTS_WITH = (String str1, String str2)->str1.startsWith(str2) ? 0 : 1;
	
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
    	System.out.println("File Saved: " + filename);
    }
}
