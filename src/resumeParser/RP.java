package resumeParser;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class RP {
	
	private enum MasterResumeLocations
	{
		TABLE_MAIN_HEADER,
		TABLE_HEADER_SUMMARY,
		PARA_SUMMARY,
		TABLE_HIGHLIGHTS,
		TABLE_CORE_COMPETENCIES,
	}
	
	// each experience could have multiple jobs associated with it. Example provided:
	// Walmart & target, worked at both as a 'sales associate', but are technically the same experience.
	private static class CompanyExperience
	{
		String companyName;
		String companyLocation;
		String yearRange;
		ArrayList<CompanyTitles> titles = new ArrayList<>();
	}
	
	// People can have various company titles at each job.
	private static class CompanyTitles
	{
		String title;
		String yearRange;
	}
	
	private static class Experience
	{
		ArrayList<CompanyExperience> companyList = new ArrayList<>();
		ArrayList<String> descriptionList = new ArrayList<>();
	}
	
	private static class Education
	{
		String schoolName;
		String cityStateOrCountry;
		String graduationYear;
		String degreeNameAndMajor;
		ArrayList<String> other = new ArrayList<>();
		
		public String toString()
		{
			String listOfOthers = "";
			for(String o : other)
			{
				listOfOthers = listOfOthers + o + "\n";
			}
			return schoolName + ", " + cityStateOrCountry + ":" + graduationYear + "\n"
					+ degreeNameAndMajor + "\n" + listOfOthers;
		}
	}
	
	private String name;
	private String location;
	private String phone;
	private String email;
	private String linkedInURL;
	private ArrayList<String> headerSummaryList = new ArrayList<>();
	private String summary;
	private ArrayList<String> highlightList = new ArrayList<>();
	private ArrayList<String> coreCompetenciesList = new ArrayList<>();
	
	private ArrayList<Experience> experienceList = new ArrayList<>();
	private ArrayList<Education> educationList = new ArrayList<>();
	
    
    public static void main(String[] args){
        try {
            new RP();
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }
    
    private void extractTemplateData() throws IOException
    {
    	InputStream is = new FileInputStream(new File("Master Resume Template-Revised.docx"));
        XWPFDocument doc = new XWPFDocument(is);
        ArrayList<IBodyElement> bodyEl = new ArrayList<>(doc.getBodyElements());
        
        IBodyElement element;
        XWPFTable table;
        XWPFParagraph para;
        List<XWPFParagraph> tempParaList;
        
        // First, remove all blank paragraphs in the rest of the document
        Iterator<IBodyElement> iter = bodyEl.iterator();
        while(iter.hasNext())
        {
        	element = iter.next();
        	if(element instanceof XWPFParagraph &&
        			((XWPFParagraph) element).getText().trim().isEmpty())
        	{
        		iter.remove();
        	}
        }
        
        for(int i=0; i < bodyEl.size(); i++)
        {
        	// location is important here, so we cannot call remove() here.
        	// we will clear the parsed items once we have gotten everything.
        	element = bodyEl.get(i);
        	
        	// these objects are of static size.
        	if(i < MasterResumeLocations.values().length)
        	{
        		switch(MasterResumeLocations.values()[i])
            	{
    				case TABLE_MAIN_HEADER:
    					// cell 0 contains name, location, phone, email, linkedIn URL
    					table = (XWPFTable) element;
    					tempParaList = table.getRow(0).getCell(0).getParagraphs();
    					name = tempParaList.get(0).getText().trim();
    					location = tempParaList.get(1).getText().trim();
    					phone = tempParaList.get(2).getText().split("\\s+")[0].trim();
    					email = tempParaList.get(2).getText().split("\\s+")[1].trim();
    					linkedInURL = tempParaList.get(3).getText().trim();

    					// cell 1 contains QR code. We will grab that elsewhere
    					break;
    					
    				case TABLE_HEADER_SUMMARY:
    					table = (XWPFTable) element;
    					tempParaList = table.getRow(0).getCell(0).getParagraphs();
    					for(XWPFParagraph parag : tempParaList)
    					{
    						for(String str : parag.getText().split("\\s\\s"))
    						{
    							headerSummaryList.add(str.trim());
    						}
    					}
    					break;
    					
    				case PARA_SUMMARY:
    					para = (XWPFParagraph) element;
    					summary = para.getText();
    					break;
    					
    				case TABLE_HIGHLIGHTS:
    					table = (XWPFTable) element;
    					tempParaList = table.getRow(0).getCell(0).getParagraphs();
    					
    					// start j=1 to skip header
    					for(int j=1; j < tempParaList.size(); j++)
    					{
    						highlightList.add(tempParaList.get(j).getText());
    					}
    					break;
    					
    				case TABLE_CORE_COMPETENCIES:
    					table = (XWPFTable) element;
    					
    					for(int j=1; j < table.getRows().size(); j++)
    					{
    						for(XWPFTableCell cell : table.getRow(j).getTableCells())
    						{
    							coreCompetenciesList.add(cell.getText());
    						}
    					}
    					break;
    					
    				default:
    					break;
            	}
        	}
        	else
        	{
        		// remove everything we have just parsed.
        		for(int j=0; j < i; j++)
        		{
        			bodyEl.remove(0);
        		}
        		break;
        	}
        }
        
        // we have parsed & removed everything up until the dynamic portions of the template.
        // up next is professional experience.
        
        // remove the professional experience header
        bodyEl.remove(0);
        
        iter = bodyEl.iterator();
        
        boolean parsingTitle = true;
        ArrayList<XWPFParagraph> unParsedData = new ArrayList<>();
        
        while(iter.hasNext())
        {
        	element = iter.next();
        	if(element instanceof XWPFTable)
        	{
        		// process the last experience we have and move on.
        		parseExperience(unParsedData);
        		unParsedData.clear();
        		break;
        	}
        	else if(element instanceof XWPFParagraph)
        	{
        		// we've hit a title. There are three cases to consider here:
        		// - we have just started parsing experiences and this is our first one.
        		// - we have just finished parsing a full experience and are about to start a new one
        		// - we are still parsing an experience header
        		if(((XWPFParagraph) element).getAlignment() == ParagraphAlignment.CENTER)
        		{
        			// we are starting to parse a new title. Pass of what we have
        			if(!parsingTitle)
        			{
        				// Send off the previous list for processing.
        				parseExperience(unParsedData);
        				unParsedData.clear();
        				parsingTitle = true;
        			}
        			
        			unParsedData.add((XWPFParagraph) element);
        		}
        		else
        		{
        			parsingTitle = false;
        			unParsedData.add((XWPFParagraph) element);
        		}
        		iter.remove();
        	}
        }
        
        // professional experience is finished. Moving on to education
        
        // remove the education title
        bodyEl.remove(0);
        iter = bodyEl.iterator();
        boolean processEducation = false;
        Education edu = new Education();
        while(iter.hasNext())
        {
        	element = iter.next();
        	if(element instanceof XWPFTable)
        	{
        		// process the last education we have and move on.
        		educationList.add(edu);
        		break;
        	}
        	else if(element instanceof XWPFParagraph)
        	{
        		if(processEducation)
        		{
        			educationList.add(edu);
        			System.out.println(edu.toString());
        			edu = new Education();
        		}
        		// chunks of formatted text are considered 'run' objects contained in XWPFParagraph
        		// we are processing school name, city or state (one whole bolded run)
        		if(((XWPFParagraph) element).getRuns().get(0).isBold() &&
				   ((XWPFParagraph) element).getRuns().size() == 1)
        		{
        			// school name, city, state/country, and year (optional)
        			processEducation = true;
        			String[] strArr = ((XWPFParagraph) element).getText().split(",");
        			edu.schoolName = strArr[0];
        			edu.cityStateOrCountry = strArr[1];
        			// last string will hold state/contry and year
        			strArr = strArr[2].split(":");
        			edu.cityStateOrCountry = edu.cityStateOrCountry + ", " + strArr[0];
        			if(strArr.length == 2)
        			{
        				// we have a year
        				edu.graduationYear = strArr[1];
        			}
        		}
        		else if(((XWPFParagraph) element).getRuns().get(0).isItalic())
        		{
        			// degree shiz
        			edu.degreeNameAndMajor = ((XWPFParagraph) element).getText();
        		}
        		else if(((XWPFParagraph) element).getNumFmt() != null &&
        				((XWPFParagraph) element).getNumFmt().toLowerCase().contains("bullet"))
        		{
        			// varied number of bullets
        			edu.other.add(((XWPFParagraph) element).getText());
        		}
        		
        		iter.remove();
        	}
        }
        
        // additional credentials are next, which is one giant table (with header)
        table = (XWPFTable) bodyEl.remove(0);
        
        // skip header
        for(int i=1; i < table.getRows().size(); i++)
        {
        	// TODO: parse credentials
        }
        
        // last thing left is a note about references. We don't care about this.
    }
    
    public RP() throws IOException
    {
    	// grab data from the template
        extractTemplateData();
    }
    
    private void parseExperience(ArrayList<XWPFParagraph> unparsedExperience)
    {
//    	for(XWPFParagraph para : unparsedExperience)
//    	{
//    		System.out.println(para.getText());
//    	}
    	
    	// TODO: process dynamic experience
    	// first row is company, city,state/country, and year
//		String[] split = ((XWPFParagraph) element).getText().split("\\s\\s");
//		Experience exp = new Experience();
//		exp.companyName = split[0].trim();
//		exp.cityStateOrCountry = split[1].trim();
//		exp.yearRange.add(split[2].trim());
//		experienceList.add(new Experience());
    }
    
    private void printAll(ArrayList<IBodyElement> bodyEl)
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
	              
	              System.out.println(para.getText());
	          }
      }
    }
}