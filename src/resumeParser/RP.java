package resumeParser;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.xmlbeans.XmlCursor;
import resumeParser.MasterResume.Credential;
import resumeParser.MasterResume.CredentialTypes;
import resumeParser.MasterResume.Education;
import resumeParser.MasterResume.MasterResumeLocations;

public class RP {
	
	MasterResume resume;
	
    public static void main(String[] args){
        try {
            new RP();
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }
    
    public RP() throws IOException
    {
    	resume = new MasterResume();
    	// grab data from the template
        extractTemplateData();
        
        // Admin CL Template work
        inputCLTemplate();
        // Admil IL Template work
        inputLITemplate();
    }
    
    private void inputLITemplate() throws IOException
    {
    	InputStream is = new FileInputStream(new File("Admin LI Template.docx"));
    	XWPFDocument doc = new XWPFDocument(is);
    	
    	// “NAME HERE” at the very top should match the name in the Resume & CL 
    	// (it should always capitalize every word, rather than having all uppercase/lowercase).
    	String newText = doc.getParagraphs().get(0).getText().replace("Name Here", resume.getName());
    	doc.getParagraphs().get(0).getRuns().get(0).setText(newText, 0);
    	
    	Utils.printAll(doc.getBodyElements(), true);
    	
    	int headingIndex = 0, backgroundIndex = 0, groupIndex = 0, followingIndex = 0;
    	
    	// lets index the document before continuing.
    	for(XWPFParagraph para : doc.getParagraphs())
    	{
    		// found a major header
    		if(para.getRuns().size() > 0 && 
			   "808080".equals(para.getRuns().get(0).getColor()) &&
			   para.getRuns().get(0).isBold() &&
			   para.getRuns().get(0).getFontSize() == 18)
    		{
    			switch(para.getText().toLowerCase().trim())
    			{
    				case "heading":
    					headingIndex = doc.getPosOfParagraph(para);
    					break;
    				case "background":
    					backgroundIndex = doc.getPosOfParagraph(para);
    					break;
    				case "groups":
    					groupIndex = doc.getPosOfParagraph(para);
    					break;
    				case "following":
    					followingIndex = doc.getPosOfParagraph(para);
    					break;
    			}
    		}
    	}
    	
    	/*Heading*/
    	// Replace “Full Name, Credentials” with name from resume. 
    	// As mentioned elsewhere, the formatting of the destination template needs to be preserved.
    	XWPFTable headingTable = (XWPFTable) doc.getBodyElements().get(headingIndex+2);
    	headingTable.getRow(0).getCell(1).getParagraphs().get(0).getRuns().get(0).setText(resume.getName(), 0);
    	
    	// Replace “Title – List Here” with the first part of the sentence in the resume’s introductory paragraph – 
    	// it should not contain “A” or “An”, and only go as far as the title.  
    	// Each word should capitalized. Example: “Savvy, Dedicated Electrical Engineer”. 
    	// Note that the title in the resume will be either in black or red font, 
    	// but in the LI profile, it should always be black.
    	headingTable.getRow(0).getCell(1).getParagraphs().get(1).getRuns().get(0).setText(resume.getTitle(), 0);
    	
    	/*Background*/
    	for(int i=backgroundIndex; i < groupIndex; i++)
    	{
    		XWPFParagraph para = (XWPFParagraph) doc.getBodyElements().get(i);
    		if(para.getText().equalsIgnoreCase("list summary text here"))
    		{
    			// “List Summary Text Here” should be replaced with the resume introductory paragraph. 
    	    	// Comments should not be preserved.
    			para.getRuns().get(0).getCTR().getRPr().unsetHighlight();
    			para.getRuns().get(0).setText(resume.getSummary(), 0);
    		}
    		else if(para.getText().equalsIgnoreCase("Highlight one"))
    		{
    			// The three highlights below should 
    			// be the first three highlights from the resume.
    			para.getRuns().get(0).getCTR().getRPr().unsetHighlight();
    			para.getRuns().get(0).setText(resume.getHighlights().get(0).getText(), 0);
    			
    			para = (XWPFParagraph) doc.getBodyElements().get(i+1);
    			i++;
    			para.getRuns().get(0).getCTR().getRPr().unsetHighlight();
    			para.getRuns().get(0).setText(resume.getHighlights().get(1).getText(), 0);
    			
    			para = (XWPFParagraph) doc.getBodyElements().get(i+1);
    			i++;
    			para.getRuns().get(0).getCTR().getRPr().unsetHighlight();
    			para.getRuns().get(0).setText(resume.getHighlights().get(2).getText(), 0);
    		}
    		else if(para.getText().equalsIgnoreCase("List ALL Selected Highlights Here"))
    		{
    			para.getRuns().get(0).getCTR().getRPr().unsetHighlight();
    			// The “ALL selected highlights” should be all highlights from the resume.
    			for(XWPFParagraph pa : resume.getHighlights())
    			{
    				XmlCursor cur = para.getCTP().newCursor();
    				XWPFParagraph p = doc.insertNewParagraph(cur);
                	Utils.cloneParagraph(p, para);
                	while(p.getRuns().size() != 0)
                	{
                		p.removeRun(0);
                	}
                	
                	XWPFRun run = p.createRun();
                	run.setText(pa.getText(), 0);
    			}
    			// remove the placeholder
            	doc.removeBodyElement(i + resume.getHighlights().size());
    		}
    	}
    	
    	Utils.saveToFile(doc, "Admin LI Template_export.docx");
    }
    
    private void inputCLTemplate() throws IOException
    {
    	InputStream is = new FileInputStream(new File("Admin CL Template.docx"));
    	XWPFDocument doc = new XWPFDocument(is);
    	
        // The heading containing personal info should be copied over in its entirety 
        // to the corresponding CL section (entire row/cell thingy).
        doc.setTable(0, resume.getPersonalTable());
        
        ArrayList<XWPFParagraph> highlightList = resume.getHighlights();
        // Some resumes do not contain highlights. If a resume does not contain highlights, 
        // the CL should still contain highlights. 
        // Leave the placeholder highlights that are already in the CL template if that’s the case.
        if(highlightList.size() > 0)
        {
        	// Selected highlights should be copied over (this does not include the “Selected Highlights” 
    		// phrase & comment from the resume, nor does it use the gray box thingamajig).
            int beginIndex = -1;
            int endIndex = -1;
            for(int i=0; i < doc.getParagraphs().size(); i++)
            {
            	XWPFParagraph para = doc.getParagraphs().get(i);
            	if(para.getText().contains("Other highlights of my career that exceed expectations of"))
            	{
            		// we've found where the highlights go. Skip the blank line
            		beginIndex = i+1;
            		i++;
            	}
            	else if(beginIndex != -1 && para.getText().trim().isEmpty())
            	{
            		// we've hit the end of the highlight list.
            		endIndex = i-1;
            		break;
            	}
            }
            
            // insert the highlights from the resume
            for(int i=highlightList.size()-1; i >= 0; i--)
            {
            	XmlCursor cur = doc.getParagraphs().get(beginIndex+1).getCTP().newCursor();
            	XWPFParagraph p = doc.insertNewParagraph(cur);
            	Utils.cloneParagraph(p, highlightList.get(i));
            }
            
//            // remove all placeholders
            endIndex += highlightList.size();
            beginIndex += highlightList.size();
            // delete all EXCEPT the last one. The last guy holds the comment we cannot lose.
            for(int i=endIndex-1; i > beginIndex; i--)
            {
            	doc.removeBodyElement(i+2); // add 2 for tables
            }
            
            // The comment in the CL template (RW6) for this section must be preserved.
            // store string of last bullet
            String lastBullet = doc.getParagraphs().get(beginIndex).getText();
            
            // remove all runs except for the last one (containing the comment)
            while(doc.getParagraphs().get(beginIndex+1).getRuns().size() != 1)
            {
            	doc.getParagraphs().get(beginIndex+1).removeRun(0);
            }
            // set text to empty for last run
            doc.getParagraphs().get(beginIndex+1).getRuns().get(0).setText("");
            XWPFRun newRun = doc.getParagraphs().get(beginIndex+1).insertNewRun(0);
            // copy formatting to last bullet
            Utils.cloneRun(newRun, doc.getParagraphs().get(beginIndex).getRuns().get(0));
            // set text
            newRun.setText(lastBullet);
            // remove the previous bullet (keeping the one that preserves the comment)
            doc.removeBodyElement(beginIndex+2);
        }
        
        // The name at the bottom should be copied over from the resume heading – 
        // preferably with each word capitalized, rather than all uppercase/lowercase.
        for(int i=0; i < doc.getParagraphs().size(); i++)
        {
        	if(doc.getParagraphs().get(i).getText().trim().equalsIgnoreCase("name"))
        	{
        		String camelName = StringUtils.capitalize(resume.getName());
        		doc.getParagraphs().get(i).getRuns().get(0).setText(camelName, 0);
        	}
        }
        
        Utils.saveToFile(doc, "Admin CL Template_export.docx");
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
    					resume.setPersonalInfo((XWPFTable) element);
    					// cell 1 contains QR code. We have that stored in the table
    					break;
    					
    				case TABLE_HEADER_SUMMARY:
    					table = (XWPFTable) element;
    					tempParaList = table.getRow(0).getCell(0).getParagraphs();
    					for(XWPFParagraph parag : tempParaList)
    					{
    						for(String str : parag.getText().split("\\s\\s"))
    						{
    							resume.getHeaderSummaryList().add(str.trim());
    						}
    					}
    					break;
    					
    				case PARA_SUMMARY:
    					para = (XWPFParagraph) element;
    					resume.setSummary(para.getText());
    					break;
    					
    				case TABLE_HIGHLIGHTS:
    					table = (XWPFTable) element;
    					tempParaList = table.getRow(0).getCell(0).getParagraphs();
    					
    					// start j=1 to skip header
    					for(int j=1; j < tempParaList.size(); j++)
    					{
    						resume.getHighlights().add(tempParaList.get(j));
    					}
    					break;
    					
    				case TABLE_CORE_COMPETENCIES:
    					table = (XWPFTable) element;
    					
    					for(int j=1; j < table.getRows().size(); j++)
    					{
    						for(XWPFTableCell cell : table.getRow(j).getTableCells())
    						{
    							resume.getCoreCompetencies().add(cell.getText());
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
        		resume.parseExperience(unParsedData);
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
        				resume.parseExperience(unParsedData);
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
        Education edu = null;
        while(iter.hasNext())
        {
        	element = iter.next();
        	if(element instanceof XWPFTable)
        	{
        		// process the last education we have and move on.
        		resume.getEducationList().add(edu);
        		break;
        	}
        	else if(element instanceof XWPFParagraph)
        	{
        		// chunks of formatted text are considered 'run' objects contained in XWPFParagraph
        		// we are processing school name, city or state (one whole bolded run)
        		if(((XWPFParagraph) element).getRuns().get(0).isBold() &&
				   ((XWPFParagraph) element).getRuns().size() == 1)
        		{
        			// we've hit another education listing. If this is not the first one,
        			// add the previous one.
        			if(edu != null)
            		{
            			resume.getEducationList().add(edu);
            			System.out.println(edu);
            		}
        			edu = new Education();
        			// school name, city, state/country, and year (optional)
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
        		else if(((XWPFParagraph) element).getRuns().get(0).getUnderline() == UnderlinePatterns.SINGLE)
        		{
        			//TODO: hit certifications/additional education. What should we do with these?
        		}
        		
        		iter.remove();
        	}
        }
        
        // the last eduation processed has not been added. Toss it on now.
        if(edu != null)
        {
        	resume.getEducationList().add(edu);
        }
        
        // additional credentials are next, which is one giant table (with header)
        table = (XWPFTable) bodyEl.remove(0);
        
        // skip header
        for(int i=1; i < table.getRows().size(); i++)
        {
        	Credential cred = new Credential();
        	
        	// cell 0 is the type of credential
        	for(CredentialTypes type : CredentialTypes.values())
        	{
        		if(type.toString().equalsIgnoreCase(table.getRow(i).getCell(0).getParagraphs().get(0).getText()))
        		{
        			cred.type = type;
        			break;
        		}
        	}
        	
        	// cell 1 is the list of creds
        	for(XWPFParagraph p : table.getRow(i).getCell(1).getParagraphs())
        	{
        		cred.credList.add(p.getText());
        	}
        }
        
        // last thing left is a note about references. We don't care about this.
        doc.close();
    }
}