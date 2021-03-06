package resumeParser;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHyperlinkRun;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRelation;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.xmlbeans.XmlCursor;
import resumeParser.MasterResume.Credential;
import resumeParser.MasterResume.CredentialTypes;
import resumeParser.MasterResume.MasterResumeLocations;

public class RP {
	
	MasterResume resume;
	File resumeFile;
    
    public RP(boolean isMale, File resumeFile) throws IOException
    {
    	this.resumeFile = resumeFile;
    	resume = new MasterResume();
    	// grab data from the template
        extractTemplateData();
        
        // Admin CL Template work
        inputCLTemplate();
        // Admin IL Template work
        inputLITemplate();
        // Thank You Template work
        inputThankYouTemplate();
        // Introduction Template work
        inputIntroductionTemplate(isMale);
    }
    
    private void inputIntroductionTemplate(boolean isMale) throws IOException
    {
    	InputStream is = new FileInputStream(new File("Introduction Template .docx"));
    	XWPFDocument doc = new XWPFDocument(is);
    	
    	// Because these documents use pronouns for gender, I have one for each. 
    	// Some way to select the preferred gender for output would be ideal.
    	Utils.replaceAll("objective pronoun", isMale ? "him" : "her", doc.getParagraphs(), false);
    	Utils.replaceAll("possessive pronoun", isMale ? "his" : "her", doc.getParagraphs(), false);
    	Utils.replaceAll("nominative pronoun", isMale ? "he" : "she", doc.getParagraphs(), false);
    	
    	// Following the trend, you’ll see parts where personal info needs to be entered.  
		// Full name, phone, email, and LinkedIn URL. If the information from the resume 
		// is in red for these parts, it should not be moved/replace info in the template.
    	if(resume.isNameBlack())
    	{
    		// replace all instances of 'Full Name'
    		Utils.replaceAll("Full Name", resume.getName(), doc.getParagraphs(), false);
    		// Under the Introduction Email/InMail section, there is one part which 
    		// will require only the first name, rather than the full name. This is in red.
    		Utils.replaceAll("First name", resume.getName().split("\\s+")[0], doc.getParagraphs(), false);
    	}
    	
    	if(resume.isPhoneBlack())
    	{
    		Utils.replaceAll("Phone", resume.getPhone(), doc.getParagraphs(), false);
    	}
    	
    	if(resume.isEmailBlack())
    	{
    		Utils.replaceAll("email", resume.getEmail(), doc.getParagraphs(), false, "FF0000");
    	}
    	
    	if(resume.islinkedInURLBlack())
    	{
    		Utils.replaceAll("LinkedIn URL", resume.getlinkedInURL(), doc.getParagraphs(), false);
    	}
   
    	// For the Introduction Email/InMail section, the “Job Title” should be 
    	// replaced with the title used in the resume introductory paragraph. 
    	// It should be in black, but should still retain the yellow highlighting.
    	Utils.replaceAll("Job Title", resume.getJobTitle(), doc.getParagraphs(), false);
    	
    	Utils.saveToFile(doc, resumeFile.getParent() + File.separator + "Introduction Template_export.docx");
    }
    
    private void inputThankYouTemplate() throws IOException
    {
    	InputStream is = new FileInputStream(new File("Thank you template.docx"));
    	XWPFDocument doc = new XWPFDocument(is);
    	
    	// For simplicity, this section just needs the personal information 
    	// from the resume moved over, while retaining the format in the thank you template. 
    	// It only needs the information that is shown in the TY template – no LI URL or QR code.
    	Utils.replaceAll("Full Name", resume.getName(), doc.getParagraphs(), true);
    	Utils.replaceAll("Location", resume.getLocation(), doc.getParagraphs(), true);
    	Utils.replaceAll("Phone", resume.getPhone(), doc.getParagraphs(), true);
    	Utils.replaceAll("person@gmail.com", resume.getEmail(), doc.getParagraphs(), true);
    	
    	// update the hyperlink in the document. Only Hyperlinks will be the email
    	String hyperID = doc.getPackagePart().addExternalRelationship("mailto:" + resume.getEmail(), XWPFRelation.HYPERLINK.getRelation()).getId();
    	for(XWPFParagraph para : doc.getParagraphs())
    	{
    		for(XWPFRun run : para.getRuns())
    		{
    			if(run instanceof XWPFHyperlinkRun)
    			{
    				((XWPFHyperlinkRun) run).setHyperlinkId(hyperID);
    			}
    		}
    	}
    	
    	Utils.saveToFile(doc, resumeFile.getParent() + File.separator + "Thank you template_export.docx");
    }
    
    
    private void inputLITemplate() throws IOException
    {
    	InputStream is = new FileInputStream(new File("Admin LI Template.docx"));
    	XWPFDocument doc = new XWPFDocument(is);
    	
    	// “NAME HERE” at the very top should match the name in the Resume & CL 
    	// (it should always capitalize every word, rather than having all uppercase/lowercase).
    	String newText = doc.getParagraphs().get(0).getText().replace("Name Here", resume.getName());
    	doc.getParagraphs().get(0).getRuns().get(0).setText(newText, 0);
    	
//    	Utils.printAll(doc.getBodyElements(), true);
    	
    	int headingIndex = 0, backgroundIndex = 0, groupIndex = 0;
    	
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
    					
					// nothing to input past 'groups'
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
    	// remove all runs but the first one
    	XWPFParagraph para = headingTable.getRow(0).getCell(1).getParagraphs().get(1); 
    	while(para.getRuns().size() != 1)
    	{
    		para.removeRun(1);
    	}
    	para.getRuns().get(0).setText(resume.getTitle(), 0);
    	
    	/*Background*/
    	for(int i=backgroundIndex; i < groupIndex; i++)
    	{
    		para = (XWPFParagraph) doc.getBodyElements().get(i);
    		if(para.getText().equalsIgnoreCase("list summary text here"))
    		{
    			// “List Summary Text Here” should be replaced with the resume introductory paragraph. 
    	    	// Comments should not be preserved.
    			para.getRuns().get(0).getCTR().getRPr().unsetHighlight();
    			para.getRuns().get(0).setText(resume.getSummary(), 0);
    		}
    		else if(para.getText().equalsIgnoreCase("Highlight one"))
    		{
    			// highlights are optional
    			if(resume.getHighlights().size() > 0)
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
    		else if(para.getText().toLowerCase().contains("admin note:"))
    		{
    			// remove everything from here until the end of the last
    			// responsibility
    			for(int j=i; j < i+10; j++)
    			{
//    				System.out.println(((XWPFParagraph)doc.getBodyElements().get(i)).getText());
    				doc.removeBodyElement(i);
    				groupIndex--;
    			}
    			
    			// since we've since deleted the paragraph, reassign it.
    			para = (XWPFParagraph) doc.getBodyElements().get(i);
    			
    			/* Experience */
    	    	// All information from the “Professional Experience” part of the resume should be moved here. 
    	    	// All comments should be omitted. Text color should be preserved.  
    	    	// As mentioned above, LI template formatting does not have to be preserved.
    			for(int j=0; j < resume.getExperiences().size(); j++)
    			{
    				XmlCursor cur = para.getCTP().newCursor();
    				XWPFParagraph p = doc.insertNewParagraph(cur);
    				groupIndex++;
    				Utils.cloneParagraph(p, resume.getExperiences().get(j));
    			}
    		}
    		else if(para.getText().toLowerCase().startsWith("education"))
    		{
    			// move forward 2 spots, as we don't want to remove the first blurb under the title.
    			i += 2;
    			para = (XWPFParagraph) doc.getBodyElements().get(i);
    			/* Education */
    			// Place any information from “Education” section of the resume here. 
    			// Do not preserve comments. If no education section is in the resume, 
    			// leave this section as is.
    			// LI formatting does not need to be preserved.
    			if(resume.getEducationList().size() > 0)
    			{
    				// remove all placeholder garbage
    				for(int j=i+1; j < i+4; j++)
    				{
//        				System.out.println(((XWPFParagraph)doc.getBodyElements().get(i+1)).getText());
    					doc.removeBodyElement(i+1);
    					groupIndex--;
    				}
    				
    				// move paragraph up one step
    				para = (XWPFParagraph) doc.getBodyElements().get(i+1);
    				
    				// insert education list
    				for(int j=0; j < resume.getEducationList().size(); j++)
    				{
    					XmlCursor cur = para.getCTP().newCursor();
        				XWPFParagraph p = doc.insertNewParagraph(cur);
        				groupIndex++;
        				Utils.cloneParagraph(p, resume.getEducationList().get(j));
    				}
    			}
    		}
    		else if(para.getText().toLowerCase().startsWith("certifications"))
    		{
    			
    		}
    		else if(para.getText().toLowerCase().startsWith("skills"))
    		{
    			/* Skills */
    			
    			// Replace placeholder here with information from “Core Competencies”.
    			// First ten bullet points are fine in whatever order, as long as it lists 
    			// it in the format of the LI template’s example filler.
    			
    			// shift forward to the location of the first skill bullet.
    			i += 5;
    			
    			// TODO: check on the formatting of this, see if numbering is OK
    			for(int j=0; j < 10; j++)
    			{
    				// assign para to that location
        			para = (XWPFParagraph) doc.getBodyElements().get(i++);
        			Utils.cloneParagraph(para, resume.getCoreCompetencies().get(j));
    			}
    			
    			// Use the first name from the resume for the part below 
    			// listed skills (capitalized, not all uppercase/lowercase).
    			
    			// skip ahead to the 'First Name' line
    			i += 1;
    			para = (XWPFParagraph) doc.getBodyElements().get(i);
    			
    			// first run contains first name. Replace it
    			para.getRuns().get(0).setText(resume.getName(), 0);
    		}
    		// it's easier to continue searching for this rather than guess how far we should skip ahead
    		// in the 'skills' branch above
    		else if(para.getText().toLowerCase().startsWith("make sure all technical skills are listed after basic skills"))
    		{
    			// For technical skills, copy the technical skills in the resume.
    			
    			// only run we want to keep is the intro one. Remove the rest
    			while(para.getRuns().size() > 1)
    			{
    				para.removeRun(1);
    			}
    			
    			// add space to make it look a bit better
    			XWPFRun run = para.createRun();
    			run.setText(" ",0);
    			para.addRun(run);
    			
    			// find the technical skills credential and add them all.
    			for(Credential cred : resume.getCredentials())
				{
    				if(cred.type.equals(CredentialTypes.TECHNICAL_SKILLS))
    				{
    					for(XWPFParagraph p : cred.credList)
    					{
    						run = para.createRun();
        					run.setText(p.getText(), 0);
        					para.addRun(run);
    					}
    					break;
    				}
    			}
    		}
    		else if(para.getText().toLowerCase().startsWith("projects"))
    		{
    			// do anything?
    		}
    		else if(para.getText().toLowerCase().startsWith("honors & awards") ||
    				para.getText().toLowerCase().startsWith("organizations") ||
    				para.getText().toLowerCase().startsWith("languages") ||
    				para.getText().toLowerCase().startsWith("volunteer") ||
    				para.getText().toLowerCase().startsWith("interests"))
    		{
    			// You’ll see the other sections in the template which correspond 
    			// with Additional Credentials in the resume.  
    			// If the information in the resume’s Additional Credentials is colored red,
    			// the information does not need to be moved or adjusted in the LI template.  
    			// If it is in black, the information should be moved into the LI template 
    			// (it should replace any placeholder info in the LI template.  
    			// This includes the example organizations, certifications, etc.)
    			
    			// determine what credential we are looking for
    			CredentialTypes type;
    			
    			switch(para.getText().toLowerCase())
    			{
    				case "honors & awards":
    					type = CredentialTypes.HONORS_AWARDS;
    					break;
    				case "organizations":
    					type = CredentialTypes.ORGANIZATIONS;
    					break;
    				case "languages":
    					type = CredentialTypes.LANGUAGES;
    					break;
    				case "volunteering experience":
    					type = CredentialTypes.VOLUNTEERING_EXPERIENCE;
    					break;
    				case "interests":
					default:
    					type = CredentialTypes.INTERESTS;
    					break;
    			}
    			
    			// skip blue comments
    			i += 3;
    			
    			for(Credential cred : resume.getCredentials())
    			{
    				if(cred.type == type && cred.isBlack)
    				{
    					// remove what was there
    	    			while(!para.getText().isEmpty())
    	    			{
    	    				doc.removeBodyElement(i);
    	    				groupIndex--;
    	    				para = (XWPFParagraph) doc.getBodyElements().get(i);
    	    			}
    	    			
    	    			// create blank line for spacing purposes
    	    			XmlCursor cur = para.getCTP().newCursor();
        				XWPFParagraph p = doc.insertNewParagraph(cur);
        				// jump ahead one
    	    			para = (XWPFParagraph) doc.getBodyElements().get(i+1);
    	    			
    	    			// add what we have
    	    			for(int j=0; j < cred.credList.size(); j++)
    	    			{
    	    				cur = para.getCTP().newCursor();
            				p = doc.insertNewParagraph(cur);
            				groupIndex++;
            				Utils.cloneParagraph(p, cred.credList.get(j));
    	    			}
    				}
    			}
    		}
    	}
    	
    	Utils.saveToFile(doc, resumeFile.getParent() + File.separator + "Admin LI Template_export.docx");
    }
    
    private void inputCLTemplate() throws IOException
    {
    	InputStream is = new FileInputStream(new File("Admin CL Template.docx"));
    	XWPFDocument doc = new XWPFDocument(is);
    	
        // The heading containing personal info should be copied over in its entirety 
        // to the corresponding CL section (entire row/cell thingy).
    	XWPFTable topTable = doc.getTables().get(0);
    	XWPFTableCell infoCell = topTable.getRow(0).getCell(0);
    	Utils.replaceAll("Name", resume.getName(), infoCell.getParagraphs(), true);
    	Utils.replaceAll("phone", resume.getPhone(), infoCell.getParagraphs(), true);
    	Utils.replaceAll("email", resume.getEmail(), infoCell.getParagraphs(), true);
    	Utils.replaceAll("LinkedIn URL", resume.getlinkedInURL(), infoCell.getParagraphs(), true);
//        doc.setTable(0, resume.getPersonalTable());
    	// TODO: MOVE QR CODE
        XWPFTableCell qrCodeCell = topTable.getRow(0).getCell(1);
        
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
        		doc.getParagraphs().get(i).getRuns().get(0).setText(resume.getName(), 0);
        	}
        }
        
        Utils.saveToFile(doc, resumeFile.getParent() + File.separator + "Admin CL Template_export.docx");
    }
    
    private void extractTemplateData() throws IOException
    {
    	InputStream is = new FileInputStream(resumeFile);
        XWPFDocument doc = new XWPFDocument(is);
        ArrayList<IBodyElement> bodyEl = new ArrayList<>(doc.getBodyElements());
        
        IBodyElement element;
        XWPFTable table;
        XWPFParagraph para;
        List<XWPFParagraph> tempParaList;
        
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
        
        // highlights or core competencies next, depending
        table = (XWPFTable) bodyEl.remove(0);
        
        if(table.getRow(0).getCell(0).getParagraphs().get(0).getText().equalsIgnoreCase("selected highlights"))
        {
        	parseHighlights(table);
        	// remove spacing afterwards
        	bodyEl.remove(0);
        	// core competencies follows
        	parseCoreCompetencies((XWPFTable) bodyEl.remove(0));
        }
        else if(table.getRow(0).getCell(0).getParagraphs().get(0).getText().equalsIgnoreCase("core competencies"))
        {
        	// selected highlights are missing. Just parse the core competencies
        	parseCoreCompetencies(table);
        	// remove spacing afterwards
        	bodyEl.remove(0);
        }
        
        // up next is professional experience or education.
        System.out.println("next thing: " + ((XWPFTable)bodyEl.get(0)).getText());
        
        // remove the professional experience/education header
        String title = ((XWPFTable)bodyEl.remove(0)).getRow(0).getCell(0).getText();
        
        Iterator<IBodyElement> iter = bodyEl.iterator();
        
        if(title.contains("Education"))
        {
        	parseEducation(iter);
        	// remove the next title
            bodyEl.remove(0);
            // reset iterator
            iter = bodyEl.iterator();
            parseExperience(iter);
        }
        else
        {
        	parseExperience(iter);
        	// remove the next title
            bodyEl.remove(0);
            // reset iterator
            iter = bodyEl.iterator();
            parseEducation(iter);
        }
        
        // could be additional credentials
        table = (XWPFTable) bodyEl.get(0);
        
        title = table.getRow(0).getCell(0).getText();
        if(!title.toLowerCase().equals("additional credentials"))
        {
            // remove the table
            bodyEl.remove(0);
            // reset iterator
            iter = bodyEl.iterator();
        	// store any other types of experience that may be before credentials
        	parseExperience(iter);
        }
        
        // next will be additional credentials
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
        		// grab one of the run's colors. 
        		if(p.getRuns().size() > 0)
        		{
        			cred.credList.add(p);
        			cred.isBlack = p.getRuns().get(0).getColor() == "000000" || p.getRuns().get(0).getColor() == null;
        		}
        	}
        	
        	resume.getCredentials().add(cred);
        }
        
        // last thing left is a note about references. We don't care about this.
        doc.close();
    }
    
    private void parseHighlights(XWPFTable table)
    {
    	List<XWPFParagraph> tempParaList = table.getRow(0).getCell(0).getParagraphs();
		
		// start j=1 to skip header
		for(int j=1; j < tempParaList.size(); j++)
		{
			resume.getHighlights().add(tempParaList.get(j));
		}
    }
    
    private void parseCoreCompetencies(XWPFTable table)
    {
    	for(int j=1; j < table.getRows().size(); j++)
		{
			for(XWPFTableCell cell : table.getRow(j).getTableCells())
			{
				resume.getCoreCompetencies().add(cell.getParagraphs().get(0));
			}
		}
    }
    
    private void parseEducation(Iterator<IBodyElement> iter)
    {
        while(iter.hasNext())
        {
        	IBodyElement element = iter.next();
        	if(element instanceof XWPFParagraph)
        	{
        		if(((XWPFParagraph) element).getText().toLowerCase().contains("certifications or additional education:"))
        		{
        			// we have additional certs/ education.
        			// remove certifications title
        	        iter.remove();
        	        
        	        while(iter.hasNext())
        	        {
        	        	element = iter.next();
        	        	if(element instanceof XWPFTable)
        	        	{
        	        		// we've hit the next title. We are finished with certs & add. education
        	        		break;
        	        	}
        	        	else
        	        	{
        	        		resume.getCertificationsList().add((XWPFParagraph) element);
        	        		iter.remove();
        	        	}
        	        }
        	        // at this point, we have finished parsing additional creds. We are finished with education.
            		return;
        		}
        		else
            	{
            		resume.getEducationList().add((XWPFParagraph) element);
            		iter.remove();
            	}
        	}
        	else
        	{
        		// we won't always have additional creds/education. This kicks us
        		// back out if we don't
        		break;
        	}
        }
    }
    
    private void parseExperience(Iterator<IBodyElement> iter)
    {
    	while(iter.hasNext())
        {
        	IBodyElement element = iter.next();
        	if(element instanceof XWPFTable)
        	{
        		// we've hit the next title. We are finished with experience
        		break;
        	}
        	else if(element instanceof XWPFParagraph)
        	{
        		resume.addExperience((XWPFParagraph) element);
        		iter.remove();
        	}
        }
    }
}