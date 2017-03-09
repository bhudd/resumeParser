package resumeParser;

import java.util.ArrayList;
import java.util.List;

import org.apache.commons.lang3.text.WordUtils;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;

public class MasterResume {
	public static enum MasterResumeLocations
	{
		TABLE_MAIN_HEADER,
		TABLE_HEADER_SUMMARY,
		PARA_SUMMARY,
		TABLE_HIGHLIGHTS,
		TABLE_CORE_COMPETENCIES,
	}
	
	public static enum CredentialTypes
	{
		TECHNICAL_SKILLS ("Technical Skills"),
		EDUCATION ("Education"),
		LANGUAGES ("Languages"),
		HONORS_AWARDS ("Honors & Awards"),
		PROFESSIONAL_DEVELOPMENT ("Professional Development"),
		ORGANIZATIONS ("Organizations"),
		VOLUNTEERING_EXPERIENCE ("Volunteering Experience"),
		INTERESTS ("Interests");
		
		final String string;
		CredentialTypes(String string) {this.string = string;}
		@Override public String toString(){return this.string;}
	}
	
	public static class Credential
	{
		CredentialTypes type;
		ArrayList<XWPFParagraph> credList = new ArrayList<>();
		boolean isBlack = false;
	}
	
	private XWPFTable PersonalInfoTable;
	private String name;
	private String location;
	private int[] phone = new int[10];
	private String email;
	private String linkedInURL;
	private ArrayList<String> headerSummaryList = new ArrayList<>();
	private String summary;
	private String title;
	private ArrayList<XWPFParagraph> highlightList = new ArrayList<>();
	private ArrayList<XWPFParagraph> coreCompetenciesList = new ArrayList<>();
	
	private ArrayList<XWPFParagraph> experienceList = new ArrayList<>();
	private ArrayList<XWPFParagraph> educationList = new ArrayList<>();
	private ArrayList<XWPFParagraph> certList = new ArrayList<>();
	private ArrayList<Credential> additionalCredList = new ArrayList<>();

	
	public String getName()
	{
		return name;
	}
	
	public String getLocation()
	{
		return location;
	}
	
	public String getPhone()
	{
		return String.format("(%d%d%d) %d%d%d-%d%d%d%d", phone[0], phone[1], phone[2], 
								phone[3], phone[4], phone[5], phone[6], phone[7], phone[8], phone[9]); 
	}
	
	public String getEmail()
	{
		return email;
	}
	
	public String getlinkedInURL()
	{
		return linkedInURL;
	}
	
	public ArrayList<String> getHeaderSummaryList()
	{
		return headerSummaryList;
	}
	
	public String getSummary()
	{
		return summary;
	}
	
	public void setSummary(String summary)
	{
		this.summary = summary;
		
		if(summary.startsWith("An"))
		{
			title = summary.substring(2, summary.indexOf("with")-1);
		}
		// starts with a
		else
		{
			title = summary.substring(1, summary.indexOf("with")-1);
		}
		
		title = WordUtils.capitalize(title);
	}
	
	public String getTitle()
	{
		return title;
	}
	
	public ArrayList<XWPFParagraph> getHighlights()
	{
		return highlightList;
	}
	
	public ArrayList<XWPFParagraph> getCoreCompetencies()
	{
		return coreCompetenciesList;
	}
	
	public ArrayList<XWPFParagraph> getExperiences()
	{
		return experienceList;
	}
	
	public ArrayList<XWPFParagraph> getEducationList()
	{
		return educationList;
	}
	
	public ArrayList<XWPFParagraph> getCertificationsList()
	{
		return certList;
	}
	
	public ArrayList<Credential> getCredentials()
	{
		return additionalCredList;
	}
	
	public XWPFTable getPersonalTable()
	{
		return PersonalInfoTable;
	}
	
	public void setPersonalInfo(XWPFTable table)
	{
		PersonalInfoTable = table;
		List<XWPFParagraph> tempParaList = PersonalInfoTable.getRow(0).getCell(0).getParagraphs();
		// capitalize each word in name
		this.name = WordUtils.capitalize(tempParaList.get(0).getText().trim().toLowerCase());
		
		this.location = tempParaList.get(1).getText().trim();
		String phoneString = tempParaList.get(2).getText().split("\\s+")[0].trim();
		
		int i=0;
		for(char c : phoneString.toCharArray())
		{
			if(Character.isDigit(c))
			{
				phone[i++] = Character.getNumericValue(c);
				if(i > 9) {break;}
			}
		}
		this.email = tempParaList.get(2).getText().split("\\s+")[1].trim();
		this.linkedInURL = tempParaList.get(3).getText().trim();
	}
}
