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
	private XWPFParagraph name;
	private XWPFParagraph location;
	private int[] phone = new int[10];
	private XWPFParagraph phoneAndEmail;
	private XWPFParagraph linkedInURL;
	private ArrayList<String> headerSummaryList = new ArrayList<>();
	private String summary;
	private String title;
	private String jobTitle = "";
	private ArrayList<XWPFParagraph> highlightList = new ArrayList<>();
	private ArrayList<XWPFParagraph> coreCompetenciesList = new ArrayList<>();
	
	private ArrayList<XWPFParagraph> experienceList = new ArrayList<>();
	private ArrayList<XWPFParagraph> educationList = new ArrayList<>();
	private ArrayList<XWPFParagraph> certList = new ArrayList<>();
	private ArrayList<Credential> additionalCredList = new ArrayList<>();

	
	public String getName()
	{
		return WordUtils.capitalize(name.getText().trim().toLowerCase());
	}
	
	public boolean isNameBlack()
	{
		String color = name.getRuns().get(0).getColor();
		return "000000".equals(color) || null == color;
	}
	
	public String getLocation()
	{
		return location.getText().trim();
	}
	
	public boolean isLocationBlack()
	{
		String color = location.getRuns().get(0).getColor();
		return "000000".equals(color) || null == color;
	}
	
	public String getPhone()
	{
		return String.format("(%d%d%d) %d%d%d-%d%d%d%d", phone[0], phone[1], phone[2], 
								phone[3], phone[4], phone[5], phone[6], phone[7], phone[8], phone[9]); 
	}
	
	public boolean isPhoneBlack()
	{
		String color = phoneAndEmail.getRuns().get(0).getColor();
		return "000000".equals(color) || null == color;
	}
	
	public String getEmail()
	{
		return phoneAndEmail.getText().split("\\s+")[1].trim();
	}
	
	public boolean isEmailBlack()
	{
		String color = phoneAndEmail.getRuns().get(phoneAndEmail.getRuns().size()-1).getColor();
		return "000000".equals(color) || null == color;
	}
	
	public String getlinkedInURL()
	{
		return linkedInURL.getText().trim();
	}
	
	public boolean islinkedInURLBlack()
	{
		String color = linkedInURL.getRuns().get(0).getColor();
		return "000000".equals(color) || null == color;
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
		
		// Job title can be distinguished by assuming it is the last few words prior to the word 'with', excluding
		// the first word after a comma. Good assumption?
		String[] commaList = title.split(",");
		String[] wordsAfterLastComma = commaList[commaList.length-1].split("\\s+");
		if(wordsAfterLastComma.length == 1)
		{
			jobTitle = wordsAfterLastComma[0];
		}
		else
		{
			// properly formatted will be 1 empty string (the space after the comma)
			// and one more descriptive term
			for(int i=2; i < wordsAfterLastComma.length; i++)
			{
				jobTitle = jobTitle + wordsAfterLastComma[i] + " ";
			}
			jobTitle = jobTitle.trim();
		}
	}
	
	public String getTitle()
	{
		return title;
	}
	
	public String getJobTitle()
	{
		return jobTitle;
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
		this.name = tempParaList.get(0);
		this.location = tempParaList.get(1);
		
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
		this.phoneAndEmail = tempParaList.get(2);
		this.linkedInURL = tempParaList.get(3);
	}
}
