import java.io.FileInputStream;
import java.io.IOException;
import java.net.URI;
import java.util.Properties;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.atlassian.jira.rest.client.api.AuthenticationHandler;
import com.atlassian.jira.rest.client.api.IssueRestClient;
import com.atlassian.jira.rest.client.api.JiraRestClient;
import com.atlassian.jira.rest.client.api.domain.BasicIssue;
import com.atlassian.jira.rest.client.api.domain.Issue;
import com.atlassian.jira.rest.client.api.domain.IssueLink;
import com.atlassian.jira.rest.client.api.domain.input.IssueInput;
import com.atlassian.jira.rest.client.api.domain.input.IssueInputBuilder;
import com.atlassian.jira.rest.client.api.domain.input.LinkIssuesInput;
import com.atlassian.jira.rest.client.auth.BasicHttpAuthenticationHandler;
import com.atlassian.jira.rest.client.internal.async.AsynchronousJiraRestClientFactory;

public class JiraIssueCreator {
	
	public static IssueRestClient issueClient;
	public static JiraRestClient restClient;
	public static final String PROJECT_KEY = "TBD";
	public static final String LINK_TYPE = "Blocks";
	public static final String EXCEL_INPUT_FILENAME = "input_file.xlsx";

	public static void main(String[] args) throws Exception {

		Properties props = new Properties();

		props.load(ClassLoader.getSystemResourceAsStream("jira.properties"));

		URI jiraServerUri = URI.create(props.getProperty("url"));
		AsynchronousJiraRestClientFactory factory = new AsynchronousJiraRestClientFactory();
		AuthenticationHandler auth = new BasicHttpAuthenticationHandler(props.getProperty("username"), props.getProperty("password"));
		restClient = factory.create(jiraServerUri, auth);
		issueClient = restClient.getIssueClient();
		processInputFile(EXCEL_INPUT_FILENAME);
	}
	
	public static void linkIssue(String issue1Key, String issue2Key) {
		LinkIssuesInput input = new LinkIssuesInput(issue1Key, issue2Key, LINK_TYPE);
		issueClient.linkIssue(input).claim();
	}

	private static String createJiraTicket(String summary, String description){
		 IssueInputBuilder iib = new IssueInputBuilder();		 
		 iib.setProjectKey(PROJECT_KEY);
		 iib.setSummary(summary);
		 iib.setDescription(description);
		 iib.setIssueTypeId(10029L);
		 IssueInput issue = iib.build();
         BasicIssue issueObj = issueClient.createIssue(issue).claim();
         String issueKey = issueObj.getKey();
         System.out.println("Issue " + issueKey + " created successfully");
         return issueKey;
	}

	private static void processInputFile(String filename) throws IOException {

		FileInputStream fis = new FileInputStream(filename);

		Workbook wb = WorkbookFactory.create(fis);

		Sheet sheet = wb.getSheet("Sheet 1");

		String summary = "", description = "";

		int i = 0;
		for (Row row : sheet) {
			i++;
			if (i < 6) //skip first 6 records
				continue;
			Cell level1 = row.getCell(1);
			Cell level2 = row.getCell(2);
			if (level1 != null && !level1.getStringCellValue().isBlank()) {
				description = level1.getStringCellValue();
			}
			if (level2 != null && !level2.getStringCellValue().isBlank()) {
				summary = level2.getStringCellValue();
			}
			String linkedIssue = "";
			Cell linkedIssueCell = row.getCell(3);
			if(linkedIssueCell != null)
				linkedIssue = linkedIssueCell.getStringCellValue();
			
			String issueKey = createJiraTicket(summary, description);
			linkIssue(linkedIssue, issueKey);
			
			System.out.println("\""+summary +"\""+ ",\"" + description+"\","+linkedIssue+","+issueKey);
		}
	}

}
