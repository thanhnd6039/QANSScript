package Steps.Reports;

import Pages.Reports.RSSGReportPage;
import SharingTestContext.TestContext;
import cucumber.api.java.en.And;
import cucumber.api.java.en.Given;
import cucumber.api.java.en.Then;

public class RSSGReportSteps {
    private TestContext testContext;
    private RSSGReportPage rssgReportPage;

    public RSSGReportSteps(TestContext context){
        testContext = context;
        rssgReportPage = testContext.getPageObjectManager().getRssgReportPage();
    }
    @Then("^I should see the title of report is (.*)$")
    public void shouldSeeTitleOfReport(String expectedTitle) throws Throwable {
        rssgReportPage.shouldSeeTitleOfReport(expectedTitle);
    }
    @And("^I get data from the SG report for quarter (.*) year (.*)$")
    public void getDataFromSGReport(String quarter, String year)throws Throwable{
        rssgReportPage.getDataFromSGReport(quarter, year);
    }
    @And("^I get data from SS Revenue Cost Dump on NS for quarter (.*) year (.*)$")
    public void getDataFromSSRevCostDump(String quarter, String year)throws Throwable{
        rssgReportPage.getDataFromSSRevCostDump(quarter, year);
    }

}
