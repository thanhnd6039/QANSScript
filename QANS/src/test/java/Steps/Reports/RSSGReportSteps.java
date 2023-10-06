package Steps.Reports;

import Pages.Reports.RSSGReportPage;
import SharingTestContext.TestContext;
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

}
