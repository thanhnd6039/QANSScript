package StepDefinitions.Reports;

import Pages.Reports.RSGapToSFPartNumberDetailsPage;
import SharingTestContext.TestContext;
import cucumber.api.java.en.Then;

public class RSGapToSFPartNumberDetailsStepDefinition {
    private TestContext testContext;
    private RSGapToSFPartNumberDetailsPage rsGapToSFPartNumberDetailsPage;
    public RSGapToSFPartNumberDetailsStepDefinition(TestContext context){
        testContext = context;
        rsGapToSFPartNumberDetailsPage = testContext.getPageObjectManager().getRsGapToSFPartNumberDetailsPage();
    }
    @Then("^I should see the title of report is (.*)$")
    public void shouldSeeTitleOfReport(String title) throws Throwable{
        rsGapToSFPartNumberDetailsPage.shouldSeeTitle(title);
    }
    public void getSourceDataForGapToSFPNDetail(){

    }


}
