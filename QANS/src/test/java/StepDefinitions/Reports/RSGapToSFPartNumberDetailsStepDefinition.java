package StepDefinitions.Reports;

import Pages.Reports.RSGapToSFPartNumberDetailsPage;
import SharingTestContext.TestContext;
import cucumber.api.java.en.And;
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
    @And("^I get OEM Group by main Sale Rep from NS$")
    public void getOEMGroupByMainSaleRep() throws Throwable{
        rsGapToSFPartNumberDetailsPage.getAllOEMGroupByMainSaleRep();
    }
    @And("^I get the source data for the GAP to SF - Part Number Detail report from NS$")
    public void getSourceDataForGapToSFPNDetailReport(){
        rsGapToSFPartNumberDetailsPage.getSourceDataForGapToSFPNDetailReport();
    }


}
