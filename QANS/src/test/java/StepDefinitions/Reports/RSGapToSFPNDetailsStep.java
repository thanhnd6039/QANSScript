package StepDefinitions.Reports;

import Pages.Reports.RSGapToSFPNDetailsPage;
import SharingTestContext.TestContext;
import cucumber.api.java.en.And;
import cucumber.api.java.en.Then;

public class RSGapToSFPNDetailsStep {
    private TestContext testContext;
    private RSGapToSFPNDetailsPage rsGapToSFPNDetailsPage;
    public RSGapToSFPNDetailsStep(TestContext context){
        testContext = context;
        rsGapToSFPNDetailsPage = testContext.getPageObjectManager().getRsGapToSFPNDetailsPage();
    }
    @Then("^I should see the title of report is (.*)$")
    public void shouldSeeTitleOfReport(String title) throws Throwable{
        rsGapToSFPNDetailsPage.shouldSeeTitle(title);
    }
    @And("^I get All OEM Group by Main Sales Rep from NS$")
    public void getAllOEMGroupByMainSaleRep() throws Throwable{
        rsGapToSFPNDetailsPage.getAllOEMGroupByMainSaleRep();
    }
    @And("^I get the source data for the GAP to SF - Part Number Detail report from NS$")
    public void getSourceDataForGapToSFPNDetailReport(){
        rsGapToSFPNDetailsPage.getSourceDataForGapToSFPNDetailReport();
    }


}
