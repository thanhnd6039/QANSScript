package StepDefinitions.Reports;

import Pages.Reports.RSLoginPage;
import SharingTestContext.TestContext;
import cucumber.api.java.en.And;
import cucumber.api.java.en.Given;
import cucumber.api.java.en.When;

public class RSLoginStepDefinition {
    private TestContext testContext;
    private RSLoginPage rsLoginPage;

    public RSLoginStepDefinition(TestContext context){
        testContext = context;
        rsLoginPage = testContext.getPageObjectManager().getRsLoginPage();
    }

    @Given("^I login to the (.*) Report$")
    public void loginToReport(String title) throws Throwable{
        rsLoginPage.loginToReport(title);
    }


}
