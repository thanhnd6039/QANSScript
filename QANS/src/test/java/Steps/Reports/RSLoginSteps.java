package Steps.Reports;

import Pages.Reports.RSLoginPage;
import SharingTestContext.TestContext;
import cucumber.api.java.en.And;
import cucumber.api.java.en.Given;
import cucumber.api.java.en.Then;

public class RSLoginSteps {
    private TestContext testContext;
    private RSLoginPage rsLoginPage;
    public RSLoginSteps(TestContext context){
        testContext = context;
        rsLoginPage = testContext.getPageObjectManager().getRsLoginPage();
    }
    @Given("^I login to the (.*) report$")
    public void loginToReport(String title) throws Throwable{
        rsLoginPage.loginToReport(title);
    }

}
