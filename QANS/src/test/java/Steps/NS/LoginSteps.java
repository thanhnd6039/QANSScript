package Steps.NS;

import Pages.NS.LoginPage;
import SharingTestContext.TestContext;
import cucumber.api.java.en.And;
import cucumber.api.java.en.Given;

public class LoginSteps {
    private TestContext testContext;
    private LoginPage loginPage;
    public LoginSteps(TestContext context){
        testContext = context;
        loginPage = testContext.getPageObjectManager().getLoginPage();
    }
    @Given("^I navigate to NS$")
    public void navigateToNS() throws Throwable{
        loginPage.navigateToNS();
    }
    @And("^I login to NS$")
    public void loginToNS() throws Throwable{

    }

}
