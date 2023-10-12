package Steps.NS;

import Pages.NS.LoginPage;
import SharingTestContext.TestContext;
import cucumber.api.java.en.And;
import cucumber.api.java.en.Given;
import cucumber.api.java.en.Then;

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
        loginPage.loginToNS();
    }
    @And("^I choose Account Type is (PRODUCTION|SANDBOX)$")
    public void chooseAccountType(String expectedAccountType)throws Throwable{
        loginPage.chooseAccountType(expectedAccountType);
    }
    @Then("^I should see the title contains (.*) to on LogIn page$")
    public void shouldSeeTitle(String expectedTitle)throws Throwable{
        loginPage.shouldSeeTitle(expectedTitle);
    }
    @And("^I input verification code$")
    public void inputVerificationCode()throws Throwable{
        loginPage.inputVerificationCode();
    }
    @And("^I check to the Trust this device for 30 days for access to this role checkbox$")
    public void checktoCheckbox() throws Throwable {
        loginPage.checktoCheckbox();
    }
    @And("^I click to Submit button$")
    public void clickToButton() throws Throwable {

    }

}
