package StepDefinitions.Reports;

import SharingTestContext.TestContext;
import cucumber.api.java.en.Given;
import cucumber.api.java.en.When;

public class RSLoginStepDefinition {
    private TestContext testContext;

    public RSLoginStepDefinition(TestContext context){
        testContext = context;
    }

    @Given("^I login to Report Viewer page$")
    public void loginToReportViewer() throws Throwable{

    }

}
