package Runners;

import cucumber.api.CucumberOptions;
import cucumber.api.junit.Cucumber;
import org.junit.runner.RunWith;

@RunWith(Cucumber.class)
@CucumberOptions(
        features = "src/test/resources/Features/Reports"
        ,glue = "Steps"
        ,tags = {"@MO-0001"}
)
public class TestRunner {
}
