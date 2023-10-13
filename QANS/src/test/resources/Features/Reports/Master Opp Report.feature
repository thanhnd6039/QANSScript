Feature: The Master Opp Report
  Description:

  @MO-0001
  Scenario: Verify data for the Master Opp report by default
    Given I navigate to NS
    And I login to NS
    And I choose Account Type is PRODUCTION
    Then I should see the title contains Logging in to on LogIn page
    And I input verification code
    And I check to the Trust this device for 30 days for access to this role checkbox
    And I click to Submit button
#    Then I should see the account type is PRODUCTION
#    And I choose Role is VT Full Developer
#    Then I should see the user role is VT Full Developer
#    And I navigate to ss Master Opp on NS
#    And I export excel from ss Master Opp on NS
#    And I open new tab
#    And I navigate to the Master Opp report
#    Then I should see the title of Master Opp report is abc
