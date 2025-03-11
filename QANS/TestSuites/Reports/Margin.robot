*** Settings ***
Resource    ../../Pages/Reports/MarginPage.robot

*** Test Cases ***
Testcase1
    Compare Data Between Margin Report And SS On NS    reportFilePath=${DOWNLOAD_DIR}\\Margin Reporting By OEM Part.xlsx    ssRevenueCostDumpFilePath=${DOWNLOAD_DIR}\\RevenueCostDump.xlsx



