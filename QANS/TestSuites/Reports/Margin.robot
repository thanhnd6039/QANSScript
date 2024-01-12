*** Settings ***
Resource    ../../Pages/Reports/MarginPage.robot

*** Test Cases ***
Validating Detailed Data For Margin Report
    Compare Data Between Margin Report And SS On NS    reportFilePath=${DOWNLOAD_DIR}\\Margin Reporting By OEM Part_V2.xlsx    ssRevenueCostDumpFilePath=${DOWNLOAD_DIR}\\RevenueCostDump.xlsx
