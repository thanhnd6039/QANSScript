*** Settings ***
Resource    ../../Pages/Reports/MarginPage.robot

*** Test Cases ***
Validating Detailed Data For Margin Report
    Compare Data Between Margin Report And SS On NS     ${DOWNLOAD_DIR}\\Margin Reporting By OEM Part.xlsx      ${DOWNLOAD_DIR}\\RevenueCostDump.xlsx
