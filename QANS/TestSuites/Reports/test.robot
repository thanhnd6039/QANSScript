*** Settings ***
Library     PandasLibrary

*** Test Cases ***
Testcase1
    ${csvPath}      Set Variable    C:\\RobotFramework\\Downloads\\data.csv
    Load Csv To DataFrame    ${csvPath}    df
    ${pivot_df}=    Pivot Table    df    index=Name    columns=Department    values=Salary    aggfunc=sum
    Log DataFrame    ${pivot_df}