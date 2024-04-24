*** Settings ***
Resource    ../CommonPage.robot

*** Variables ***
${wowChangeReportResultFilePath}    C:\\RobotFramework\\Results\\WoWChangeReportResult.xlsx

*** Keywords ***
Write The Test Result Of WoW Change Report To Excel
    [Arguments]     ${item}     ${oemGroup}     ${valueOnWoWChangeReport}   ${valueOnSGWeeklyActionReport}
    File Should Exist    ${wowChangeReportResultFilePath}
    Open Excel Document    ${wowChangeReportResultFilePath}    doc_id=WoWChangeReportResult
    Switch Current Excel Document    doc_id=WoWChangeReportResult
    ${latestRowInWoWchangeReportResultFile}   Get Number Of Rows In Excel    ${wowChangeReportResultFilePath}
    ${nextRow}    Evaluate    ${latestRowInWoWchangeReportResultFile}+1
    Write Excel Cell    row_num=${nextRow}    col_num=1    value=${item}
    Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oemGroup}
    Write Excel Cell    row_num=${nextRow}    col_num=3    value=${valueOnWoWChangeReport}
    Write Excel Cell    row_num=${nextRow}    col_num=4    value=${valueOnSGWeeklyActionReport}
    Save Excel Document    ${wowChangeReportResultFilePath}
    Close Current Excel Document

Compare The Prev Quarter Ship Data For The Strategic Table Between WoW Change Report And SG Weekly Action DB Report
    [Arguments]     ${wowChangeReportFilePath}  ${sgWeeklyActionDBReportFilePath}
    ${result}   Set Variable    ${True}
    
    File Should Exist    ${wowChangeReportFilePath}
    Open Excel Document    ${wowChangeReportFilePath}    doc_id=WoWChangeReport
    
    File Should Exist    ${sgWeeklyActionDBReportFilePath}
    Open Excel Document    ${sgWeeklyActionDBReportFilePath}    doc_id=SGWeeklyActionDBReport

#   Verify the Pre Quarter Ship data
    FOR    ${rowIndexOnWoWChangeReport}    IN RANGE    2    6
        Switch Current Excel Document    doc_id=WoWChangeReport
        ${oemGroupColOnWoWChangeReport}      Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=1
        ${prevQShipColOnWoWChangeReport}     Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=2
        Switch Current Excel Document    doc_id=SGWeeklyActionDBReport
        ${numOfRowsOnSGWeeklyActionDBReport}    Get Number Of Rows In Excel    ${sgWeeklyActionDBReportFilePath}
        FOR    ${rowIndexOnSGWeeklyActionDBReport}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDBReport}+1
            ${oemGroupColOnSGWeeklyActionDBReport}      Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=1
            ${revColOnSGWeeklyActionDBReport}           Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=5
            IF    '${oemGroupColOnWoWChangeReport}' == '${oemGroupColOnSGWeeklyActionDBReport}'
                 IF    ${prevQShipColOnWoWChangeReport} != ${revColOnSGWeeklyActionDBReport}
                      ${result}     Set Variable    ${False}
                      Write The Test Result Of WoW Change Report To Excel    Prev Quarter Ship    ${oemGroupColOnWoWChangeReport}    ${prevQShipColOnWoWChangeReport}    ${revColOnSGWeeklyActionDBReport}
                 END
                 BREAK
            END
        END
    END
#   Verify the Total data
    Switch Current Excel Document    doc_id=SGWeeklyActionDBReport
    ${revTotalOnSGWeeklyActionDBReport}  Set Variable    0
    FOR    ${rowIndexOnSGWeeklyActionDBReport}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDBReport}+1
        ${mainSalesRepColOnSGWeeklyActionDBReport}      Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=2
        ${revColOnSGWeeklyActionDBReport}               Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=5
        IF    '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Cameron Sinclair' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Huan Tran' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Nicole Lau'
             ${revTotalOnSGWeeklyActionDBReport}     Evaluate    ${revTotalOnSGWeeklyActionDBReport}+${revColOnSGWeeklyActionDBReport}
        END
    END
    Switch Current Excel Document    doc_id=WoWChangeReport
    ${preQShipTotalOnWoWchangeReport}   Read Excel Cell    row_num=7    col_num=2
    IF    ${preQShipTotalOnWoWchangeReport} != ${revTotalOnSGWeeklyActionDBReport}
         ${result}     Set Variable    ${False}
         Write The Test Result Of WoW Change Report To Excel    Prev Quarter Ship    Strategic Total    ${preQShipTotalOnWoWchangeReport}    ${revTotalOnSGWeeklyActionDBReport}
    END
 #  Verify the Others data
    Switch Current Excel Document    doc_id=SGWeeklyActionDBReport
    ${revOthersOnSGWeeklyActionDBReport}  Set Variable    0
    FOR    ${rowIndexOnSGWeeklyActionDBReport}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDBReport}+1
        ${oemGroupColOnSGWeeklyActionDBReport}          Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=1
        ${mainSalesRepColOnSGWeeklyActionDBReport}      Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=2
        ${revColOnSGWeeklyActionDBReport}               Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=5
        IF    '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Cameron Sinclair' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Huan Tran' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Nicole Lau'
             IF    '${oemGroupColOnSGWeeklyActionDBReport}' != 'NOKIA/ALCATEL LUCENT WORLDWIDE' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'PALO ALTO NETWORKS' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'ARISTA' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'CIENA GROUP'
                  ${revOthersOnSGWeeklyActionDBReport}     Evaluate    ${revOthersOnSGWeeklyActionDBReport}+${revColOnSGWeeklyActionDBReport}
             END
        END
    END
    Switch Current Excel Document    doc_id=WoWChangeReport
    ${preQShipOthersOnWoWChangeReport}   Read Excel Cell    row_num=6    col_num=2
    IF    ${preQShipOthersOnWoWChangeReport} != ${revOthersOnSGWeeklyActionDBReport}
         ${result}     Set Variable    ${False}
         Write The Test Result Of WoW Change Report To Excel    Prev Quarter Ship    Strategic Others    ${preQShipOthersOnWoWChangeReport}    ${revOthersOnSGWeeklyActionDBReport}
    END

    IF    '${result}' == '${False}'
         Fail   The Pre Quarter Ship data between the WoW Change Report and SG Weekly Action Report is different
    END
    Close All Excel Documents

Compare The Prev Quarter Ship Data For The OEM East Table Between WoW Change Report And SG Weekly Action DB Report
    [Arguments]     ${wowChangeReportFilePath}  ${sgWeeklyActionDBReportFilePath}
    ${result}   Set Variable    ${True}

    File Should Exist    ${wowChangeReportFilePath}
    Open Excel Document    ${wowChangeReportFilePath}    doc_id=WoWChangeReport

    File Should Exist    ${sgWeeklyActionDBReportFilePath}
    Open Excel Document    ${sgWeeklyActionDBReportFilePath}    doc_id=SGWeeklyActionDBReport

#   Verify the Pre Quarter Ship data
    FOR    ${rowIndexOnWoWChangeReport}    IN RANGE    10    15
        Switch Current Excel Document    doc_id=WoWChangeReport
        ${oemGroupColOnWoWChangeReport}      Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=1
        ${prevQShipColOnWoWChangeReport}     Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=2
        Switch Current Excel Document    doc_id=SGWeeklyActionDBReport
        ${numOfRowsOnSGWeeklyActionDBReport}    Get Number Of Rows In Excel    ${sgWeeklyActionDBReportFilePath}
        FOR    ${rowIndexOnSGWeeklyActionDBReport}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDBReport}+1
            ${oemGroupColOnSGWeeklyActionDBReport}      Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=1
            ${revColOnSGWeeklyActionDBReport}           Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=5
            IF    '${oemGroupColOnWoWChangeReport}' == '${oemGroupColOnSGWeeklyActionDBReport}'
                 IF    ${prevQShipColOnWoWChangeReport} != ${revColOnSGWeeklyActionDBReport}
                      ${result}     Set Variable    ${False}
                      Write The Test Result Of WoW Change Report To Excel    Prev Quarter Ship    ${oemGroupColOnWoWChangeReport}    ${prevQShipColOnWoWChangeReport}    ${revColOnSGWeeklyActionDBReport}
                 END
                 BREAK
            END
        END
    END
#   Verify the Total data
    Switch Current Excel Document    doc_id=SGWeeklyActionDBReport
    ${revTotalOnSGWeeklyActionDBReport}  Set Variable    0
    FOR    ${rowIndexOnSGWeeklyActionDBReport}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDBReport}+1
        ${mainSalesRepColOnSGWeeklyActionDBReport}      Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=2
        ${revColOnSGWeeklyActionDBReport}               Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=5
        IF    '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Chris Seitz' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Daniel Schmidt' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Eli Tiomkin' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Michael Pauser'
             ${revTotalOnSGWeeklyActionDBReport}     Evaluate    ${revTotalOnSGWeeklyActionDBReport}+${revColOnSGWeeklyActionDBReport}
        END
    END
    Switch Current Excel Document    doc_id=WoWChangeReport
    ${preQShipTotalOnWoWchangeReport}   Read Excel Cell    row_num=16    col_num=2
    IF    ${preQShipTotalOnWoWchangeReport} != ${revTotalOnSGWeeklyActionDBReport}
         ${result}     Set Variable    ${False}
         Write The Test Result Of WoW Change Report To Excel    Prev Quarter Ship    OEM East Total    ${preQShipTotalOnWoWchangeReport}    ${revTotalOnSGWeeklyActionDBReport}
    END
 #  Verify the Others data
    Switch Current Excel Document    doc_id=SGWeeklyActionDBReport
    ${revOthersOnSGWeeklyActionDBReport}  Set Variable    0
    FOR    ${rowIndexOnSGWeeklyActionDBReport}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDBReport}+1
        ${oemGroupColOnSGWeeklyActionDBReport}          Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=1
        ${mainSalesRepColOnSGWeeklyActionDBReport}      Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=2
        ${revColOnSGWeeklyActionDBReport}               Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=5
        IF    '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Chris Seitz' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Daniel Schmidt' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Eli Tiomkin' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Michael Pauser'
             IF    '${oemGroupColOnSGWeeklyActionDBReport}' != 'ERICSSON WORLDWIDE' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'MELLANOX GROUP' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'CURTISS WRIGHT GROUP' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'JUNIPER NETWORKS' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'KONTRON NORTH AMERICA'
                  ${revOthersOnSGWeeklyActionDBReport}     Evaluate    ${revOthersOnSGWeeklyActionDBReport}+${revColOnSGWeeklyActionDBReport}
             END
        END
    END
    Switch Current Excel Document    doc_id=WoWChangeReport
    ${preQShipOthersOnWoWChangeReport}   Read Excel Cell    row_num=15    col_num=2
    IF    ${preQShipOthersOnWoWChangeReport} != ${revOthersOnSGWeeklyActionDBReport}
         ${result}     Set Variable    ${False}
         Write The Test Result Of WoW Change Report To Excel    Prev Quarter Ship    OEM East Others    ${preQShipOthersOnWoWChangeReport}    ${revOthersOnSGWeeklyActionDBReport}
    END

    IF    '${result}' == '${False}'
         Fail   The Pre Quarter Ship data between the WoW Change Report and SG Weekly Action Report is different
    END
    Close All Excel Documents

Compare The Prev Quarter Ship Data For The OEM West Table Between WoW Change Report And SG Weekly Action DB Report
    [Arguments]     ${wowChangeReportFilePath}  ${sgWeeklyActionDBReportFilePath}
    ${result}   Set Variable    ${True}

    File Should Exist    ${wowChangeReportFilePath}
    Open Excel Document    ${wowChangeReportFilePath}    doc_id=WoWChangeReport

    File Should Exist    ${sgWeeklyActionDBReportFilePath}
    Open Excel Document    ${sgWeeklyActionDBReportFilePath}    doc_id=SGWeeklyActionDBReport

#   Verify the Pre Quarter Ship data
    FOR    ${rowIndexOnWoWChangeReport}    IN RANGE    19    26
        Switch Current Excel Document    doc_id=WoWChangeReport
        ${oemGroupColOnWoWChangeReport}      Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=1
        ${prevQShipColOnWoWChangeReport}     Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=2
        Switch Current Excel Document    doc_id=SGWeeklyActionDBReport
        ${numOfRowsOnSGWeeklyActionDBReport}    Get Number Of Rows In Excel    ${sgWeeklyActionDBReportFilePath}
        FOR    ${rowIndexOnSGWeeklyActionDBReport}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDBReport}+1
            ${oemGroupColOnSGWeeklyActionDBReport}      Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=1
            ${revColOnSGWeeklyActionDBReport}           Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=5
            IF    '${oemGroupColOnWoWChangeReport}' == '${oemGroupColOnSGWeeklyActionDBReport}'
                 IF    ${prevQShipColOnWoWChangeReport} != ${revColOnSGWeeklyActionDBReport}
                      ${result}     Set Variable    ${False}
                      Write The Test Result Of WoW Change Report To Excel    Prev Quarter Ship    ${oemGroupColOnWoWChangeReport}    ${prevQShipColOnWoWChangeReport}    ${revColOnSGWeeklyActionDBReport}
                 END
                 BREAK
            END
        END
    END
#   Verify the Total data
    Switch Current Excel Document    doc_id=SGWeeklyActionDBReport
    ${revTotalOnSGWeeklyActionDBReport}  Set Variable    0
    FOR    ${rowIndexOnSGWeeklyActionDBReport}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDBReport}+1
        ${mainSalesRepColOnSGWeeklyActionDBReport}      Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=2
        ${revColOnSGWeeklyActionDBReport}               Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=5
        IF    '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Amy Duong' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Caden Douglas' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Michael Nilsson' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Tiger Wang' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Yoda Yasunobu'
             ${revTotalOnSGWeeklyActionDBReport}     Evaluate    ${revTotalOnSGWeeklyActionDBReport}+${revColOnSGWeeklyActionDBReport}
        END
    END
    Switch Current Excel Document    doc_id=WoWChangeReport
    ${preQShipTotalOnWoWchangeReport}   Read Excel Cell    row_num=27    col_num=2
    IF    ${preQShipTotalOnWoWchangeReport} != ${revTotalOnSGWeeklyActionDBReport}
         ${result}     Set Variable    ${False}
         Write The Test Result Of WoW Change Report To Excel    Prev Quarter Ship    OEM West Total    ${preQShipTotalOnWoWchangeReport}    ${revTotalOnSGWeeklyActionDBReport}
    END
 #  Verify the Others data
    Switch Current Excel Document    doc_id=SGWeeklyActionDBReport
    ${revOthersOnSGWeeklyActionDBReport}  Set Variable    0
    FOR    ${rowIndexOnSGWeeklyActionDBReport}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDBReport}+1
        ${oemGroupColOnSGWeeklyActionDBReport}          Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=1
        ${mainSalesRepColOnSGWeeklyActionDBReport}      Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=2
        ${revColOnSGWeeklyActionDBReport}               Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=5
        IF    '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Amy Duong' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Caden Douglas' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Michael Nilsson' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Tiger Wang' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Yoda Yasunobu'
             IF    '${oemGroupColOnSGWeeklyActionDBReport}' != 'SCHWEITZER ENGINEERING LABORATORIES (SEL)' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'PANASONIC AVIONICS' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'RADISYS CORPORATION' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'TEKTRONIX' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'TELEDYNE CONTROLS' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'NATIONAL INSTRUMENTS' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'ARROW ELECTRONICS, INC.'
                  ${revOthersOnSGWeeklyActionDBReport}     Evaluate    ${revOthersOnSGWeeklyActionDBReport}+${revColOnSGWeeklyActionDBReport}
             END
        END
    END
    Switch Current Excel Document    doc_id=WoWChangeReport
    ${preQShipOthersOnWoWChangeReport}   Read Excel Cell    row_num=26    col_num=2
    IF    ${preQShipOthersOnWoWChangeReport} != ${revOthersOnSGWeeklyActionDBReport}
         ${result}     Set Variable    ${False}
         Write The Test Result Of WoW Change Report To Excel    Prev Quarter Ship    OEM West Others    ${preQShipOthersOnWoWChangeReport}    ${revOthersOnSGWeeklyActionDBReport}
    END

    IF    '${result}' == '${False}'
         Fail   The Pre Quarter Ship data between the WoW Change Report and SG Weekly Action Report is different
    END
    Close All Excel Documents






Compare The Current Quarter Budget Data For The Strategic Table Between WoW Change Report And SG Weekly Action DB Report
    [Arguments]     ${wowChangeReportFilePath}  ${sgWeeklyActionDBReportFilePath}
    ${result}   Set Variable    ${True}

    File Should Exist    ${wowChangeReportFilePath}
    Open Excel Document    ${wowChangeReportFilePath}    doc_id=WoWChangeReport

    File Should Exist    ${sgWeeklyActionDBReportFilePath}
    Open Excel Document    ${sgWeeklyActionDBReportFilePath}    doc_id=SGWeeklyActionDBReport

#   Verify the Pre Quarter Ship data
    FOR    ${rowIndexOnWoWChangeReport}    IN RANGE    2    6
        Switch Current Excel Document    doc_id=WoWChangeReport
        ${oemGroupColOnWoWChangeReport}           Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=1
        ${currentQBudgetColOnWoWChangeReport}     Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=3
        Switch Current Excel Document    doc_id=SGWeeklyActionDBReport
        ${numOfRowsOnSGWeeklyActionDBReport}    Get Number Of Rows In Excel    ${sgWeeklyActionDBReportFilePath}
        FOR    ${rowIndexOnSGWeeklyActionDBReport}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDBReport}+1
            ${oemGroupColOnSGWeeklyActionDBReport}      Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=1
            ${budgetColOnSGWeeklyActionDBReport}           Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=3
            IF    '${oemGroupColOnWoWChangeReport}' == '${oemGroupColOnSGWeeklyActionDBReport}'
                 IF    ${currentQBudgetColOnWoWChangeReport} != ${budgetColOnSGWeeklyActionDBReport}
                      ${result}     Set Variable    ${False}
                      Write The Test Result Of WoW Change Report To Excel    Current Q Budget    ${oemGroupColOnWoWChangeReport}    ${currentQBudgetColOnWoWChangeReport}    ${budgetColOnSGWeeklyActionDBReport}
                 END
                 BREAK
            END
        END
    END
#   Verify the Total data
    Switch Current Excel Document    doc_id=SGWeeklyActionDBReport
    ${budgetTotalOnSGWeeklyActionDBReport}  Set Variable    0
    FOR    ${rowIndexOnSGWeeklyActionDBReport}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDBReport}+1
        ${mainSalesRepColOnSGWeeklyActionDBReport}         Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=2
        ${budgetColOnSGWeeklyActionDBReport}               Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=3
        IF    '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Cameron Sinclair' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Huan Tran' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Nicole Lau'
             ${budgetTotalOnSGWeeklyActionDBReport}     Evaluate    ${budgetTotalOnSGWeeklyActionDBReport}+${budgetColOnSGWeeklyActionDBReport}
        END
    END
    Switch Current Excel Document    doc_id=WoWChangeReport
    ${currentQBudgetTotalOnWoWchangeReport}   Read Excel Cell    row_num=7    col_num=3
    IF    ${currentQBudgetTotalOnWoWchangeReport} != ${budgetTotalOnSGWeeklyActionDBReport}
         ${result}     Set Variable    ${False}
         Write The Test Result Of WoW Change Report To Excel    Current Q Budget   Strategic Total    ${currentQBudgetTotalOnWoWchangeReport}    ${budgetTotalOnSGWeeklyActionDBReport}
    END
 #  Verify the Others data
    Switch Current Excel Document    doc_id=SGWeeklyActionDBReport
    ${budgetOthersOnSGWeeklyActionDBReport}  Set Variable    0
    FOR    ${rowIndexOnSGWeeklyActionDBReport}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDBReport}+1
        ${oemGroupColOnSGWeeklyActionDBReport}             Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=1
        ${mainSalesRepColOnSGWeeklyActionDBReport}         Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=2
        ${budgetColOnSGWeeklyActionDBReport}               Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=3
        IF    '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Cameron Sinclair' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Huan Tran' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Nicole Lau'
             IF    '${oemGroupColOnSGWeeklyActionDBReport}' != 'NOKIA/ALCATEL LUCENT WORLDWIDE' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'PALO ALTO NETWORKS' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'ARISTA' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'CIENA GROUP'
                  ${budgetOthersOnSGWeeklyActionDBReport}     Evaluate    ${budgetOthersOnSGWeeklyActionDBReport}+${budgetColOnSGWeeklyActionDBReport}
             END
        END
    END
    Switch Current Excel Document    doc_id=WoWChangeReport
    ${currentQBudgetOthersOnWoWChangeReport}   Read Excel Cell    row_num=6    col_num=3
    IF    ${currentQBudgetOthersOnWoWChangeReport} != ${budgetOthersOnSGWeeklyActionDBReport}
         ${result}     Set Variable    ${False}
         Write The Test Result Of WoW Change Report To Excel    Current Q Budget    Strategic Others    ${currentQBudgetOthersOnWoWChangeReport}    ${budgetOthersOnSGWeeklyActionDBReport}
    END

    IF    '${result}' == '${False}'
         Fail   The Current Q Budget data between the WoW Change Report and SG Weekly Action Report is different
    END
    Close All Excel Documents
