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

Compare Data For The Strategic Table Between WoW Change Report And SG Weekly Action DB Report
    [Arguments]     ${wowChangeReportFilePath}  ${sgWeeklyActionDBReportFilePath}   ${posOfColOnWoWChangeReport}    ${posOfColOnSGWeeklyActionDBReport}     ${nameOfCol}
    ${result}   Set Variable    ${True}
    
    File Should Exist    ${wowChangeReportFilePath}
    Open Excel Document    ${wowChangeReportFilePath}    doc_id=WoWChangeReport
    
    File Should Exist    ${sgWeeklyActionDBReportFilePath}
    Open Excel Document    ${sgWeeklyActionDBReportFilePath}    doc_id=SGWeeklyActionDBReport

#   Verify the data for each OEM Group
    FOR    ${rowIndexOnWoWChangeReport}    IN RANGE    2    6
        Switch Current Excel Document    doc_id=WoWChangeReport
        ${oemGroupColOnWoWChangeReport}      Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=1
        ${dataColOnWoWChangeReport}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=${posOfColOnWoWChangeReport}
        Switch Current Excel Document    doc_id=SGWeeklyActionDBReport
        ${numOfRowsOnSGWeeklyActionDBReport}    Get Number Of Rows In Excel    ${sgWeeklyActionDBReportFilePath}
        FOR    ${rowIndexOnSGWeeklyActionDBReport}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDBReport}+1
            ${oemGroupColOnSGWeeklyActionDBReport}      Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=1
            ${dataColOnSGWeeklyActionDBReport}           Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=${posOfColOnSGWeeklyActionDBReport}
            ${dataColOnSGWeeklyActionDBReport}   Evaluate  "%.2f" % ${dataColOnSGWeeklyActionDBReport}
            IF    '${oemGroupColOnWoWChangeReport}' == '${oemGroupColOnSGWeeklyActionDBReport}'
                 IF    ${dataColOnWoWChangeReport} != ${dataColOnSGWeeklyActionDBReport}
                      ${result}     Set Variable    ${False}
                      Write The Test Result Of WoW Change Report To Excel    ${nameOfCol} for Strategic table    ${oemGroupColOnWoWChangeReport}    ${dataColOnWoWChangeReport}    ${dataColOnSGWeeklyActionDBReport}
                 END
                 BREAK
            END
        END
    END
#   Verify the Total data
    Switch Current Excel Document    doc_id=SGWeeklyActionDBReport
    ${dataTotalOnSGWeeklyActionDBReport}  Set Variable    0
    FOR    ${rowIndexOnSGWeeklyActionDBReport}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDBReport}+1
        ${mainSalesRepColOnSGWeeklyActionDBReport}      Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=2
        ${dataColOnSGWeeklyActionDBReport}               Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=${posOfColOnSGWeeklyActionDBReport}
        IF    '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Cameron Sinclair' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Huan Tran' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Nicole Lau'
             ${dataTotalOnSGWeeklyActionDBReport}     Evaluate    ${dataTotalOnSGWeeklyActionDBReport}+${dataColOnSGWeeklyActionDBReport}
        END
    END
    ${dataTotalOnSGWeeklyActionDBReport}   Evaluate  "%.2f" % ${dataTotalOnSGWeeklyActionDBReport}
    Switch Current Excel Document    doc_id=WoWChangeReport
    ${dataTotalOnWoWchangeReport}   Read Excel Cell    row_num=7    col_num=${posOfColOnWoWChangeReport}
    IF    ${dataTotalOnWoWchangeReport} != ${dataTotalOnSGWeeklyActionDBReport}
         ${result}     Set Variable    ${False}
         Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    Strategic Total    ${dataTotalOnWoWchangeReport}    ${dataTotalOnSGWeeklyActionDBReport}
    END
 #  Verify the Others data
    Switch Current Excel Document    doc_id=SGWeeklyActionDBReport
    ${dataOthersOnSGWeeklyActionDBReport}  Set Variable    0
    FOR    ${rowIndexOnSGWeeklyActionDBReport}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDBReport}+1
        ${oemGroupColOnSGWeeklyActionDBReport}          Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=1
        ${mainSalesRepColOnSGWeeklyActionDBReport}      Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=2
        ${dataColOnSGWeeklyActionDBReport}               Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=${posOfColOnSGWeeklyActionDBReport}
        IF    '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Cameron Sinclair' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Huan Tran' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Nicole Lau'
             IF    '${oemGroupColOnSGWeeklyActionDBReport}' != 'NOKIA/ALCATEL LUCENT WORLDWIDE' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'PALO ALTO NETWORKS' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'ARISTA' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'CIENA GROUP'
                  ${dataOthersOnSGWeeklyActionDBReport}     Evaluate    ${dataOthersOnSGWeeklyActionDBReport}+${dataColOnSGWeeklyActionDBReport}
             END
        END
    END
    ${dataOthersOnSGWeeklyActionDBReport}   Evaluate  "%.2f" % ${dataOthersOnSGWeeklyActionDBReport}
    Switch Current Excel Document    doc_id=WoWChangeReport
    ${dataOthersOnWoWChangeReport}   Read Excel Cell    row_num=6    col_num=${posOfColOnWoWChangeReport}
    IF    ${dataOthersOnWoWChangeReport} != ${dataOthersOnSGWeeklyActionDBReport}
         ${result}     Set Variable    ${False}
         Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    Strategic Others    ${dataOthersOnWoWChangeReport}    ${dataOthersOnSGWeeklyActionDBReport}
    END

    IF    '${result}' == '${False}'
         Close All Excel Documents
         Fail   The ${nameOfCol} data between the WoW Change Report and SG Weekly Action Report is different
    END
    Close All Excel Documents

Compare Data For The OEM East Table Between WoW Change Report And SG Weekly Action DB Report
    [Arguments]     ${wowChangeReportFilePath}  ${sgWeeklyActionDBReportFilePath}   ${posOfColOnWoWChangeReport}    ${posOfColOnSGWeeklyActionDBReport}     ${nameOfCol}
    ${result}   Set Variable    ${True}

    File Should Exist    ${wowChangeReportFilePath}
    Open Excel Document    ${wowChangeReportFilePath}    doc_id=WoWChangeReport

    File Should Exist    ${sgWeeklyActionDBReportFilePath}
    Open Excel Document    ${sgWeeklyActionDBReportFilePath}    doc_id=SGWeeklyActionDBReport

#   Verify the data for each OEM Group
    FOR    ${rowIndexOnWoWChangeReport}    IN RANGE    10    15
        Switch Current Excel Document    doc_id=WoWChangeReport
        ${oemGroupColOnWoWChangeReport}      Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=1
        ${dataColOnWoWChangeReport}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=${posOfColOnWoWChangeReport}
        Switch Current Excel Document    doc_id=SGWeeklyActionDBReport
        ${numOfRowsOnSGWeeklyActionDBReport}    Get Number Of Rows In Excel    ${sgWeeklyActionDBReportFilePath}
        FOR    ${rowIndexOnSGWeeklyActionDBReport}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDBReport}+1
            ${oemGroupColOnSGWeeklyActionDBReport}       Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=1
            ${dataColOnSGWeeklyActionDBReport}           Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=${posOfColOnSGWeeklyActionDBReport}
            ${dataColOnSGWeeklyActionDBReport}   Evaluate  "%.2f" % ${dataColOnSGWeeklyActionDBReport}
            IF    '${oemGroupColOnWoWChangeReport}' == '${oemGroupColOnSGWeeklyActionDBReport}'
                 IF    ${dataColOnWoWChangeReport} != ${dataColOnSGWeeklyActionDBReport}
                      ${result}     Set Variable    ${False}
                      Write The Test Result Of WoW Change Report To Excel    ${nameOfCol} for OEM East table    ${oemGroupColOnWoWChangeReport}    ${dataColOnWoWChangeReport}    ${dataColOnSGWeeklyActionDBReport}
                 END
                 BREAK
            END
        END
    END
#   Verify the Total data
    Switch Current Excel Document    doc_id=SGWeeklyActionDBReport
    ${dataTotalOnSGWeeklyActionDBReport}  Set Variable    0
    FOR    ${rowIndexOnSGWeeklyActionDBReport}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDBReport}+1
        ${mainSalesRepColOnSGWeeklyActionDBReport}      Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=2
        ${dataColOnSGWeeklyActionDBReport}               Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=${posOfColOnSGWeeklyActionDBReport}
        IF    '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Chris Seitz' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Daniel Schmidt' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Eli Tiomkin' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Michael Pauser'
             ${dataTotalOnSGWeeklyActionDBReport}     Evaluate    ${dataTotalOnSGWeeklyActionDBReport}+${dataColOnSGWeeklyActionDBReport}
        END
    END
    ${dataTotalOnSGWeeklyActionDBReport}   Evaluate  "%.2f" % ${dataTotalOnSGWeeklyActionDBReport}
    Switch Current Excel Document    doc_id=WoWChangeReport
    ${dataTotalOnWoWchangeReport}   Read Excel Cell    row_num=16    col_num=${posOfColOnWoWChangeReport}
    IF    ${dataTotalOnWoWchangeReport} != ${dataTotalOnSGWeeklyActionDBReport}
         ${result}     Set Variable    ${False}
         Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM East Total    ${dataTotalOnWoWchangeReport}    ${dataTotalOnSGWeeklyActionDBReport}
    END
 #  Verify the Others data
    Switch Current Excel Document    doc_id=SGWeeklyActionDBReport
    ${dataOthersOnSGWeeklyActionDBReport}  Set Variable    0
    FOR    ${rowIndexOnSGWeeklyActionDBReport}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDBReport}+1
        ${oemGroupColOnSGWeeklyActionDBReport}          Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=1
        ${mainSalesRepColOnSGWeeklyActionDBReport}      Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=2
        ${dataColOnSGWeeklyActionDBReport}               Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=${posOfColOnSGWeeklyActionDBReport}
        IF    '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Chris Seitz' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Daniel Schmidt' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Eli Tiomkin' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Michael Pauser'
             IF    '${oemGroupColOnSGWeeklyActionDBReport}' != 'ERICSSON WORLDWIDE' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'MELLANOX GROUP' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'CURTISS WRIGHT GROUP' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'JUNIPER NETWORKS' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'KONTRON NORTH AMERICA'
                  ${dataOthersOnSGWeeklyActionDBReport}     Evaluate    ${dataOthersOnSGWeeklyActionDBReport}+${dataColOnSGWeeklyActionDBReport}
             END
        END
    END
    ${dataOthersOnSGWeeklyActionDBReport}   Evaluate  "%.2f" % ${dataOthersOnSGWeeklyActionDBReport}
    Switch Current Excel Document    doc_id=WoWChangeReport
    ${dataOthersOnWoWChangeReport}   Read Excel Cell    row_num=15    col_num=${posOfColOnWoWChangeReport}
    IF    ${dataOthersOnWoWChangeReport} != ${dataOthersOnSGWeeklyActionDBReport}
         ${result}     Set Variable    ${False}
         Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM East Others    ${dataOthersOnWoWChangeReport}    ${dataOthersOnSGWeeklyActionDBReport}
    END

    IF    '${result}' == '${False}'
         Close All Excel Documents
         Fail   The ${nameOfCol} data for the OEM East table between the WoW Change Report and SG Weekly Action Report is different
    END
    Close All Excel Documents

Compare Data For The OEM West Table Between WoW Change Report And SG Weekly Action DB Report
    [Arguments]     ${wowChangeReportFilePath}  ${sgWeeklyActionDBReportFilePath}   ${posOfColOnWoWChangeReport}    ${posOfColOnSGWeeklyActionDBReport}     ${nameOfCol}
    ${result}   Set Variable    ${True}

    File Should Exist    ${wowChangeReportFilePath}
    Open Excel Document    ${wowChangeReportFilePath}    doc_id=WoWChangeReport

    File Should Exist    ${sgWeeklyActionDBReportFilePath}
    Open Excel Document    ${sgWeeklyActionDBReportFilePath}    doc_id=SGWeeklyActionDBReport

#   Verify the data for each OEM Group
    FOR    ${rowIndexOnWoWChangeReport}    IN RANGE    19    26
        Switch Current Excel Document    doc_id=WoWChangeReport
        ${oemGroupColOnWoWChangeReport}      Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=1
        ${dataColOnWoWChangeReport}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=${posOfColOnWoWChangeReport}
        Switch Current Excel Document    doc_id=SGWeeklyActionDBReport
        ${numOfRowsOnSGWeeklyActionDBReport}    Get Number Of Rows In Excel    ${sgWeeklyActionDBReportFilePath}
        FOR    ${rowIndexOnSGWeeklyActionDBReport}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDBReport}+1
            ${oemGroupColOnSGWeeklyActionDBReport}       Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=1
            ${dataColOnSGWeeklyActionDBReport}           Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=${posOfColOnSGWeeklyActionDBReport}
            ${dataColOnSGWeeklyActionDBReport}   Evaluate  "%.2f" % ${dataColOnSGWeeklyActionDBReport}
            IF    '${oemGroupColOnWoWChangeReport}' == '${oemGroupColOnSGWeeklyActionDBReport}'
                 IF    ${dataColOnWoWChangeReport} != ${dataColOnSGWeeklyActionDBReport}
                      ${result}     Set Variable    ${False}
                      Write The Test Result Of WoW Change Report To Excel    ${nameOfCol} for OEM West Table    ${oemGroupColOnWoWChangeReport}    ${dataColOnWoWChangeReport}    ${dataColOnSGWeeklyActionDBReport}
                 END
                 BREAK
            END
        END
    END
#   Verify the Total data
    Switch Current Excel Document    doc_id=SGWeeklyActionDBReport
    ${dataTotalOnSGWeeklyActionDBReport}  Set Variable    0
    FOR    ${rowIndexOnSGWeeklyActionDBReport}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDBReport}+1
        ${mainSalesRepColOnSGWeeklyActionDBReport}      Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=2
        ${dataColOnSGWeeklyActionDBReport}               Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=${posOfColOnSGWeeklyActionDBReport}
        IF    '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Amy Duong' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Caden Douglas' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Michael Nilsson' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Tiger Wang' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Yoda Yasunobu'
             ${dataTotalOnSGWeeklyActionDBReport}     Evaluate    ${dataTotalOnSGWeeklyActionDBReport}+${dataColOnSGWeeklyActionDBReport}
        END
    END
    ${dataTotalOnSGWeeklyActionDBReport}   Evaluate  "%.2f" % ${dataTotalOnSGWeeklyActionDBReport}
    Switch Current Excel Document    doc_id=WoWChangeReport
    ${dataTotalOnWoWchangeReport}   Read Excel Cell    row_num=27    col_num=${posOfColOnWoWChangeReport}
    IF    ${dataTotalOnWoWchangeReport} != ${dataTotalOnSGWeeklyActionDBReport}
         ${result}     Set Variable    ${False}
         Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM West Total    ${dataTotalOnWoWchangeReport}    ${dataTotalOnSGWeeklyActionDBReport}
    END
 #  Verify the Others data
    Switch Current Excel Document    doc_id=SGWeeklyActionDBReport
    ${dataOthersOnSGWeeklyActionDBReport}  Set Variable    0
    FOR    ${rowIndexOnSGWeeklyActionDBReport}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDBReport}+1
        ${oemGroupColOnSGWeeklyActionDBReport}          Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=1
        ${mainSalesRepColOnSGWeeklyActionDBReport}      Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=2
        ${dataColOnSGWeeklyActionDBReport}               Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=${posOfColOnSGWeeklyActionDBReport}
        IF    '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Amy Duong' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Caden Douglas' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Michael Nilsson' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Tiger Wang' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Yoda Yasunobu'
             IF    '${oemGroupColOnSGWeeklyActionDBReport}' != 'SCHWEITZER ENGINEERING LABORATORIES (SEL)' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'PANASONIC AVIONICS' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'RADISYS CORPORATION' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'TEKTRONIX' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'TELEDYNE CONTROLS' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'NATIONAL INSTRUMENTS' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'ARROW ELECTRONICS, INC.'
                  ${dataOthersOnSGWeeklyActionDBReport}     Evaluate    ${dataOthersOnSGWeeklyActionDBReport}+${dataColOnSGWeeklyActionDBReport}
             END
        END
    END
    ${dataOthersOnSGWeeklyActionDBReport}   Evaluate  "%.2f" % ${dataOthersOnSGWeeklyActionDBReport}
    Switch Current Excel Document    doc_id=WoWChangeReport
    ${dataOthersOnWoWChangeReport}   Read Excel Cell    row_num=26    col_num=${posOfColOnWoWChangeReport}
    IF    ${dataOthersOnWoWChangeReport} != ${dataOthersOnSGWeeklyActionDBReport}
         ${result}     Set Variable    ${False}
         Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM West Others    ${dataOthersOnWoWChangeReport}    ${dataOthersOnSGWeeklyActionDBReport}
    END

    IF    '${result}' == '${False}'
         Close All Excel Documents
         Fail   The ${nameOfCol} data for the OEM West table between the WoW Change Report and SG Weekly Action Report is different
    END
    Close All Excel Documents

Compare The LW Commit Or Comment Data Between WoW Change Report And WoW Change Report On VDC
    [Arguments]     ${wowChangeReportFilePath}  ${wowChangeReportOnVDCFilePath}   ${posOfColOnWoWChangeReport}    ${nameOfCol}   ${table}
    ${result}   Set Variable    ${True}

    File Should Exist    ${wowChangeReportFilePath}
    Open Excel Document    ${wowChangeReportFilePath}    doc_id=WoWChangeReport

    File Should Exist    ${wowChangeReportOnVDCFilePath}
    Open Excel Document    ${wowChangeReportOnVDCFilePath}    doc_id=WoWChangeReportOnVDC

    IF    '${table}' == 'Strategic'
         FOR    ${rowIndexOnWoWChangeReport}    IN RANGE    2    8
            Switch Current Excel Document    doc_id=WoWChangeReport
            ${oemGroupColOnWoWChangeReport}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=1
            ${dataColOnWoWChangeReport}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=${posOfColOnWoWChangeReport}
            Switch Current Excel Document    doc_id=WoWChangeReportOnVDC
            FOR    ${rowIndexOnWoWChangeReportOnVDC}    IN RANGE    2   8
                ${oemGroupColOnWoWChangeReportOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReportOnVDC}    col_num=1
                IF    '${nameOfCol}' == 'LW Commit'
                     ${posOfColOnWoWChangeReportOnVDC}   Evaluate    ${posOfColOnWoWChangeReport}+1
                     ${dataColOnWoWChangeReportOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReportOnVDC}    col_num=${posOfColOnWoWChangeReportOnVDC}
                ELSE IF  '${nameOfCol}' == 'Comments'
                     ${dataColOnWoWChangeReportOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReportOnVDC}    col_num=${posOfColOnWoWChangeReport}
                ELSE
                   Close All Excel Documents
                   Fail  The name of column ${nameOfCol} is not valid
                END

                IF    '${oemGroupColOnWoWChangeReport}' == '${oemGroupColOnWoWChangeReportOnVDC}'
                     IF    '${dataColOnWoWChangeReport}' == 'None'
                          ${dataColOnWoWChangeReport}  Set Variable    ${EMPTY}
                     END
                     IF    '${dataColOnWoWChangeReportOnVDC}' == 'None'
                          ${dataColOnWoWChangeReportOnVDC}  Set Variable    ${EMPTY}
                     END

                     IF   '${dataColOnWoWChangeReport}' != '${dataColOnWoWChangeReportOnVDC}'
                         Log To Console    dataColOnWoWChangeReport: ${dataColOnWoWChangeReport}
                         Log To Console    dataColOnWoWChangeReportOnVDC: ${dataColOnWoWChangeReportOnVDC}
                         ${result}  Set Variable    ${False}
                         IF    '${oemGroupColOnWoWChangeReport}' == 'OTHERS'
                              Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    Strategic Others    ${dataColOnWoWChangeReport}    ${dataColOnWoWChangeReportOnVDC}
                         ELSE IF  '${oemGroupColOnWoWChangeReport}' == 'Total'
                              Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    Strategic Total    ${dataColOnWoWChangeReport}    ${dataColOnWoWChangeReportOnVDC}
                         ELSE
                              Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    ${oemGroupColOnWoWChangeReport}    ${dataColOnWoWChangeReport}    ${dataColOnWoWChangeReportOnVDC}
                         END
                     END
                     BREAK
                END
            END
         END
    END

    IF    '${table}' == 'OEM East'
        FOR    ${rowIndexOnWoWChangeReport}    IN RANGE    10    17
            Switch Current Excel Document    doc_id=WoWChangeReport
            ${oemGroupColOnWoWChangeReport}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=1
            ${dataColOnWoWChangeReport}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=${posOfColOnWoWChangeReport}
            Switch Current Excel Document    doc_id=WoWChangeReportOnVDC
            FOR    ${rowIndexOnWoWChangeReportOnVDC}    IN RANGE    10   17
                ${oemGroupColOnWoWChangeReportOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReportOnVDC}    col_num=1
                IF    '${nameOfCol}' == 'LW Commit'
                     ${posOfColOnWoWChangeReportOnVDC}   Evaluate    ${posOfColOnWoWChangeReport}+1
                     ${dataColOnWoWChangeReportOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReportOnVDC}    col_num=${posOfColOnWoWChangeReportOnVDC}
                ELSE IF  '${nameOfCol}' == 'Comments'
                     ${dataColOnWoWChangeReportOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReportOnVDC}    col_num=${posOfColOnWoWChangeReport}
                ELSE
                   Close All Excel Documents
                   Fail  The name of column ${nameOfCol} is not valid
                END

                IF    '${oemGroupColOnWoWChangeReport}' == '${oemGroupColOnWoWChangeReportOnVDC}'
                     IF    '${dataColOnWoWChangeReport}' == 'None'
                          ${dataColOnWoWChangeReport}  Set Variable    ${EMPTY}
                     END
                     IF    '${dataColOnWoWChangeReportOnVDC}' == 'None'
                          ${dataColOnWoWChangeReportOnVDC}  Set Variable    ${EMPTY}
                     END

                     IF   '${dataColOnWoWChangeReport}' != '${dataColOnWoWChangeReportOnVDC}'

                         ${result}  Set Variable    ${False}
                         IF    '${oemGroupColOnWoWChangeReport}' == 'OTHERS'
                              Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM East Others    ${dataColOnWoWChangeReport}    ${dataColOnWoWChangeReportOnVDC}
                         ELSE IF  '${oemGroupColOnWoWChangeReport}' == 'Total'
                              Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM East Total    ${dataColOnWoWChangeReport}    ${dataColOnWoWChangeReportOnVDC}
                         ELSE
                              Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    ${oemGroupColOnWoWChangeReport}    ${dataColOnWoWChangeReport}    ${dataColOnWoWChangeReportOnVDC}
                         END
                     END
                     BREAK
                END
            END
        END
    END

    IF    '${table}' == 'OEM West'
        FOR    ${rowIndexOnWoWChangeReport}    IN RANGE    19    28
            Switch Current Excel Document    doc_id=WoWChangeReport
            ${oemGroupColOnWoWChangeReport}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=1
            ${dataColOnWoWChangeReport}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=${posOfColOnWoWChangeReport}
            Switch Current Excel Document    doc_id=WoWChangeReportOnVDC
            FOR    ${rowIndexOnWoWChangeReportOnVDC}    IN RANGE    19   28
                ${oemGroupColOnWoWChangeReportOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReportOnVDC}    col_num=1
                IF    '${nameOfCol}' == 'LW Commit'
                     ${posOfColOnWoWChangeReportOnVDC}   Evaluate    ${posOfColOnWoWChangeReport}+1
                     ${dataColOnWoWChangeReportOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReportOnVDC}    col_num=${posOfColOnWoWChangeReportOnVDC}
                ELSE IF  '${nameOfCol}' == 'Comments'
                     ${dataColOnWoWChangeReportOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReportOnVDC}    col_num=${posOfColOnWoWChangeReport}
                ELSE
                   Close All Excel Documents
                   Fail  The name of column ${nameOfCol} is not valid
                END

                IF    '${oemGroupColOnWoWChangeReport}' == '${oemGroupColOnWoWChangeReportOnVDC}'
                     IF    '${dataColOnWoWChangeReport}' == 'None'
                          ${dataColOnWoWChangeReport}  Set Variable    ${EMPTY}
                     END
                     IF    '${dataColOnWoWChangeReportOnVDC}' == 'None'
                          ${dataColOnWoWChangeReportOnVDC}  Set Variable    ${EMPTY}
                     END

                     IF   '${dataColOnWoWChangeReport}' != '${dataColOnWoWChangeReportOnVDC}'

                         ${result}  Set Variable    ${False}
                         IF    '${oemGroupColOnWoWChangeReport}' == 'OTHERS'
                              Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM West Others    ${dataColOnWoWChangeReport}    ${dataColOnWoWChangeReportOnVDC}
                         ELSE IF  '${oemGroupColOnWoWChangeReport}' == 'Total'
                              Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM West Total    ${dataColOnWoWChangeReport}    ${dataColOnWoWChangeReportOnVDC}
                         ELSE
                              Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    ${oemGroupColOnWoWChangeReport}    ${dataColOnWoWChangeReport}    ${dataColOnWoWChangeReportOnVDC}
                         END
                     END
                     BREAK
                END
            END
        END
    END

    IF    '${result}' == '${False}'
         Close All Excel Documents
         Fail   The ${nameOfCol} data for the ${table} table between the WoW Change Report and WoW Change Report On VDC is different
    END
    Close All Excel Documents

Verify The WoW Data On WoW Change Report
    [Arguments]     ${wowChangeReportFilePath}  ${wowChangeReportOnVDCFilePath}     ${table}    ${posOfColOnWoWChangeReport}    ${nameOfCol}

    ${result}   Set Variable    ${True}

    File Should Exist    ${wowChangeReportFilePath}
    Open Excel Document    ${wowChangeReportFilePath}    doc_id=WoWChangeReport

    File Should Exist    ${wowChangeReportOnVDCFilePath}
    Open Excel Document    ${wowChangeReportOnVDCFilePath}    doc_id=WoWChangeReportOnVDC

    IF    '${table}' == 'Strategic'
        FOR    ${rowIndexOnWoWChangeReport}    IN RANGE    2    8
            Switch Current Excel Document    doc_id=WoWChangeReport
            ${oemGroupColOnWoWChangeReport}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=1
            ${wowColOnWoWChangeReport}               Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=${posOfColOnWoWChangeReport}
            ${posOfDataColOnWoWChangeReport}   Evaluate    ${posOfColOnWoWChangeReport}-1
            ${dataColOnWoWChangeReport}               Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=${posOfDataColOnWoWChangeReport}
            Switch Current Excel Document    doc_id=WoWChangeReportOnVDC
            FOR    ${rowIndexOnWoWChangeReportOnVDC}    IN RANGE    2   8
                ${oemGroupColOnWoWChangeReportOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReportOnVDC}    col_num=1
                IF    '${oemGroupColOnWoWChangeReport}' == '${oemGroupColOnWoWChangeReportOnVDC}'
                    ${dataColOnWoWChangeReportOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReportOnVDC}    col_num=${posOfDataColOnWoWChangeReport}
                    ${wowData}  Evaluate    ${dataColOnWoWChangeReport}-${dataColOnWoWChangeReportOnVDC}
                    ${wowColOnWoWChangeReport}   Evaluate  "%.2f" % ${wowColOnWoWChangeReport}
                    ${wowData}   Evaluate  "%.2f" % ${wowData}
                    IF    '${wowColOnWoWChangeReport}' != '${wowData}'
                        ${result}  Set Variable    ${False}
                         IF    '${oemGroupColOnWoWChangeReport}' == 'OTHERS'
                              Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    Strategic Others    ${wowColOnWoWChangeReport}    ${wowData}
                         ELSE IF  '${oemGroupColOnWoWChangeReport}' == 'Total'
                              Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    Strategic Total    ${wowColOnWoWChangeReport}    ${wowData}
                         ELSE
                              Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    ${oemGroupColOnWoWChangeReport}    ${wowColOnWoWChangeReport}    ${wowData}
                         END
                    END
                    BREAK
                END
            END
        END
    END

    IF    '${table}' == 'OEM East'
        FOR    ${rowIndexOnWoWChangeReport}    IN RANGE    10    17
            Switch Current Excel Document    doc_id=WoWChangeReport
            ${oemGroupColOnWoWChangeReport}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=1
            ${wowColOnWoWChangeReport}               Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=${posOfColOnWoWChangeReport}
            ${posOfDataColOnWoWChangeReport}   Evaluate    ${posOfColOnWoWChangeReport}-1
            ${dataColOnWoWChangeReport}               Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=${posOfDataColOnWoWChangeReport}
            Switch Current Excel Document    doc_id=WoWChangeReportOnVDC
            FOR    ${rowIndexOnWoWChangeReportOnVDC}    IN RANGE    10   17
                ${oemGroupColOnWoWChangeReportOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReportOnVDC}    col_num=1
                IF    '${oemGroupColOnWoWChangeReport}' == '${oemGroupColOnWoWChangeReportOnVDC}'
                    ${dataColOnWoWChangeReportOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReportOnVDC}    col_num=${posOfDataColOnWoWChangeReport}
                    ${wowData}  Evaluate    ${dataColOnWoWChangeReport}-${dataColOnWoWChangeReportOnVDC}
                    ${wowColOnWoWChangeReport}   Evaluate  "%.2f" % ${wowColOnWoWChangeReport}
                    ${wowData}   Evaluate  "%.2f" % ${wowData}
                    IF    '${wowColOnWoWChangeReport}' != '${wowData}'
                        ${result}  Set Variable    ${False}
                         IF    '${oemGroupColOnWoWChangeReport}' == 'OTHERS'
                              Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM East Others    ${wowColOnWoWChangeReport}    ${wowData}
                         ELSE IF  '${oemGroupColOnWoWChangeReport}' == 'Total'
                              Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM East Total    ${wowColOnWoWChangeReport}    ${wowData}
                         ELSE
                              Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    ${oemGroupColOnWoWChangeReport}    ${wowColOnWoWChangeReport}    ${wowData}
                         END
                    END
                    BREAK
                END
            END
        END
    END

    IF    '${table}' == 'OEM West'
        FOR    ${rowIndexOnWoWChangeReport}    IN RANGE    19    28
            Switch Current Excel Document    doc_id=WoWChangeReport
            ${oemGroupColOnWoWChangeReport}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=1
            ${wowColOnWoWChangeReport}               Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=${posOfColOnWoWChangeReport}
            ${posOfDataColOnWoWChangeReport}   Evaluate    ${posOfColOnWoWChangeReport}-1
            ${dataColOnWoWChangeReport}               Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=${posOfDataColOnWoWChangeReport}
            Switch Current Excel Document    doc_id=WoWChangeReportOnVDC
            FOR    ${rowIndexOnWoWChangeReportOnVDC}    IN RANGE    19   28
                ${oemGroupColOnWoWChangeReportOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReportOnVDC}    col_num=1
                IF    '${oemGroupColOnWoWChangeReport}' == '${oemGroupColOnWoWChangeReportOnVDC}'
                    ${dataColOnWoWChangeReportOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReportOnVDC}    col_num=${posOfDataColOnWoWChangeReport}
                    ${wowData}  Evaluate    ${dataColOnWoWChangeReport}-${dataColOnWoWChangeReportOnVDC}
                    ${wowColOnWoWChangeReport}   Evaluate  "%.2f" % ${wowColOnWoWChangeReport}
                    ${wowData}   Evaluate  "%.2f" % ${wowData}
                    IF    '${wowColOnWoWChangeReport}' != '${wowData}'
                        ${result}  Set Variable    ${False}
                         IF    '${oemGroupColOnWoWChangeReport}' == 'OTHERS'
                              Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM West Others    ${wowColOnWoWChangeReport}    ${wowData}
                         ELSE IF  '${oemGroupColOnWoWChangeReport}' == 'Total'
                              Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM West Total    ${wowColOnWoWChangeReport}    ${wowData}
                         ELSE
                              Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    ${oemGroupColOnWoWChangeReport}    ${wowColOnWoWChangeReport}    ${wowData}
                         END
                    END
                    BREAK
                END
            END
        END
    END

    IF    '${result}' == '${False}'
         Close All Excel Documents
         Fail   The ${nameOfCol} data for the ${table} table is wrong
    END
    Close All Excel Documents



