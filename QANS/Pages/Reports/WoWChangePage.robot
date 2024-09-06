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

Check Data For The Strategic Table
    [Arguments]     ${wowChangeReportFilePath}  ${sgWeeklyActionDBReportFilePath}   ${posOfColOnWoWChangeReport}    ${posOfColOnSGWeeklyActionDBReport}     ${nameOfCol}
    ${result}   Set Variable    ${True}
    ${totalDataOnSGWeeklyActionDBReport}    Set Variable    0
    
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
                 ${totalDataOnSGWeeklyActionDBReport}    Evaluate    ${totalDataOnSGWeeklyActionDBReport}+${dataColOnSGWeeklyActionDBReport}
                 IF    ${dataColOnWoWChangeReport} != ${dataColOnSGWeeklyActionDBReport}
                      ${result}     Set Variable    ${False}
                      Write The Test Result Of WoW Change Report To Excel    ${nameOfCol} for Strategic table    ${oemGroupColOnWoWChangeReport}    ${dataColOnWoWChangeReport}    ${dataColOnSGWeeklyActionDBReport}
                 END
                 BREAK
            END
        END
    END

    ${totalDataOnSGWeeklyActionDBReport}   Evaluate  "%.2f" % ${totalDataOnSGWeeklyActionDBReport}
    Switch Current Excel Document    doc_id=WoWChangeReport
    ${totalDataOnWoWchangeReport}   Read Excel Cell    row_num=6    col_num=${posOfColOnWoWChangeReport}
    IF    ${totalDataOnWoWchangeReport} != ${totalDataOnSGWeeklyActionDBReport}
         ${result}     Set Variable    ${False}
         Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    Strategic Total    ${totalDataOnWoWchangeReport}    ${totalDataOnSGWeeklyActionDBReport}
    END

    IF    '${result}' == '${False}'
         Close All Excel Documents
         Fail   The ${nameOfCol} data between the WoW Change Report and SG Weekly Action Report is different
    END
    Close All Excel Documents

Check Data For The OEM East Table
    [Arguments]     ${wowChangeReportFilePath}  ${sgWeeklyActionDBReportFilePath}   ${posOfColOnWoWChangeReport}    ${posOfColOnSGWeeklyActionDBReport}     ${nameOfCol}
    ${result}   Set Variable    ${True}

    File Should Exist    ${wowChangeReportFilePath}
    Open Excel Document    ${wowChangeReportFilePath}    doc_id=WoWChangeReport

    File Should Exist    ${sgWeeklyActionDBReportFilePath}
    Open Excel Document    ${sgWeeklyActionDBReportFilePath}    doc_id=SGWeeklyActionDBReport

#   Verify the data for each OEM Group
    FOR    ${rowIndexOnWoWChangeReport}    IN RANGE    9    14
        Switch Current Excel Document    doc_id=WoWChangeReport
        ${oemGroupColOnWoWChangeReport}      Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=1
        ${dataColOnWoWChangeReport}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=${posOfColOnWoWChangeReport}
        Switch Current Excel Document    doc_id=SGWeeklyActionDBReport
        ${numOfRowsOnSGWeeklyActionDBReport}    Get Number Of Rows In Excel    ${sgWeeklyActionDBReportFilePath}
        FOR    ${rowIndexOnSGWeeklyActionDBReport}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDBReport}+1
            ${oemGroupColOnSGWeeklyActionDBReport}       Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=1
            ${dataColOnSGWeeklyActionDBReport}           Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=${posOfColOnSGWeeklyActionDBReport}
            IF    '${dataColOnSGWeeklyActionDBReport}' == 'None'
                 Continue For Loop
            END
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
    ${totalDataOnSGWeeklyActionDBReport}  Set Variable    0
    FOR    ${rowIndexOnSGWeeklyActionDBReport}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDBReport}+1
        ${mainSalesRepColOnSGWeeklyActionDBReport}      Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=2
        ${dataColOnSGWeeklyActionDBReport}               Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=${posOfColOnSGWeeklyActionDBReport}
        IF    '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Chris Seitz' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Daniel Schmidt' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Eli Tiomkin' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Michael Pauser'
             ${totalDataOnSGWeeklyActionDBReport}     Evaluate    ${totalDataOnSGWeeklyActionDBReport}+${dataColOnSGWeeklyActionDBReport}
        END
    END
    ${totalDataOnSGWeeklyActionDBReport}   Evaluate  "%.2f" % ${totalDataOnSGWeeklyActionDBReport}
    Switch Current Excel Document    doc_id=WoWChangeReport
    ${totalDataOnWoWchangeReport}   Read Excel Cell    row_num=15    col_num=${posOfColOnWoWChangeReport}
    IF    ${totalDataOnWoWchangeReport} != ${totalDataOnSGWeeklyActionDBReport}
         ${result}     Set Variable    ${False}
         Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM East Total    ${totalDataOnWoWchangeReport}    ${totalDataOnSGWeeklyActionDBReport}
    END
 #  Verify the Others data
    Switch Current Excel Document    doc_id=SGWeeklyActionDBReport
    ${othersDataOnSGWeeklyActionDBReport}  Set Variable    0
    FOR    ${rowIndexOnSGWeeklyActionDBReport}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDBReport}+1
        ${oemGroupColOnSGWeeklyActionDBReport}          Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=1
        ${mainSalesRepColOnSGWeeklyActionDBReport}      Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=2
        ${dataColOnSGWeeklyActionDBReport}               Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=${posOfColOnSGWeeklyActionDBReport}
        IF    '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Chris Seitz' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Daniel Schmidt' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Eli Tiomkin' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Michael Pauser'
             IF    '${oemGroupColOnSGWeeklyActionDBReport}' != 'ERICSSON WORLDWIDE' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'MELLANOX GROUP' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'CURTISS WRIGHT GROUP' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'JUNIPER NETWORKS' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'KONTRON NORTH AMERICA'
                  ${othersDataOnSGWeeklyActionDBReport}     Evaluate    ${othersDataOnSGWeeklyActionDBReport}+${dataColOnSGWeeklyActionDBReport}
             END
        END
    END
    ${othersDataOnSGWeeklyActionDBReport}   Evaluate  "%.2f" % ${othersDataOnSGWeeklyActionDBReport}
    Switch Current Excel Document    doc_id=WoWChangeReport
    ${othersDataOnWoWChangeReport}   Read Excel Cell    row_num=14    col_num=${posOfColOnWoWChangeReport}
    IF    ${othersDataOnWoWChangeReport} != ${othersDataOnSGWeeklyActionDBReport}
         ${result}     Set Variable    ${False}
         Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM East Others    ${othersDataOnWoWChangeReport}    ${othersDataOnSGWeeklyActionDBReport}
    END

    IF    '${result}' == '${False}'
         Close All Excel Documents
         Fail   The ${nameOfCol} data for the OEM East table between the WoW Change Report and SG Weekly Action Report is different
    END
    Close All Excel Documents

Check Data For The OEM West Table
    [Arguments]     ${wowChangeReportFilePath}  ${sgWeeklyActionDBReportFilePath}   ${posOfColOnWoWChangeReport}    ${posOfColOnSGWeeklyActionDBReport}     ${nameOfCol}
    ${result}   Set Variable    ${True}

    File Should Exist    ${wowChangeReportFilePath}
    Open Excel Document    ${wowChangeReportFilePath}    doc_id=WoWChangeReport

    File Should Exist    ${sgWeeklyActionDBReportFilePath}
    Open Excel Document    ${sgWeeklyActionDBReportFilePath}    doc_id=SGWeeklyActionDBReport

#   Verify the data for each OEM Group
    ${totalDataOnSGWeeklyActionDBReport}    Set Variable    0
    FOR    ${rowIndexOnWoWChangeReport}    IN RANGE    18    26
        Switch Current Excel Document    doc_id=WoWChangeReport
        ${oemGroupColOnWoWChangeReport}      Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=1
        ${dataColOnWoWChangeReport}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=${posOfColOnWoWChangeReport}
        Switch Current Excel Document    doc_id=SGWeeklyActionDBReport
        ${numOfRowsOnSGWeeklyActionDBReport}    Get Number Of Rows In Excel    ${sgWeeklyActionDBReportFilePath}
        FOR    ${rowIndexOnSGWeeklyActionDBReport}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDBReport}+1
            ${oemGroupColOnSGWeeklyActionDBReport}       Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=1
            ${dataColOnSGWeeklyActionDBReport}           Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=${posOfColOnSGWeeklyActionDBReport}

            IF    '${dataColOnSGWeeklyActionDBReport}' == 'None'
                 ${dataColOnSGWeeklyActionDBReport}     Set Variable    0
            END
            ${dataColOnSGWeeklyActionDBReport}   Evaluate  "%.2f" % ${dataColOnSGWeeklyActionDBReport}
            IF    '${oemGroupColOnWoWChangeReport}' == '${oemGroupColOnSGWeeklyActionDBReport}'
                 ${totalDataOnSGWeeklyActionDBReport}   Evaluate    ${totalDataOnSGWeeklyActionDBReport}+${dataColOnSGWeeklyActionDBReport}
                 IF    ${dataColOnWoWChangeReport} != ${dataColOnSGWeeklyActionDBReport}
                      ${result}     Set Variable    ${False}
                      Write The Test Result Of WoW Change Report To Excel    ${nameOfCol} for OEM West Table    ${oemGroupColOnWoWChangeReport}    ${dataColOnWoWChangeReport}    ${dataColOnSGWeeklyActionDBReport}
                 END
                 BREAK
            END
        END
    END

 #  Verify the Others data
    Switch Current Excel Document    doc_id=SGWeeklyActionDBReport
    ${othersDataOnSGWeeklyActionDBReport}                      Set Variable    0
    ${othersDataForTheOEMWestTableOnSGWeeklyActionDBReport}    Set Variable    0
    ${othersDataForTheStrategicTableOnSGWeeklyActionDBReport}  Set Variable    0
    FOR    ${rowIndexOnSGWeeklyActionDBReport}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDBReport}+1
        ${oemGroupColOnSGWeeklyActionDBReport}          Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=1
        ${mainSalesRepColOnSGWeeklyActionDBReport}      Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=2
        ${dataColOnSGWeeklyActionDBReport}              Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=${posOfColOnSGWeeklyActionDBReport}
        IF    '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Amy Duong' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Caden Douglas' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Michael Nilsson' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Tiger Wang' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Yoda Yasunobu'
             IF    '${oemGroupColOnSGWeeklyActionDBReport}' != 'SCHWEITZER ENGINEERING LABORATORIES (SEL)' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'PANASONIC AVIONICS' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'RADISYS CORPORATION' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'TEKTRONIX' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'TELEDYNE CONTROLS' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'NATIONAL INSTRUMENTS' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'ARROW ELECTRONICS, INC.' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'ASTRONICS CORPORATION'
                  ${othersDataForTheOEMWestTableOnSGWeeklyActionDBReport}     Evaluate    ${othersDataForTheOEMWestTableOnSGWeeklyActionDBReport}+${dataColOnSGWeeklyActionDBReport}
             END
        END
    END

    FOR    ${rowIndexOnSGWeeklyActionDBReport}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDBReport}+1
        ${oemGroupColOnSGWeeklyActionDBReport}          Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=1
        ${mainSalesRepColOnSGWeeklyActionDBReport}      Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=2
        ${dataColOnSGWeeklyActionDBReport}              Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=${posOfColOnSGWeeklyActionDBReport}
        IF    '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Cameron Sinclair' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Huan Tran' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Nicole Lau'
             IF    '${oemGroupColOnSGWeeklyActionDBReport}' != 'NOKIA/ALCATEL LUCENT WORLDWIDE' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'PALO ALTO NETWORKS' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'ARISTA' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'CIENA GROUP'
                  ${othersDataForTheStrategicTableOnSGWeeklyActionDBReport}     Evaluate    ${othersDataForTheStrategicTableOnSGWeeklyActionDBReport}+${dataColOnSGWeeklyActionDBReport}
             END
        END
    END
    ${othersDataOnSGWeeklyActionDBReport}   Evaluate    ${othersDataForTheOEMWestTableOnSGWeeklyActionDBReport}+${othersDataForTheStrategicTableOnSGWeeklyActionDBReport}
    ${othersDataOnSGWeeklyActionDBReport}   Evaluate  "%.2f" % ${othersDataOnSGWeeklyActionDBReport}
    Switch Current Excel Document    doc_id=WoWChangeReport
    ${othersDataOnWoWChangeReport}   Read Excel Cell    row_num=26    col_num=${posOfColOnWoWChangeReport}
    IF    ${othersDataOnWoWChangeReport} != ${othersDataOnSGWeeklyActionDBReport}
         ${result}     Set Variable    ${False}
         Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM West Others    ${othersDataOnWoWChangeReport}    ${othersDataOnSGWeeklyActionDBReport}
    END

    #   Verify the Total data
    ${totalDataOnSGWeeklyActionDBReport}    Evaluate    ${totalDataOnSGWeeklyActionDBReport}+${othersDataOnSGWeeklyActionDBReport}
    ${totalDataOnSGWeeklyActionDBReport}   Evaluate  "%.2f" % ${totalDataOnSGWeeklyActionDBReport}
    Switch Current Excel Document    doc_id=WoWChangeReport
    ${totalDataOnWoWchangeReport}   Read Excel Cell    row_num=27    col_num=${posOfColOnWoWChangeReport}
    IF    ${totalDataOnWoWchangeReport} != ${totalDataOnSGWeeklyActionDBReport}
         ${result}     Set Variable    ${False}
         Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM West Total    ${totalDataOnWoWchangeReport}    ${totalDataOnSGWeeklyActionDBReport}
    END

    IF    '${result}' == '${False}'
         Close All Excel Documents
         Fail   The ${nameOfCol} data for the OEM West table between the WoW Change Report and SG Weekly Action Report is different
    END
    Close All Excel Documents

Check The Commit Or Comment Data
    [Arguments]     ${wowChangeReportFilePath}  ${wowChangeReportOnVDCFilePath}   ${posOfColOnWoWChangeReport}    ${nameOfCol}   ${table}
    ${result}   Set Variable    ${True}

    File Should Exist    ${wowChangeReportFilePath}
    Open Excel Document    ${wowChangeReportFilePath}    doc_id=WoWChangeReport

    File Should Exist    ${wowChangeReportOnVDCFilePath}
    Open Excel Document    ${wowChangeReportOnVDCFilePath}    doc_id=WoWChangeReportOnVDC

    IF    '${table}' == 'Strategic'
         FOR    ${rowIndexOnWoWChangeReport}    IN RANGE    2    7
            Switch Current Excel Document    doc_id=WoWChangeReport
            ${oemGroupColOnWoWChangeReport}      Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=1
            ${dataColOnWoWChangeReport}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=${posOfColOnWoWChangeReport}
            IF    '${nameOfCol}' == 'TW Commit'
                 IF    '${dataColOnWoWChangeReport}' != '${EMPTY}'
                      ${result}  Set Variable    ${False}
                      IF  '${oemGroupColOnWoWChangeReport}' == 'Total'
                           Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    Strategic Total    ${dataColOnWoWChangeReport}    ${EMPTY}
                      ELSE
                           Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    ${oemGroupColOnWoWChangeReport}    ${dataColOnWoWChangeReport}    ${EMPTY}
                      END
                 END
                 Continue For Loop
            END
            Switch Current Excel Document    doc_id=WoWChangeReportOnVDC
            FOR    ${rowIndexOnWoWChangeReportOnVDC}    IN RANGE    2   7
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
                         IF  '${oemGroupColOnWoWChangeReport}' == 'Total'
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
        FOR    ${rowIndexOnWoWChangeReport}    IN RANGE    9    16
            Switch Current Excel Document    doc_id=WoWChangeReport
            ${oemGroupColOnWoWChangeReport}      Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=1
            ${dataColOnWoWChangeReport}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=${posOfColOnWoWChangeReport}
            IF    '${nameOfCol}' == 'TW Commit'
                 IF    '${dataColOnWoWChangeReport}' != '${EMPTY}'
                      ${result}  Set Variable    ${False}
                      IF    '${oemGroupColOnWoWChangeReport}' == 'OTHERS'
                           Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM East Others    ${dataColOnWoWChangeReport}    ${EMPTY}
                      ELSE IF  '${oemGroupColOnWoWChangeReport}' == 'Total'
                           Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM East Total    ${dataColOnWoWChangeReport}    ${EMPTY}
                      ELSE
                           Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    ${oemGroupColOnWoWChangeReport}    ${dataColOnWoWChangeReport}    ${EMPTY}
                      END
                 END
                 Continue For Loop
            END
            Switch Current Excel Document    doc_id=WoWChangeReportOnVDC
            FOR    ${rowIndexOnWoWChangeReportOnVDC}    IN RANGE    9   16
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
        FOR    ${rowIndexOnWoWChangeReport}    IN RANGE    18    28
            Switch Current Excel Document    doc_id=WoWChangeReport
            ${oemGroupColOnWoWChangeReport}      Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=1
            ${dataColOnWoWChangeReport}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=${posOfColOnWoWChangeReport}
            IF    '${nameOfCol}' == 'TW Commit'
                 IF    '${dataColOnWoWChangeReport}' != '${EMPTY}'
                      ${result}  Set Variable    ${False}
                      IF    '${oemGroupColOnWoWChangeReport}' == 'OTHERS'
                           Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM West Others    ${dataColOnWoWChangeReport}    ${EMPTY}
                      ELSE IF  '${oemGroupColOnWoWChangeReport}' == 'Total'
                           Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM West Total    ${dataColOnWoWChangeReport}    ${EMPTY}
                      ELSE
                           Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    ${oemGroupColOnWoWChangeReport}    ${dataColOnWoWChangeReport}    ${EMPTY}
                      END
                 END
                 Continue For Loop
            END
            Switch Current Excel Document    doc_id=WoWChangeReportOnVDC
            FOR    ${rowIndexOnWoWChangeReportOnVDC}    IN RANGE    18   28
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
                         Log To Console    dataColOnWoWChangeReport:${dataColOnWoWChangeReport};dataColOnWoWChangeReportOnVDC:${dataColOnWoWChangeReportOnVDC}
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

Check The WoW Data
    [Arguments]     ${wowChangeReportFilePath}  ${wowChangeReportOnVDCFilePath}     ${table}    ${posOfColOnWoWChangeReport}    ${nameOfCol}

    ${result}   Set Variable    ${True}

    File Should Exist    ${wowChangeReportFilePath}
    Open Excel Document    ${wowChangeReportFilePath}    doc_id=WoWChangeReport

    File Should Exist    ${wowChangeReportOnVDCFilePath}
    Open Excel Document    ${wowChangeReportOnVDCFilePath}    doc_id=WoWChangeReportOnVDC

    IF    '${table}' == 'Strategic'
        FOR    ${rowIndexOnWoWChangeReport}    IN RANGE    2    7
            Switch Current Excel Document    doc_id=WoWChangeReport
            ${oemGroupColOnWoWChangeReport}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=1
            ${wowColOnWoWChangeReport}               Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=${posOfColOnWoWChangeReport}
            ${posOfDataColOnWoWChangeReport}   Evaluate    ${posOfColOnWoWChangeReport}-1
            ${dataColOnWoWChangeReport}               Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=${posOfDataColOnWoWChangeReport}
            Switch Current Excel Document    doc_id=WoWChangeReportOnVDC
            FOR    ${rowIndexOnWoWChangeReportOnVDC}    IN RANGE    2   7
                ${oemGroupColOnWoWChangeReportOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReportOnVDC}    col_num=1
                IF    '${oemGroupColOnWoWChangeReport}' == '${oemGroupColOnWoWChangeReportOnVDC}'
                    ${dataColOnWoWChangeReportOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReportOnVDC}    col_num=${posOfDataColOnWoWChangeReport}
                    ${wowData}  Evaluate    ${dataColOnWoWChangeReport}-${dataColOnWoWChangeReportOnVDC}
                    ${wowColOnWoWChangeReport}   Evaluate  "%.2f" % ${wowColOnWoWChangeReport}
                    ${wowData}   Evaluate  "%.2f" % ${wowData}
                    IF    '${wowColOnWoWChangeReport}' == '-0.00'
                         ${wowColOnWoWChangeReport}     Set Variable    0.00
                    END
                    IF    '${wowColOnWoWChangeReport}' != '${wowData}'
                        ${result}  Set Variable    ${False}
                         IF  '${oemGroupColOnWoWChangeReport}' == 'Total'
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
        FOR    ${rowIndexOnWoWChangeReport}    IN RANGE    9    16
            Switch Current Excel Document    doc_id=WoWChangeReport
            ${oemGroupColOnWoWChangeReport}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=1
            ${wowColOnWoWChangeReport}               Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=${posOfColOnWoWChangeReport}
            ${posOfDataColOnWoWChangeReport}   Evaluate    ${posOfColOnWoWChangeReport}-1
            ${dataColOnWoWChangeReport}               Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=${posOfDataColOnWoWChangeReport}
            Switch Current Excel Document    doc_id=WoWChangeReportOnVDC
            FOR    ${rowIndexOnWoWChangeReportOnVDC}    IN RANGE    9   16
                ${oemGroupColOnWoWChangeReportOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReportOnVDC}    col_num=1
                IF    '${oemGroupColOnWoWChangeReport}' == '${oemGroupColOnWoWChangeReportOnVDC}'
                    ${dataColOnWoWChangeReportOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReportOnVDC}    col_num=${posOfDataColOnWoWChangeReport}
                    ${wowData}  Evaluate    ${dataColOnWoWChangeReport}-${dataColOnWoWChangeReportOnVDC}
                    ${wowColOnWoWChangeReport}   Evaluate  "%.2f" % ${wowColOnWoWChangeReport}
                    ${wowData}   Evaluate  "%.2f" % ${wowData}
                    IF    '${wowColOnWoWChangeReport}' == '-0.00'
                         ${wowColOnWoWChangeReport}     Set Variable    0.00
                    END
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
        FOR    ${rowIndexOnWoWChangeReport}    IN RANGE    18    28
            Switch Current Excel Document    doc_id=WoWChangeReport
            ${oemGroupColOnWoWChangeReport}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=1
            ${wowColOnWoWChangeReport}               Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=${posOfColOnWoWChangeReport}
            ${posOfDataColOnWoWChangeReport}   Evaluate    ${posOfColOnWoWChangeReport}-1
            ${dataColOnWoWChangeReport}               Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=${posOfDataColOnWoWChangeReport}
            Switch Current Excel Document    doc_id=WoWChangeReportOnVDC
            FOR    ${rowIndexOnWoWChangeReportOnVDC}    IN RANGE    18   28
                ${oemGroupColOnWoWChangeReportOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReportOnVDC}    col_num=1
                IF    '${oemGroupColOnWoWChangeReport}' == '${oemGroupColOnWoWChangeReportOnVDC}'
                    ${dataColOnWoWChangeReportOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReportOnVDC}    col_num=${posOfDataColOnWoWChangeReport}
                    ${wowData}  Evaluate    ${dataColOnWoWChangeReport}-${dataColOnWoWChangeReportOnVDC}
                    ${wowColOnWoWChangeReport}   Evaluate  "%.2f" % ${wowColOnWoWChangeReport}
                    ${wowData}   Evaluate  "%.2f" % ${wowData}
                    IF    '${wowColOnWoWChangeReport}' == '-0.00'
                         ${wowColOnWoWChangeReport}     Set Variable    0.00
                    END
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

Check The GAP Data
    [Arguments]     ${wowChangeReportOnVDCFilePath}    ${sgWeeklyActionDBReportFilePath}     ${wowChangeReportFilePath}    ${table}    ${posOfColOnWoWChangeReport}    ${nameOfCol}
    ${result}   Set Variable    ${True}

    File Should Exist    ${wowChangeReportOnVDCFilePath}
    Open Excel Document    ${wowChangeReportOnVDCFilePath}    doc_id=WoWChangeReportOnVDC

    File Should Exist    ${sgWeeklyActionDBReportFilePath}
    Open Excel Document    ${sgWeeklyActionDBReportFilePath}    doc_id=SGWeeklyActionDBReport
    Switch Current Excel Document    doc_id=SGWeeklyActionDBReport
    ${numOfRowsOnSGWeeklyActionDBReport}    Get Number Of Rows In Excel    ${sgWeeklyActionDBReportFilePath}

    File Should Exist    ${wowChangeReportFilePath}
    Open Excel Document    ${wowChangeReportFilePath}    doc_id=WoWChangeReport

    IF    '${table}' == 'Strategic'
        FOR    ${rowIndexOnWoWChangeReport}    IN RANGE    2    7
            Switch Current Excel Document    doc_id=WoWChangeReport
            ${oemGroupColOnWoWChangeReport}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=1
            ${gapColOnWoWChangeReport}               Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=${posOfColOnWoWChangeReport}
            ${gapColOnWoWChangeReport}   Evaluate  "%.2f" % ${gapColOnWoWChangeReport}

            ${losData}      Set Variable    0
            Switch Current Excel Document    doc_id=SGWeeklyActionDBReport
            IF    '${oemGroupColOnWoWChangeReport}' != 'Total'                
                 FOR    ${rowIndexOnSGWeeklyActionDBReport}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDBReport}+1
                     ${oemGroupColOnSGWeeklyActionDBReport}     Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=1
                     ${losColOnSGWeeklyActionDBReport}          Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=7
                     IF    '${oemGroupColOnWoWChangeReport}' == '${oemGroupColOnSGWeeklyActionDBReport}'
                         ${losData}     Set Variable    ${losColOnSGWeeklyActionDBReport}
                         BREAK
                     END
                 END
            ELSE               
                FOR    ${rowIndexOnSGWeeklyActionDBReport}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDBReport}+1
                    ${oemGroupColOnSGWeeklyActionDBReport}      Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=1
                    ${mainSalesRepColOnSGWeeklyActionDBReport}  Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=2
                    ${losColOnSGWeeklyActionDBReport}           Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=7
                    IF    '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Cameron Sinclair' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Huan Tran' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Nicole Lau' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Dave Beasley'
                          IF    '${oemGroupColOnSGWeeklyActionDBReport}' == 'NOKIA/ALCATEL LUCENT WORLDWIDE' or '${oemGroupColOnSGWeeklyActionDBReport}' == 'PALO ALTO NETWORKS' or '${oemGroupColOnSGWeeklyActionDBReport}' == 'ARISTA' or '${oemGroupColOnSGWeeklyActionDBReport}' == 'CIENA GROUP'
                               ${losData}   Evaluate    ${losData}+${losColOnSGWeeklyActionDBReport}
                          END
                    END
                END
            END
            Switch Current Excel Document    doc_id=WoWChangeReportOnVDC
            ${commitData}   Set Variable    0
            FOR    ${rowIndexOnWoWChangeReportOnVDC}    IN RANGE    2    7
                ${oemGroupColOnWoWChangeReportOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReportOnVDC}    col_num=1
                ${twCommitColOnWoWChangeReportOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReportOnVDC}    col_num=5

                IF    '${oemGroupColOnWoWChangeReport}' == '${oemGroupColOnWoWChangeReportOnVDC}'
                    ${commitData}   Set Variable    ${twCommitColOnWoWChangeReportOnVDC}
                    BREAK
                END
            END
            Log To Console    LOS:${losData};COMMIT:${commitData}
            ${gapDataByFormular}    Evaluate    ${losData}-${commitData}
            ${gapDataByFormular}   Evaluate  "%.2f" % ${gapDataByFormular}

            IF    '${gapColOnWoWChangeReport}' != '${gapDataByFormular}'
                  ${result}   Set Variable    ${False}
                  IF    '${oemGroupColOnWoWChangeReport}' == 'Total'
                       Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    Strategic Total     ${gapColOnWoWChangeReport}    ${gapDataByFormular}
                  ELSE
                       Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    ${oemGroupColOnWoWChangeReport}    ${gapColOnWoWChangeReport}    ${gapDataByFormular}
                  END
            END
        END
    END

    IF    '${table}' == 'OEM East'
        FOR    ${rowIndexOnWoWChangeReport}    IN RANGE    9    16
            Switch Current Excel Document    doc_id=WoWChangeReport
            ${oemGroupColOnWoWChangeReport}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=1
            ${gapColOnWoWChangeReport}               Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=${posOfColOnWoWChangeReport}
            ${gapColOnWoWChangeReport}   Evaluate  "%.2f" % ${gapColOnWoWChangeReport}

            ${losData}      Set Variable    0
            Switch Current Excel Document    doc_id=SGWeeklyActionDBReport
            IF  '${oemGroupColOnWoWChangeReport}' == 'OTHERS'
                 FOR    ${rowIndexOnSGWeeklyActionDBReport}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDBReport}+1
                     ${oemGroupColOnSGWeeklyActionDBReport}        Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=1
                     ${mainSalesRepColOnSGWeeklyActionDBReport}    Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=2
                     ${losColOnSGWeeklyActionDBReport}             Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=7
                     IF    '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Chris Seitz' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Daniel Schmidt' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Eli Tiomkin' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Michael Pauser'
                          IF    '${oemGroupColOnSGWeeklyActionDBReport}' != 'ERICSSON WORLDWIDE' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'MELLANOX GROUP' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'CURTISS WRIGHT GROUP' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'JUNIPER NETWORKS' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'KONTRON NORTH AMERICA'
                                ${losData}   Evaluate    ${losData}+${losColOnSGWeeklyActionDBReport}
                          END
                     END

                 END
            ELSE IF  '${oemGroupColOnWoWChangeReport}' == 'Total'
                FOR    ${rowIndexOnSGWeeklyActionDBReport}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDBReport}+1
                    ${oemGroupColOnSGWeeklyActionDBReport}      Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=1
                    ${mainSalesRepColOnSGWeeklyActionDBReport}  Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=2
                    ${losColOnSGWeeklyActionDBReport}           Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=7
                    IF    '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Chris Seitz' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Daniel Schmidt' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Eli Tiomkin' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Michael Pauser'
                          ${losData}   Evaluate    ${losData}+${losColOnSGWeeklyActionDBReport}
                    END
                END
            ELSE
                FOR    ${rowIndexOnSGWeeklyActionDBReport}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDBReport}+1
                     ${oemGroupColOnSGWeeklyActionDBReport}     Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=1
                     ${losColOnSGWeeklyActionDBReport}          Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=7
                     IF    '${oemGroupColOnWoWChangeReport}' == '${oemGroupColOnSGWeeklyActionDBReport}'
                         ${losData}     Set Variable    ${losColOnSGWeeklyActionDBReport}
                         BREAK
                     END
                 END
            END

            Switch Current Excel Document    doc_id=WoWChangeReportOnVDC
            ${commitData}   Set Variable    0
            FOR    ${rowIndexOnWoWChangeReportOnVDC}    IN RANGE    9    16
                ${oemGroupColOnWoWChangeReportOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReportOnVDC}    col_num=1
                ${twCommitColOnWoWChangeReportOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReportOnVDC}    col_num=5

                IF    '${oemGroupColOnWoWChangeReport}' == '${oemGroupColOnWoWChangeReportOnVDC}'
                    ${commitData}   Set Variable    ${twCommitColOnWoWChangeReportOnVDC}
                    BREAK
                END
            END
            ${gapDataByFormular}    Evaluate    ${losData}-${commitData}
            ${gapDataByFormular}   Evaluate  "%.2f" % ${gapDataByFormular}

            IF    '${gapColOnWoWChangeReport}' != '${gapDataByFormular}'
                  ${result}   Set Variable    ${False}
                  IF    '${oemGroupColOnWoWChangeReport}' == 'Total'
                       Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM East Total     ${gapColOnWoWChangeReport}    ${gapDataByFormular}
                  ELSE IF  '${oemGroupColOnWoWChangeReport}' == 'OTHERS'
                       Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM East OTHERS     ${gapColOnWoWChangeReport}    ${gapDataByFormular}
                  ELSE
                       Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    ${oemGroupColOnWoWChangeReport}    ${gapColOnWoWChangeReport}    ${gapDataByFormular}
                  END
            END
        END
    END

    IF    '${table}' == 'OEM West'
        FOR    ${rowIndexOnWoWChangeReport}    IN RANGE    18    28
            Switch Current Excel Document    doc_id=WoWChangeReport
            ${oemGroupColOnWoWChangeReport}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=1
            ${gapColOnWoWChangeReport}               Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=${posOfColOnWoWChangeReport}
            ${gapColOnWoWChangeReport}   Evaluate  "%.2f" % ${gapColOnWoWChangeReport}

            ${losData}      Set Variable    0
            Switch Current Excel Document    doc_id=SGWeeklyActionDBReport
            IF  '${oemGroupColOnWoWChangeReport}' == 'OTHERS'
                 FOR    ${rowIndexOnSGWeeklyActionDBReport}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDBReport}+1
                     ${oemGroupColOnSGWeeklyActionDBReport}        Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=1
                     ${mainSalesRepColOnSGWeeklyActionDBReport}    Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=2
                     ${losColOnSGWeeklyActionDBReport}             Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=7
                     IF    '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Cameron Sinclair' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Huan Tran' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Nicole Lau' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Dave Beasley'
                        IF    '${oemGroupColOnSGWeeklyActionDBReport}' != 'NOKIA/ALCATEL LUCENT WORLDWIDE' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'PALO ALTO NETWORKS' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'ARISTA' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'CIENA GROUP'
                            ${losData}   Evaluate    ${losData}+${losColOnSGWeeklyActionDBReport}
                        END
                     END
                     IF    '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Amy Duong' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Caden Douglas' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Michael Nilsson' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Tiger Wang' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Yoda Yasunobu'
                          IF    '${oemGroupColOnSGWeeklyActionDBReport}' != 'SCHWEITZER ENGINEERING LABORATORIES (SEL)' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'PANASONIC AVIONICS' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'RADISYS CORPORATION' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'TEKTRONIX' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'TELEDYNE CONTROLS' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'NATIONAL INSTRUMENTS' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'ARROW ELECTRONICS, INC.' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'ASTRONICS CORPORATION'
                                ${losData}   Evaluate    ${losData}+${losColOnSGWeeklyActionDBReport}
                          END
                     END

                 END
            ELSE IF  '${oemGroupColOnWoWChangeReport}' == 'Total'
                FOR    ${rowIndexOnSGWeeklyActionDBReport}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDBReport}+1
                    ${oemGroupColOnSGWeeklyActionDBReport}      Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=1
                    ${mainSalesRepColOnSGWeeklyActionDBReport}  Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=2
                    ${losColOnSGWeeklyActionDBReport}           Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=7
                    
                    IF    '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Cameron Sinclair' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Huan Tran' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Nicole Lau' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Dave Beasley'
                        IF    '${oemGroupColOnSGWeeklyActionDBReport}' != 'NOKIA/ALCATEL LUCENT WORLDWIDE' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'PALO ALTO NETWORKS' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'ARISTA' and '${oemGroupColOnSGWeeklyActionDBReport}' != 'CIENA GROUP'
                            ${losData}   Evaluate    ${losData}+${losColOnSGWeeklyActionDBReport}
                        END
                    END
                    
                    IF    '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Amy Duong' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Caden Douglas' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Michael Nilsson' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Tiger Wang' or '${mainSalesRepColOnSGWeeklyActionDBReport}' == 'Yoda Yasunobu'
                          ${losData}   Evaluate    ${losData}+${losColOnSGWeeklyActionDBReport}
                    END
                END
            ELSE
                FOR    ${rowIndexOnSGWeeklyActionDBReport}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDBReport}+1
                     ${oemGroupColOnSGWeeklyActionDBReport}     Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=1
                     ${losColOnSGWeeklyActionDBReport}          Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=7
                     IF    '${oemGroupColOnWoWChangeReport}' == '${oemGroupColOnSGWeeklyActionDBReport}'
                         ${losData}     Set Variable    ${losColOnSGWeeklyActionDBReport}
                         BREAK
                     END
                 END
            END

            Switch Current Excel Document    doc_id=WoWChangeReportOnVDC
            ${commitData}   Set Variable    0
            FOR    ${rowIndexOnWoWChangeReportOnVDC}    IN RANGE    18    28
                ${oemGroupColOnWoWChangeReportOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReportOnVDC}    col_num=1
                ${twCommitColOnWoWChangeReportOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReportOnVDC}    col_num=5

                IF    '${oemGroupColOnWoWChangeReport}' == '${oemGroupColOnWoWChangeReportOnVDC}'
                    ${commitData}   Set Variable    ${twCommitColOnWoWChangeReportOnVDC}
                    BREAK
                END
            END
            ${gapDataByFormular}    Evaluate    ${losData}-${commitData}
            ${gapDataByFormular}   Evaluate  "%.2f" % ${gapDataByFormular}

            IF    '${gapColOnWoWChangeReport}' != '${gapDataByFormular}'
                  ${result}   Set Variable    ${False}
                  IF    '${oemGroupColOnWoWChangeReport}' == 'Total'
                       Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM West Total     ${gapColOnWoWChangeReport}    ${gapDataByFormular}
                  ELSE IF  '${oemGroupColOnWoWChangeReport}' == 'OTHERS'
                       Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM West OTHERS     ${gapColOnWoWChangeReport}    ${gapDataByFormular}
                  ELSE
                       Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    ${oemGroupColOnWoWChangeReport}    ${gapColOnWoWChangeReport}    ${gapDataByFormular}
                  END
            END
        END
    END

    IF    '${result}' == '${False}'
         Close All Excel Documents
         Fail   The ${nameOfCol} data for the ${table} table is wrong
    END
    Close All Excel Documents



