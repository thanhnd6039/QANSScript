*** Settings ***
Resource    ../CommonPage.robot
Resource    SGPage.robot

*** Variables ***
${wowChangeResultFilePath}        C:\\RobotFramework\\Results\\WoWChangeReportResult.xlsx
${wowChangeFilePath}              C:\\RobotFramework\\Downloads\\Wow Change [Current Week].xlsx
${wowChangeOnVDCFilePath}         C:\\RobotFramework\\Downloads\\Wow Change [Current Week] On VDC.xlsx
${SGFilePath}                     C:\\RobotFramework\\Downloads\\Sales Gap Report NS With SO Forecast.xlsx
${SGWeeklyActionDBFilePath}       C:\\RobotFramework\\Downloads\\SalesGap Weekly Actions Gap Week DB.xlsx

*** Keywords ***
Write The Test Result Of WoW Change Report To Excel
    [Arguments]     ${item}     ${oemGroup}     ${valueOnWoWChange}   ${valueOnSG}
    File Should Exist      path=${wowChangeResultFilePath}
    Open Excel Document    filename=${wowChangeResultFilePath}    doc_id=WoWChangeReportResult
    Switch Current Excel Document    doc_id=WoWChangeReportResult
    ${latestRowInWoWchangeResult}   Get Number Of Rows In Excel    ${wowChangeResultFilePath}
    ${nextRow}    Evaluate    ${latestRowInWoWchangeResult}+1
    Write Excel Cell    row_num=${nextRow}    col_num=1    value=${item}
    Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oemGroup}
    Write Excel Cell    row_num=${nextRow}    col_num=3    value=${valueOnWoWChange}
    Write Excel Cell    row_num=${nextRow}    col_num=4    value=${valueOnSG}
    Save Excel Document    ${wowChangeResultFilePath}
    Close Current Excel Document

Get List Of Sales Member In OEM East Table
    @{listOfSalesMember}    Create List
    Append To List    ${listOfSalesMember}      Chris Seitz
    Append To List    ${listOfSalesMember}      Daniel Schmidt
    Append To List    ${listOfSalesMember}      Eli Tiomkin
    Append To List    ${listOfSalesMember}      Michael Pauser

    [Return]    ${listOfSalesMember}

Get List Of Sales Member In OEM West Table
    @{listOfSalesMember}    Create List
    Append To List    ${listOfSalesMember}      Amy Duong
    Append To List    ${listOfSalesMember}      Caden Douglas
    Append To List    ${listOfSalesMember}      Michael Nilsson
    Append To List    ${listOfSalesMember}      Tiger Wang
    Append To List    ${listOfSalesMember}      Yoda Yasunobu
    Append To List    ${listOfSalesMember}      Nicole Lau
    Append To List    ${listOfSalesMember}      Huan Tran
    Append To List    ${listOfSalesMember}      Dave Beasley
    Append To List    ${listOfSalesMember}      Cameron Sinclair
    Append To List    ${listOfSalesMember}      Clint Stalker

    [Return]    ${listOfSalesMember}

Get List Of OEM Group Shown In OEM East Table
    @{listOfOEMGroup}   Create List
    Append To List    ${listOfOEMGroup}      NVIDIA/MELLANOX
    Append To List    ${listOfOEMGroup}      NOKIA/ALCATEL LUCENT WORLDWIDE
    Append To List    ${listOfOEMGroup}      CURTISS WRIGHT GROUP
    Append To List    ${listOfOEMGroup}      JUNIPER NETWORKS
    Append To List    ${listOfOEMGroup}      ERICSSON WORLDWIDE

    [Return]    ${listOfOEMGroup}

Get List Of OEM Group Shown In OEM West Table
    @{listOfOEMGroup}   Create List
    Append To List    ${listOfOEMGroup}      PALO ALTO NETWORKS
    Append To List    ${listOfOEMGroup}      ARISTA
    Append To List    ${listOfOEMGroup}      SCHWEITZER ENGINEERING LABORATORIES (SEL)
    Append To List    ${listOfOEMGroup}      PANASONIC AVIONICS
    Append To List    ${listOfOEMGroup}      ZTE KANGXUN TELECOM CO. LTD.
    Append To List    ${listOfOEMGroup}      NATIONAL INSTRUMENTS

    [Return]    ${listOfOEMGroup}

Check The Ship, Backlog, LOS Data
    [Arguments]     ${table}    ${nameOfCol}    ${posOfColOnWoWChange}    ${posOfRColOnSG}=0      ${posOfBColOnSG}=0

    ${result}   Set Variable    ${True}
    ${startRowIndexForOEMGroupOnWoWChange}   Set Variable    0
    ${endRowIndexForOEMGroupOnWoWChange}     Set Variable    0
    ${othersRowIndexOnWoWChange}             Set Variable    0
    ${totalRowIndexOnWoWChange}              Set Variable    0

    ${listOfSalesMemberInOEMEastTable}       Get List Of Sales Member In OEM East Table
    ${listOfOEMGroupShownInOEMEastTable}     Get List Of OEM Group Shown In OEM East Table
    ${listOfSalesMemberInOEMWestTable}       Get List Of Sales Member In OEM West Table
    ${listOfOEMGroupShownInOEMWestTable}     Get List Of OEM Group Shown In OEM West Table

    File Should Exist      path=${wowChangeFilePath}
    Open Excel Document    filename=${wowChangeFilePath}           doc_id=WoWChange
    ${numOfRowsOnWoWChange}  Get Number Of Rows In Excel    filePath=${wowChangeFilePath}

    FOR    ${rowIndexOnWoWChange}    IN RANGE    1    ${numOfRowsOnWoWChange}+1
        ${oemGroupColOnWoWChange}     Read Excel Cell    row_num=${rowIndexOnWoWChange}    col_num=1
        IF    '${oemGroupColOnWoWChange}' == '${table}'
            ${startRowIndexForOEMGroupOnWoWChange}  Evaluate    ${rowIndexOnWoWChange}+1
            BREAK
        END
    END

    IF    '${table}' == 'OEM East'
        ${countOthers}  Set Variable    0
        FOR    ${rowIndexOnWoWChange}    IN RANGE    1    ${numOfRowsOnWoWChange}+1
            ${oemGroupColOnWoWChange}     Read Excel Cell    row_num=${rowIndexOnWoWChange}    col_num=1
            IF    '${oemGroupColOnWoWChange}' == 'OTHERS'
                 ${countOthers}     Evaluate    ${countOthers}+1
            END
            IF    '${countOthers}' == '1'
                 ${endRowIndexForOEMGroupOnWoWChange}   Evaluate    ${rowIndexOnWoWChange}-1
                 ${othersRowIndexOnWoWChange}   Set Variable    ${rowIndexOnWoWChange}
                 ${totalRowIndexOnWoWChange}    Evaluate    ${rowIndexOnWoWChange}+1
                 BREAK
            END
        END

    ELSE IF     '${table}' == 'OEM West + Channel'
        ${countOthers}  Set Variable    0
        FOR    ${rowIndexOnWoWChange}    IN RANGE    1    ${numOfRowsOnWoWChange}+1
            ${oemGroupColOnWoWChange}     Read Excel Cell    row_num=${rowIndexOnWoWChange}    col_num=1
            IF    '${oemGroupColOnWoWChange}' == 'OTHERS'
                 ${countOthers}     Evaluate    ${countOthers}+1
            END
            IF    '${countOthers}' == '2'
                 ${endRowIndexForOEMGroupOnWoWChange}   Evaluate    ${rowIndexOnWoWChange}-1
                 ${othersRowIndexOnWoWChange}   Set Variable    ${rowIndexOnWoWChange}
                 ${totalRowIndexOnWoWChange}    Evaluate    ${rowIndexOnWoWChange}+1
                 BREAK
            END
        END
    ELSE
        Fail    The table parameter ${table} is invalid. Please contact with the Administrator for supporting
    END

   #   Verify the data for each OEM Group
    FOR    ${rowIndexOnWoWChange}    IN RANGE    ${startRowIndexForOEMGroupOnWoWChange}    ${endRowIndexForOEMGroupOnWoWChange}+1

        ${oemGroupColOnWoWChange}      Read Excel Cell    row_num=${rowIndexOnWoWChange}    col_num=1
        ${dataOnWoWChange}          Read Excel Cell    row_num=${rowIndexOnWoWChange}    col_num=${posOfColOnWoWChange}
        IF    '${nameOfCol}' == 'Ships' or '${nameOfCol}' == 'Pre Q Ships'
            ${dataOnSG}     Get The Amount Data By OEM Group On SG Report    oemGroup=${oemGroupColOnWoWChange}    posOfCol=${posOfRColOnSG}
        ELSE IF  '${nameOfCol}' == 'Backlog'
            ${dataOnSG}     Get The Amount Data By OEM Group On SG Report    oemGroup=${oemGroupColOnWoWChange}    posOfCol=${posOfBColOnSG}
        ELSE IF  '${nameOfCol}' == 'LOS'
            ${dataROnSG}     Get The Amount Data By OEM Group On SG Report    oemGroup=${oemGroupColOnWoWChange}    posOfCol=${posOfRColOnSG}
            ${dataBOnSG}     Get The Amount Data By OEM Group On SG Report    oemGroup=${oemGroupColOnWoWChange}    posOfCol=${posOfBColOnSG}
            ${dataOnSG}     Evaluate    ${dataROnSG}+${dataBOnSG}
        END

        ${dataOnWoWChange}          Evaluate  "%.2f" % ${dataOnWoWChange}
        ${dataOnSG}                 Evaluate  "%.2f" % ${dataOnSG}
        Log To Console    OEM GRoup:${oemGroupColOnWoWChange}; dataOnWoWChange:${dataOnWoWChange}; dataOnSG:${dataOnSG}
        IF    ${dataOnWoWChange} != ${dataOnSG}
              ${result}     Set Variable    ${False}
              Write The Test Result Of WoW Change Report To Excel    item=${nameOfCol}    oemGroup=${oemGroupColOnWoWChange}    valueOnWoWChange=${dataOnWoWChange}    valueOnSG=${dataOnSG}
        END

    END
#    #   Verify the Total data
#    Switch Current Excel Document    doc_id=SG
#    ${totalOnSG}  Set Variable    0
#    FOR    ${rowIndexOnSG}    IN RANGE    6    ${numOfRowsOnSG}+1
#        ${mainSalesRepColOnSG}      Read Excel Cell    row_num=${rowIndexOnSG}    col_num=3
#        IF    '${nameOfCol}' == 'Ships' or '${nameOfCol}' == 'Pre Q Ships'
#             ${dataColOnSG}           Read Excel Cell    row_num=${rowIndexOnSG}    col_num=${posOfRColOnSG}
#        ELSE IF  '${nameOfCol}' == 'Backlog'
#             ${dataColOnSG}           Read Excel Cell    row_num=${rowIndexOnSG}    col_num=${posOfBColOnSG}
#        ELSE IF  '${nameOfCol}' == 'LOS'
#             ${dataRColOnSG}           Read Excel Cell    row_num=${rowIndexOnSG}    col_num=${posOfRColOnSG}
#             ${dataBolOnSG}           Read Excel Cell    row_num=${rowIndexOnSG}    col_num=${posOfBColOnSG}
#             IF    '${dataRColOnSG}' == 'None'
#                  ${dataRColOnSG}   Set Variable    0
#             END
#             IF    '${dataBolOnSG}' == 'None'
#                  ${dataBolOnSG}    Set Variable    0
#             END
#             ${dataColOnSG}     Evaluate    ${dataRColOnSG}+${dataBolOnSG}
#        ELSE
#             Fail   The nameOfCol parameter ${nameOfCol} is invalid. Please contact with the Administrator for supporting
#        END
#
#        IF    '${dataColOnSG}' == 'None'
#               ${dataColOnSG}     Set Variable    0
#        END
#
#        IF    '${table}' == 'OEM East'
#             IF    '${mainSalesRepColOnSG}' in ${listOfSalesMemberInOEMEastTable}
#                   ${totalOnSG}     Evaluate    ${totalOnSG}+${dataColOnSG}
#             END
#        ELSE IF   '${table}' == 'OEM West'
#            IF    '${mainSalesRepColOnSG}' in ${listOfSalesMemberInOEMWestTable}
#                   ${totalOnSG}     Evaluate    ${totalOnSG}+${dataColOnSG}
#            END
#        ELSE
#            Fail    The table parameter ${table} is invalid. Please contact with the Administrator for supporting
#        END
#    END
#    ${totalOnSG}   Evaluate  "%.2f" % ${totalOnSG}
#    Switch Current Excel Document    doc_id=WoWChange
#    ${totalOnWoWchange}   Read Excel Cell    row_num=${totalRowIndexOnWoWChange}    col_num=${posOfColOnWoWChange}
#    ${totalOnWoWchange}   Evaluate  "%.2f" % ${totalOnWoWchange}
#    IF    ${totalOnWoWchange} != ${totalOnSG}
#         ${result}     Set Variable    ${False}
#         Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    ${table} Total    ${totalOnWoWchange}    ${totalOnSG}
#    END

#    #  Verify the OTHERS data
#    Switch Current Excel Document    doc_id=SG
#    ${othersOnSG}  Set Variable    0
#    FOR    ${rowIndexOnSG}    IN RANGE    6    ${numOfRowsOnSG}+1
#        ${oemGroupColOnSG}          Read Excel Cell    row_num=${rowIndexOnSG}    col_num=2
#        ${mainSalesRepColOnSG}      Read Excel Cell    row_num=${rowIndexOnSG}    col_num=3
#        IF    '${nameOfCol}' == 'Ships' or '${nameOfCol}' == 'Pre Q Ships'
#             ${dataColOnSG}           Read Excel Cell    row_num=${rowIndexOnSG}    col_num=${posOfRColOnSG}
#        ELSE IF  '${nameOfCol}' == 'Backlog'
#             ${dataColOnSG}           Read Excel Cell    row_num=${rowIndexOnSG}    col_num=${posOfBColOnSG}
#        ELSE IF  '${nameOfCol}' == 'LOS'
#             ${dataRColOnSG}           Read Excel Cell    row_num=${rowIndexOnSG}    col_num=${posOfRColOnSG}
#             ${dataBolOnSG}           Read Excel Cell    row_num=${rowIndexOnSG}    col_num=${posOfBColOnSG}
#             IF    '${dataRColOnSG}' == 'None'
#                  ${dataRColOnSG}   Set Variable    0
#             END
#             IF    '${dataBolOnSG}' == 'None'
#                  ${dataBolOnSG}    Set Variable    0
#             END
#             ${dataColOnSG}     Evaluate    ${dataRColOnSG}+${dataBolOnSG}
#        ELSE
#             Fail   The nameOfCol parameter ${nameOfCol} is invalid. Please contact with the Administrator for supporting
#        END
#
#        IF    '${dataColOnSG}' == 'None'
#               ${dataColOnSG}     Set Variable    0
#        END
#
#        IF    '${table}' == 'OEM East'
#             IF    '${mainSalesRepColOnSG}' in ${listOfSalesMemberInOEMEastTable}
#                IF    '${oemGroupColOnSG}' not in ${listOfOEMGroupShownInOEMEastTable}
#                   ${othersOnSG}     Evaluate    ${othersOnSG}+${dataColOnSG}
#                END
#             END
#        ELSE IF   '${table}' == 'OEM West'
#            IF    '${mainSalesRepColOnSG}' in ${listOfSalesMemberInOEMWestTable}
#                IF    '${oemGroupColOnSG}' not in ${listOfOEMGroupShownInOEMWestTable}
#                   ${othersOnSG}     Evaluate    ${othersOnSG}+${dataColOnSG}
#                END
#            END
#        ELSE
#            Fail    The table parameter ${table} is invalid. Please contact with the Administrator for supporting
#        END
#
#    END
#    ${othersOnSG}   Evaluate  "%.2f" % ${othersOnSG}
#    Switch Current Excel Document    doc_id=WoWChange
#    ${othersOnWoWChange}   Read Excel Cell    row_num=${othersRowIndexOnWoWChange}    col_num=${posOfColOnWoWChange}
#    ${othersOnWoWChange}   Evaluate  "%.2f" % ${othersOnWoWChange}
#    IF    ${othersOnWoWChange} != ${othersOnSG}
#         ${result}     Set Variable    ${False}
#         Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    ${table} OTHERS    ${othersOnWoWChange}    ${othersOnSG}
#    END
#
#    IF    '${result}' == '${False}'
#         Close All Excel Documents
#         Fail   The ${nameOfCol} data for the ${table} table between the WoW Change Report and SG Report is different
#    END
#    Close All Excel Documents

