*** Settings ***
Resource    ../CommonPage.robot
Resource    SGPageV2.robot

*** Variables ***
${WOW_CHANGE_FILE_PATH}     ${OUTPUT_DIR}\\Wow Change [Current Week].xlsx
${POS_OEM_GROUP_COL_ON_WOW_CHANGE}   1


*** Keywords ***
Check BGT, Ship, Backlog On WoW Change
    [Arguments]     ${nameOftable}    ${nameOfCol}     ${transType}   ${attribute}   ${year}     ${quarter}
    ${result}       Set Variable    ${True}
    ${sumOfValueOfOEMGroup}     Set Variable    0

    ${listOfSalesMemberInOEMEastTable}       Get List Of Sales Member In OEM East Table
    ${listOfOEMGroupShownInOEMEastTable}     Get List Of OEM Group Shown In OEM East Table
    ${listOfSalesMemberInOEMWestTable}       Get List Of Sales Member In OEM West Table
    ${listOfOEMGroupShownInOEMWestTable}     Get List Of OEM Group Shown In OEM West Table

    ${tableOnWoWChange}     Create Table On WoW Change    nameOftable=${nameOftable}    nameOfCol=${nameOfCol}
    ${tableOnSG}            Create Table For SG Report    transType=${transType}    attribute=${attribute}    year=${year}    quarter=${quarter}
    #   Verify the data for each OEM Group
    FOR    ${rowOnWoWChange}    IN    @{tableOnWoWChange}
        ${oemGroupCol}          Set Variable    ${rowOnWoWChange[0]}
        ${valueOnWoWChange}     Set Variable    ${rowOnWoWChange[1]}
        IF    '${oemGroupCol}' == 'OTHERS' or '${oemGroupCol}' == 'Total'
             Continue For Loop
        END
        ${valueOnSG}    Get Value By OEM Group On SG Report     tableOnSG=${tableOnSG}    oemGroup=${oemGroupCol}
        ${sumOfValueOfOEMGroup}     Evaluate    ${sumOfValueOfOEMGroup}+${valueOnSG}
        ${valueOnWoWChange}      Evaluate  "%.2f" % ${valueOnWoWChange}
        ${valueOnSG}             Evaluate  "%.2f" % ${valueOnSG}
        IF    '${valueOnWoWChange}' != '${valueOnSG}'
             ${result}     Set Variable    ${False}
              Write Test Result Of WoW Change Report To Excel    item=${nameOfCol}    oemGroup=${oemGroupCol}    valueOnWoWChange=${valueOnWoWChange}    valueOnSG=${valueOnSG}
        END
    END
#    #   Verify the Total data
#    ${totalOnSG}    Set Variable    0
#    ${valueOnSG}    Set Variable    0
#    FOR    ${rawData}    IN    @{tableOnSG}
#        ${mainSalesRepColOnSG}  Set Variable    ${rawData[1]}
#        IF    '${table}' == 'OEM East'
#             IF    '${mainSalesRepColOnSG}' in ${listOfSalesMemberInOEMEastTable}
#                ${valueOnSG}    Set Variable    ${rawData[3]}
#                ${totalOnSG}    Evaluate    ${totalOnSG}+${valueOnSG}
#             END
#        ELSE
#            IF    '${mainSalesRepColOnSG}' in ${listOfSalesMemberInOEMWestTable}
#                ${valueOnSG}    Set Variable    ${rawData[3]}
#                ${totalOnSG}    Evaluate    ${totalOnSG}+${valueOnSG}
#            END
#        END
#    END
#    ${totalOnWoWchange}     Set Variable    0
#    FOR    ${rowOnWoWChange}    IN    @{tableOnWoWChange}
#        ${oemGroupCol}          Set Variable    ${rowOnWoWChange[0]}
#        IF    '${oemGroupCol}' == 'Total'
#             ${totalOnWoWchange}    Set Variable    ${rowOnWoWChange[1]}
#             BREAK
#        END
#    END
#    ${totalOnSG}          Evaluate  "%.2f" % ${totalOnSG}
#    ${totalOnWoWchange}   Evaluate  "%.2f" % ${totalOnWoWchange}
#    IF    ${totalOnWoWchange} != ${totalOnSG}
#         ${result}     Set Variable    ${False}
#         Write Test Result Of WoW Change Report To Excel    ${nameOfCol}    ${table} Total    ${totalOnWoWchange}    ${totalOnSG}
#    END
#    #  Verify the OTHERS data
#    ${othersOnSG}   Evaluate    ${totalOnSG}-${sumOfValueOfOEMGroup}
#    ${othersOnWoWChange}     Set Variable    0
#    FOR    ${rowOnWoWChange}    IN    @{tableOnWoWChange}
#        ${oemGroupCol}          Set Variable    ${rowOnWoWChange[0]}
#        IF    '${oemGroupCol}' == 'OTHERS'
#             ${othersOnWoWChange}    Set Variable    ${rowOnWoWChange[1]}
#             BREAK
#        END
#    END
#    ${othersOnSG}          Evaluate  "%.2f" % ${othersOnSG}
#    ${othersOnWoWChange}   Evaluate  "%.2f" % ${othersOnWoWChange}
#    IF    ${othersOnWoWChange} != ${othersOnSG}
#         ${result}     Set Variable    ${False}
#         Write Test Result Of WoW Change Report To Excel    ${nameOfCol}    ${table} OTHERS    ${othersOnWoWChange}    ${othersOnSG}
#    END
#
#    IF    '${result}' == '${False}'
#         Fail   The ${nameOfCol} data for the ${table} table is different between the WoW Change Report and SG Report
#    END


Get Row Index For Search Col
    [Arguments]     ${nameOftable}
    ${rowIndexForSearchCol}     Set Variable    0

    File Should Exist    path=${WOW_CHANGE_FILE_PATH}
    Open Excel Document    filename=${WOW_CHANGE_FILE_PATH}    doc_id=WoWChange
    ${numOfRows}    Get Number Of Rows In Excel    filePath=${WOW_CHANGE_FILE_PATH}
    FOR    ${rowIndex}    IN RANGE    1    ${numOfRows}+1
        ${oemGroupCol}      Read Excel Cell    row_num=${rowIndex}    col_num=1
        IF    '${oemGroupCol}' == '${nameOftable}'
             ${rowIndexForSearchCol}    Set Variable    ${rowIndex}
             BREAK
        END
    END

    Close Current Excel Document
    [Return]    ${rowIndexForSearchCol}

Get Position Of Column On WoW Change
    [Arguments]     ${nameOftable}    ${nameOfCol}
    ${pos}  Set Variable    0
    ${rowIndexForSearchCol}     Set Variable    0

    ${rowIndexForSearchCol}     Get Row Index For Search Col    nameOftable=${nameOftable}
    ${pos}      Get Position Of Column    filePath=${WOW_CHANGE_FILE_PATH}    rowIndex=${rowIndexForSearchCol}    searchStr=${nameOfCol}

    [Return]    ${pos}

Get Start Row On WoW Change
    [Arguments]     ${nameOftable}

    ${startRow}     Set Variable    0
    ${posOfCol}     Get Position Of Column On WoW Change    nameOftable=${nameOftable}   nameOfCol=${nameOftable}

    File Should Exist    path=${WOW_CHANGE_FILE_PATH}
    Open Excel Document    filename=${WOW_CHANGE_FILE_PATH}    doc_id=WoWChange
    ${numOfRows}    Get Number Of Rows In Excel    filePath=${WOW_CHANGE_FILE_PATH}

    FOR    ${rowIndex}    IN RANGE    1    ${numOfRows}+1
        ${oemGroupCol}     Read Excel Cell    row_num=${rowIndex}    col_num=${posOfCol}
        IF    '${oemGroupCol}' == '${nameOftable}'
             ${startRow}    Evaluate    ${rowIndex}+1
             BREAK
        END
    END

    Close Current Excel Document
    [Return]    ${startRow}

Get End Row On WoW Change
    [Arguments]     ${nameOftable}

    ${endRow}   Set Variable    0
    ${count}    Set Variable    0

    ${posOfCol}     Get Position Of Column On WoW Change    table=${nameOftable}   nameOfCol=${nameOftable}

    File Should Exist      path=${WOW_CHANGE_FILE_PATH}
    Open Excel Document    filename=${WOW_CHANGE_FILE_PATH}    doc_id=WoWChange
    ${numOfRows}    Get Number Of Rows In Excel    filePath=${WOW_CHANGE_FILE_PATH}

    FOR    ${rowIndex}    IN RANGE    1    ${numOfRows}+1
        ${oemGroupCol}     Read Excel Cell    row_num=${rowIndex}    col_num=${posOfCol}
        IF    '${oemGroupCol}' == 'OTHERS'
             ${count}   Evaluate    ${count}+1
             IF    '${nameOftable}' == 'OEM East'
                  IF    '${count}' == '1'
                       ${endRow}    Evaluate    ${rowIndex}-1
                       BREAK
                  END
             ELSE IF     '${nameOftable}' == 'OEM West + Channel'
                  IF    '${count}' == '2'
                       ${endRow}    Evaluate    ${rowIndex}-1
                       BREAK
                  END
             END
        END
    END

    Close Current Excel Document
    [Return]    ${endRow}

Create Table On WoW Change
    [Arguments]     ${nameOftable}    ${nameOfCol}

    @{table}     Create List
    ${startRow}     Set Variable    0
    ${endRow}       Set Variable    0
    ${othersRow}    Set Variable    0
    ${totalRow}     Set Variable    0

    IF    '${nameOftable}' != 'OEM East' and '${nameOftable}' != 'OEM West + Channel'
         Fail    The nameOftable parameter ${nameOftable} is invalid
    END

    ${startRow}     Get Start Row On WoW Change    nameOftable=${nameOftable}
    ${endRow}       Get End Row On WoW Change      nameOftable=${nameOftable}
    ${othersRow}    Evaluate    ${endRow}+1
    ${totalRow}     Evaluate    ${endRow}+2

    IF    '${nameOfCol}' == 'Pre Q Ships'
         ${posOfValueCol}    Set Variable    2
    ELSE IF    '${nameOfCol}' == 'Current Q Budget'
         ${posOfValueCol}    Set Variable    3
    ELSE IF    '${nameOfCol}' == 'LW Commit'
         ${posOfValueCol}    Set Variable    4
    ELSE IF  '${nameOfCol}' == 'TW Commit'
         ${posOfValueCol}    Set Variable    5
    ELSE IF  '${nameOfCol}' == 'WoW Of Ships'
         ${posOfValueCol}    Set Variable    7
    ELSE IF  '${nameOfCol}' == 'WoW Of LOS'
         ${posOfValueCol}    Set Variable    10
    ELSE
        ${posOfValueCol}     Get Position Of Column On WoW Change    nameOftable=${nameOftable}    nameOfCol=${nameOfCol}
    END

    File Should Exist      path=${WOW_CHANGE_FILE_PATH}
    Open Excel Document    filename=${WOW_CHANGE_FILE_PATH}           doc_id=WoWChange
    FOR    ${rowIndex}    IN RANGE    ${startRow}    ${totalRow}+1
        ${oemGroupColOnWoWChange}          Read Excel Cell    row_num=${rowIndex}    col_num=${POS_OEM_GROUP_COL_ON_WOW_CHANGE}
        ${valueColOnWoWChange}             Read Excel Cell    row_num=${rowIndex}    col_num=${posOfValueCol}
        ${rowOnTable}   Create List
        ...             ${oemGroupColOnWoWChange}
        ...             ${valueColOnWoWChange}
        Append To List    ${table}   ${rowOnTable}
    END

    Close Current Excel Document
    [Return]    ${table}

Get List Of Sales Member In OEM East Table
    @{listOfSalesMember}    Create List
    Append To List    ${listOfSalesMember}      Chris Seitz
    Append To List    ${listOfSalesMember}      Daniel Schmidt
    Append To List    ${listOfSalesMember}      Eli Tiomkin
    Append To List    ${listOfSalesMember}      Michael Pauser

    [Return]    ${listOfSalesMember}

Get List Of OEM Group Shown In OEM East Table
    @{listOfOEMGroup}   Create List
    Append To List    ${listOfOEMGroup}      NVIDIA/MELLANOX
    Append To List    ${listOfOEMGroup}      NOKIA/ALCATEL LUCENT WORLDWIDE
    Append To List    ${listOfOEMGroup}      CURTISS WRIGHT GROUP
    Append To List    ${listOfOEMGroup}      JUNIPER NETWORKS
    Append To List    ${listOfOEMGroup}      ERICSSON WORLDWIDE

    [Return]    ${listOfOEMGroup}

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

Get List Of OEM Group Shown In OEM West Table
    @{listOfOEMGroup}   Create List
    Append To List    ${listOfOEMGroup}      PALO ALTO NETWORKS
    Append To List    ${listOfOEMGroup}      ARISTA
    Append To List    ${listOfOEMGroup}      SCHWEITZER ENGINEERING LABORATORIES (SEL)
    Append To List    ${listOfOEMGroup}      PANASONIC AVIONICS
    Append To List    ${listOfOEMGroup}      ZTE KANGXUN TELECOM CO. LTD.
    Append To List    ${listOfOEMGroup}      NATIONAL INSTRUMENTS

    [Return]    ${listOfOEMGroup}