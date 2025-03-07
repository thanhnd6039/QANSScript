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
Write Test Result Of WoW Change Report To Excel
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


Get Row Index For Search Col
    [Arguments]     ${table}
    ${rowIndexForSearchCol}     Set Variable    0

    File Should Exist    path=${wowChangeFilePath}
    Open Excel Document    filename=${wowChangeFilePath}    doc_id=WoWChange
    ${numOfRows}    Get Number Of Rows In Excel    filePath=${wowChangeFilePath}
    FOR    ${rowIndex}    IN RANGE    1    ${numOfRows}+1
        ${oemGroupCol}      Read Excel Cell    row_num=${rowIndex}    col_num=1
        IF    '${oemGroupCol}' == '${table}'
             ${rowIndexForSearchCol}    Set Variable    ${rowIndex}
             BREAK
        END
    END

    Close Current Excel Document
    [Return]    ${rowIndexForSearchCol}

Get Position Of Col
    [Arguments]     ${table}    ${nameOfCol}
    ${pos}  Set Variable    0
    ${rowIndexForSearchCol}     Set Variable    0

    ${rowIndexForSearchCol}     Get Row Index For Search Col    ${table}
    ${pos}      Get Position Of Column    filePath=${wowChangeFilePath}    rowIndex=${rowIndexForSearchCol}    searchStr=${nameOfCol}

    [Return]    ${pos}


Get Start Row On WoW Change
    [Arguments]     ${table}

    ${startRow}     Set Variable    0
    ${posOfCol}     Get Position Of Col    table=${table}   nameOfCol=${table}

    File Should Exist    path=${wowChangeFilePath}
    Open Excel Document    filename=${wowChangeFilePath}    doc_id=WoWChange
    ${numOfRows}    Get Number Of Rows In Excel    filePath=${wowChangeFilePath}

    FOR    ${rowIndex}    IN RANGE    1    ${numOfRows}+1
        ${oemGroupCol}     Read Excel Cell    row_num=${rowIndex}    col_num=${posOfCol}
        IF    '${oemGroupCol}' == '${table}'
             ${startRow}    Evaluate    ${rowIndex}+1
             BREAK
        END
    END

    Close Current Excel Document
    [Return]    ${startRow}

Get End Row On WoW Change
    [Arguments]     ${table}

    ${endRow}   Set Variable    0
    ${count}    Set Variable    0

    ${posOfCol}     Get Position Of Col    table=${table}   nameOfCol=${table}

    File Should Exist    path=${wowChangeFilePath}
    Open Excel Document    filename=${wowChangeFilePath}    doc_id=WoWChange
    ${numOfRows}    Get Number Of Rows In Excel    filePath=${wowChangeFilePath}

    FOR    ${rowIndex}    IN RANGE    1    ${numOfRows}+1
        ${oemGroupCol}     Read Excel Cell    row_num=${rowIndex}    col_num=${posOfCol}
        IF    '${oemGroupCol}' == 'OTHERS'
             ${count}   Evaluate    ${count}+1
             IF    '${table}' == 'OEM East'
                  IF    '${count}' == '1'
                       ${endRow}    Evaluate    ${rowIndex}-1
                       BREAK
                  END
             ELSE IF     '${table}' == 'OEM West + Channel'
                  IF    '${count}' == '2'
                       ${endRow}    Evaluate    ${rowIndex}-1
                       BREAK
                  END
             END
        END
    END

    Close Current Excel Document
    [Return]    ${endRow}

Get Trans Type
    [Arguments]     ${nameOfCol}
    ${transType}    Set Variable    ${EMPTY}

    IF    '${nameOfCol}' == 'Ships' or '${nameOfCol}' == 'Pre Q Ships'
         ${transType}   Set Variable    R
    END
    [Return]    ${transType}

Check BGT, Ship, Backlog, LOS On WoW Change
    [Arguments]     ${table}    ${nameOfCol}    ${year}     ${quarter}

    ${result}       Set Variable    ${True}
    ${startRow}     Set Variable    0
    ${endRow}       Set Variable    0
    ${othersRow}    Set Variable    0
    ${totalRow}     Set Variable    0
    ${posOfCol}     Set Variable    0
    ${transType}    Set Variable    ${EMPTY}
    ${sumOfValueOfOEMGroup}     Set Variable    0

    IF    '${table}' != 'OEM East' and '${table}' != 'OEM West + Channel'
         Fail    The table parameter ${table} is invalid. Please contact with the Administrator for supporting
    END

    ${transType}    Get Trans Type    nameOfCol=${nameOfCol}
    ${startRow}     Get Start Row On WoW Change    table=${table}
    ${endRow}       Get End Row On WoW Change    table=${table}
    ${othersRow}    Evaluate    ${endRow}+1
    ${totalRow}     Evaluate    ${endRow}+2

    IF    '${nameOfCol}' == 'Pre Q Ships'
         ${posOfCol}    Set Variable    2
    ELSE
        ${posOfCol}     Get Position Of Col    table=${table}    nameOfCol=${nameOfCol}
    END

    ${listOfSalesMemberInOEMEastTable}       Get List Of Sales Member In OEM East Table
    ${listOfOEMGroupShownInOEMEastTable}     Get List Of OEM Group Shown In OEM East Table
    ${listOfSalesMemberInOEMWestTable}       Get List Of Sales Member In OEM West Table
    ${listOfOEMGroupShownInOEMWestTable}     Get List Of OEM Group Shown In OEM West Table

    File Should Exist      path=${wowChangeFilePath}
    Open Excel Document    filename=${wowChangeFilePath}           doc_id=WoWChange
    ${numOfRows}  Get Number Of Rows In Excel    filePath=${wowChangeFilePath}
   #   Verify the data for each OEM Group
    FOR    ${rowIndex}    IN RANGE    ${startRow}    ${endRow}+1
        ${oemGroupCol}          Read Excel Cell    row_num=${rowIndex}    col_num=1
        ${valueOnWoWChange}     Read Excel Cell    row_num=${rowIndex}    col_num=${posOfCol}
        ${valueOnSG}            Get Value By OEM Group On SG Report    oemGroup=${oemGroupCol}    transType=${transType}    year=${year}    quarter=${quarter}  attribute=AMOUNT
        ${valueOnWoWChange}      Evaluate  "%.2f" % ${valueOnWoWChange}
        ${valueOnSG}              Evaluate  "%.2f" % ${valueOnSG}
        ${sumOfValueOfOEMGroup}     Evaluate    ${sumOfValueOfOEMGroup}+${valueOnSG}
        IF    ${valueOnWoWChange} != ${valueOnSG}
              ${result}     Set Variable    ${False}
              Write Test Result Of WoW Change Report To Excel    item=${nameOfCol}    oemGroup=${oemGroupCol}    valueOnWoWChange=${valueOnWoWChange}    valueOnSG=${valueOnSG}
        END

    END
#    #   Verify the Total data
    ${totalOnSG}  Get Value By Main Sales Rep On SG Report    mainSalesRep=${listOfSalesMemberInOEMEastTable}    transType=${transType}    year=${year}    quarter=${quarter}    attribute=AMOUNT
    ${totalOnWoWchange}   Read Excel Cell    row_num=${totalRow}    col_num=${posOfCol}
    ${totalOnSG}   Evaluate  "%.2f" % ${totalOnSG}
    ${totalOnWoWchange}   Evaluate  "%.2f" % ${totalOnWoWchange}
    IF    ${totalOnWoWchange} != ${totalOnSG}
         ${result}     Set Variable    ${False}
         Write Test Result Of WoW Change Report To Excel    ${nameOfCol}    ${table} Total    ${totalOnWoWchange}    ${totalOnSG}
    END

    #  Verify the OTHERS data
    ${othersOnSG}   Evaluate    ${totalOnSG}-${sumOfValueOfOEMGroup}
    ${othersOnWoWChange}   Read Excel Cell    row_num=${othersRow}    col_num=${posOfCol}
    ${othersOnSG}   Evaluate  "%.2f" % ${othersOnSG}
    ${othersOnWoWChange}   Evaluate  "%.2f" % ${othersOnWoWChange}
    IF    ${othersOnWoWChange} != ${othersOnSG}
         ${result}     Set Variable    ${False}
         Write Test Result Of WoW Change Report To Excel    ${nameOfCol}    ${table} OTHERS    ${othersOnWoWChange}    ${othersOnSG}
    END

    IF    '${result}' == '${False}'
         Close Current Excel Document
         Fail   The ${nameOfCol} data for the ${table} table between the WoW Change Report and SG Report is different
    END
    Close Current Excel Document

