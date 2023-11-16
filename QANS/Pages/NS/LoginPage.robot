*** Settings ***
Resource    ../CommonPage.robot

*** Variables ***
${txtEmail}        //*[@id='email']
${txtPassword}     //*[@id='password']
${btnLogin}        //*[@id='login-submit']
${txtLoginTitle}    //*[@id='uif43']
${txtAccountTitle}  //*[@class='tableTitle']
${txtVerificationCode}    //*[@id='uif51']
${chkTrustThisDeviceFor30Days}    //*[@id='uif67']
${btnSubmit}    //*[@id='uif71']
${imgSANDBOXIcon}   //*[@id='uif128']
${imgLogoIcon}      //*[@id='uif129']

*** Keywords ***
Login To NS With Account
    [Arguments]     ${account}
    Login To NS
    Should See Account Title    Choose account
    Choose Account    ${account}
    IF    '${account}' == 'PRODUCTION'
        Should See Login Title    Logging in to Virtium
    ELSE IF     '${account}' == 'SANDBOX4'
        Should See Login Title    Logging in to Virtium__SB4
    END
    Input Verification Code And Click Submit
    IF    '${account}' == 'PRODUCTION'
        Should See Account Is PRODUCTION
    ELSE IF     '${account}' == 'SANDBOX4'
        Should See Account Is SANDBOX
    END


Login To NS
    ${configFileObject}     Load Json From File    ${CONFIG_FILE}
    ${url}               Get Value From Json    ${configFileObject}    $.nsUrl
    ${url}               Set Variable       ${url}[0]
    ${email}             Get Value From Json    ${configFileObject}    $.accounts[1].email
    ${email}             Set Variable       ${email}[0]
    ${pass}              Get Value From Json    ${configFileObject}    $.accounts[1].password
    ${pass}              Set Variable       ${pass}[0]
    Go To    ${url}
    Wait Until Element Is Enabled    ${txtEmail}       ${TIMEOUT}
    Input Text    ${txtEmail}    ${email}
    Wait Until Element Is Enabled    ${txtPassword}    ${TIMEOUT}
    Input Text    ${txtPassword}    ${pass}
    Wait Until Element Is Enabled    ${btnLogin}       ${TIMEOUT}
    Click Element    ${btnLogin}

Choose Account
    [Arguments]     ${account}
    @{listOfAccountElements}    Get WebElements    //*[@class='listTable']/tbody/tr
    ${numOfAccounts}     Get Length    ${listOfAccountElements}

    FOR    ${accountIndex}    IN RANGE    2    ${numOfAccounts}+1
        ${company}      Get Text    //*[@class='listTable']/tbody/tr[${accountIndex}]/td[1]
        ${accountType}  Get Text    //*[@class='listTable']/tbody/tr[${accountIndex}]/td[2]
        IF    '${account}' == 'PRODUCTION'
            IF    '${company}' == 'Virtium' and '${accountType}' == 'PRODUCTION'
                 Wait Until Element Is Visible    //*[@class='listTable']/tbody/tr[${accountIndex}]/td[3]/a     ${TIMEOUT}
                 Click Element    //*[@class='listTable']/tbody/tr[${accountIndex}]/td[3]/a
                 Exit For Loop
            END
        ELSE IF     '${account}' == 'SANDBOX4'
            IF    '${company}' == 'Virtium__SB4' and '${accountType}' == 'SANDBOX'
                Wait Until Element Is Visible    //*[@class='listTable']/tbody/tr[${accountIndex}]/td[3]/a      ${TIMEOUT}
                Click Element    //*[@class='listTable']/tbody/tr[${accountIndex}]/td[3]/a
                Exit For Loop
            END
        ELSE
             Fail   The Account ${account} is invalid. Please contact with Admin!
        END
    END

Should See Login Title
    [Arguments]     ${title}
    Wait Until Element Is Visible    ${txtLoginTitle}       ${TIMEOUT}
    Element Text Should Be    ${txtLoginTitle}    ${title}

Should See Account Title
    [Arguments]     ${title}
    Wait Until Element Is Visible    ${txtAccountTitle}       ${TIMEOUT}
    Element Text Should Be    ${txtAccountTitle}    ${title}

Input Verification Code And Click Submit
    ${configFileObject}     Load Json From File    ${CONFIG_FILE}
    ${key}      Get Value From Json    ${configFileObject}    $.nsKey
    ${key}      Set Variable    ${key}[0]
    ${otp}      Generate Otp    ${key}
    Wait Until Element Is Enabled    ${txtVerificationCode}     ${TIMEOUT}
    Input Text    ${txtVerificationCode}    ${otp}
    Wait Until Element Is Enabled    ${chkTrustThisDeviceFor30Days}     ${TIMEOUT}
    Click Element    ${chkTrustThisDeviceFor30Days}
    Wait Until Element Is Enabled    ${btnSubmit}   ${TIMEOUT}
    Click Element    ${btnSubmit}
    
    







Should See Account Is SANDBOX
    Wait Until Element Is Visible    ${imgSANDBOXIcon}      ${TIMEOUT}
    Wait Until Element Is Visible    ${imgLogoIcon}     ${TIMEOUT}
    
Should See Account Is PRODUCTION
    Wait Until Element Is Visible    ${imgLogoIcon}     ${TIMEOUT}
    
