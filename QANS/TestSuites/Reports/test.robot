*** Settings ***
Library     DependencyLibrary

*** Test Cases ***
Testcase1
    Log To Console    Testcase1 Running
#    Fail    Testcase1 failed

Testcase2
    Log To Console    Testcase2 Running

Testcase3
    Depends On Test    Testcase1
    Depends On Test    Testcase2
    Log To Console    Testcase3 Running