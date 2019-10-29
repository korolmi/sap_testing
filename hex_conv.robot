*** Test Cases ***
convert_smoke
  ${res} =  Convert To Hex  255
  Should Be Equal  ${res}  FF
convert_prefix
  ${res} =  Convert To Hex  255  prefix=0X
  Should Be Equal  ${res}  0XFF
convert_prefix_length
  ${res} =  Convert To Hex  255  prefix=0X  length=4
  Should Be Equal  ${res}  0X00FF
convert_prefix_lowercase
  ${res} =  Convert To Hex  255  prefix=0X  lowercase=yes
  Should Be Equal  ${res}  0Xff
use own keyword
  ${res} =  Do Smoke  255  FF
  Log  \nKey result is: ${res}  console=yes

*** Keywords ***
do smoke
  [Arguments]  ${value}  ${result}
  [Return]  ${res}
  ${res} =  Convert To Hex  ${value}
  Should Be Equal  ${res}  ${result}
