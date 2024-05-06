## Patterns which have issues
1.We need to have fix pattern like the start character should be same as end character
  Ex - #Prospect_ID# or {Prospect_ID}
2. Different Patterns in files which do not have same start and end values
  * \$\$Prospect_ID# 
  * OTP##\$\$ Enter OTP ( This should be excluded if #1 is in place)
    * Some special characters like circle are added - https://github.com/Pavs1605/ExcelParser/blob/main/img_1.png
  * <@PaidAmount> - Has start and end values but has an @ as well
  *  #shortcode - No ending character

  * \$\$Enter OTP \$\$OTP\$\$ => This string is a part of OTP##\$\$Enter OTP \$\$OTP\$\$, here \$\$ is available for enter OTP, and enter OTP can be considered as a token
  * OTP##\$\$Your One Time Password for adding \$\$BNF_NAME\$\$( Here the token can be considered as Your One Time Password for adding )
3. Url's
   * It's better to not include '/' as dynamic values character
      Ex -https://play.google.com/store/ => here store can be considered as a token
4. Some SMS templates column values are empty
5. Column name is template, SUBJECT, PARAMETER_VALUE, PARAM_VALUE, SMS_Text in few data files.
6. Exclusion of double quotes & single quotes for dynamic values, else the whole string would be considered as a token
   "Dear Customer, An amount of AED {Amount} has been withdrawn from your account {AccountNumber}, in {BranchName} {TermianlNumber} on Date {Date} Time {Time}."
7. 
