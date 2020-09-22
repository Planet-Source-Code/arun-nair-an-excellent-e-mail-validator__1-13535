<div align="center">

## An Excellent E\-Mail Validator


</div>

### Description

Validating an email address is a big headache to both the web and stand alone app developers. Especially when the e-mail is used as the handle or key to your transactions, its structure becomes of importance and has to be properly entered. Unfortunately developing a module or an application to solve this problem is not without its own set of problems ....besides having to spend valuable time in combining and interpreting the validations.

As a developer, I faced this problem more than once and decided to find a permanent solution...and what better way than to develop a COM component!! and make it universally perusable!

Yes, you now have a powerful piece of component(with its source code) at your disposal. This COM not only comes as a boon(because of its inherent quality of being a COM and being multiusable), but also a great time saver . I'm sure you'll find this component worthy, as has many other developers in my organisation(yes!! its widely used in my company :)).

I have covered all aspects of email validations here ...so when you enter an email address for verification, you are sure it'll be verified!!

I have also added the DLL along with the ZIP :)
 
### More Info
 
An email address to be verified.

To use the COM follow the steps:

1.) Reference the COM through Project references(if you are using VB)

2.) If you are not using VB .for eg. within your ASP script, write the code

dim mvObj

set mvObj=Server.CreateObject("ArunUtils.MailValidator")

3.) Now refer the checkMailVal method within the object. and store the output in a var eg: result=mvObj.checkMailVal("n_arun@vsnl.com")

4.) If the email is structurely right, the return value is an empty string ...else the error description is returned as the string

Basic knowledge of Visual Basic and also usage of COM components

Whether the email supplied is valid or not!

Nothing!!


<span>             |<span>
---                |---
**Submitted On**   |2000-12-12 17:35:46
**By**             |[Arun Nair](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/arun-nair.md)
**Level**          |Intermediate
**User Rating**    |4.1 (29 globes from 7 users)
**Compatibility**  |VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[OLE/ COM/ DCOM/ Active\-X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/ole-com-dcom-active-x__1-29.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD1263612132000\.zip](https://github.com/Planet-Source-Code/arun-nair-an-excellent-e-mail-validator__1-13535/archive/master.zip)








