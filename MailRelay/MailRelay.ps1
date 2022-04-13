# SECTION 1 - VARIABLES
$MailFrom = "Thin Mints <ThinMints@thickmints.dev>"
$MailTo = "CarmelDeLites@thickmints.dev"
$MailCC = "Lemonades@thickmints.dev,Smores@thickmints.dev,"
$ReplyTo = "$MailTo" #Or set the reply to an email box you control
$ReturnPath = "$MailTo"
$MessageSubject = "Please Send a Crate of Girl Scout Cookies ASAP"
$MessageBody = "Carmel, please make sure to send a crate of Girl Scout cookies to the crew over at ThickMints.dev.Put it on the corporate account. Thanks."
$MessageSignature = "Thin"
$MessageName = "Thin Mints"
$MessageTitle = "Chief Cookie Officer"
$MessageAddress = "123 Cookie Way, Sugar City, NY 12345" 
$MessagePhone = "555-867-5309"
$MessageEmail = "ThinMints@thickmints.dev"
$MessageSensitivity = "Confidential"

# SECTION 2 - SMTP SERVER CONFIG
$Username = "user"
$Secret = ConvertTo-SecureString -String "password" -AsPlainText -Force
$Creds = New-Object System.Management.Automation.PSCredential($Username,$Secret)
$SmtpServer = "mail.thickmints.dev"
$SmtpPort = "25"

# SECTION 3 - BUILD E-MAIL MESSAGE
$Message = New-Object System.Net.Mail.MailMessage $MailFrom,$MailTo
# Optional: Here's where to attach images 
$Message.Attachments.Add("C:\Temp\MailSpoof\CompanyLogo.png")
$Message.Attachments.Add("C:\Temp\MailSpoof\Award1.png")
$Message.Attachments.Add("C:\Temp\MailSpoof\Award2.jpg")
$Message.Attachments.Add("C:\Temp\MailSpoof\Award3.png")
$Message.Attachments.Add("C:\Temp\MailSpoof\Award4.jpg")
$Message.Headers.Add("Return-Path",$ReturnPath)
$Message.CC.Add($MailCC)
$Message.ReplyTo=$ReplyTo
$Message.IsBodyHTML = $true
$Message.Subject = $MessageSubject
$Message.Body = @"
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40">
  <head>
    <meta http-equiv=Content-Type content="text/html; charset=us-ascii">
    <meta name=Generator content="Microsoft Word 15 (filtered medium)">
    <!--[if !mso]>
				<style>v\:* {behavior:url(#default#VML);}
o\:* {behavior:url(#default#VML);}
w\:* {behavior:url(#default#VML);}
.shape {behavior:url(#default#VML);}
</style>
				<![endif]-->
    <style>
      <!--
      /* Font Definitions */
      @font-face {
        font-family: "Cambria Math";
        panose-1: 2 4 5 3 5 4 6 3 2 4;
      }

      @font-face {
        font-family: Calibri;
        panose-1: 2 15 5 2 2 2 4 3 2 4;
      }

      /* Style Definitions */
      p.MsoNormal,
      li.MsoNormal,
      div.MsoNormal {
        margin: 0in;
        font-size: 11.0pt;
        font-family: "Calibri", sans-serif;
      }

      a:link,
      span.MsoHyperlink {
        mso-style-priority: 99;
        color: #0563C1;
        text-decoration: underline;
      }

      span.EmailStyle17 {
        mso-style-type: personal-compose;
        font-family: "Calibri", sans-serif;
        color: windowtext;
      }

      .MsoChpDefault {
        mso-style-type: export-only;
        font-family: "Calibri", sans-serif;
      }

      @page WordSection1 {
        size: 8.5in 11.0in;
        margin: 1.0in 1.0in 1.0in 1.0in;
      }

      div.WordSection1 {
        page: WordSection1;
      }
      -->
    </style>
    <!--[if gte mso 9]>
				<xml>
					<o:shapedefaults v:ext="edit" spidmax="1026" />
				</xml>
				<![endif]-->
    <!--[if gte mso 9]>
				<xml>
					<o:shapelayout v:ext="edit">
						<o:idmap v:ext="edit" data="1" />
					</o:shapelayout>
				</xml>
				<![endif]-->
  </head>
  <body lang=EN-US link="#0563C1" vlink="#954F72" style='word-wrap:break-word'>
    <p class=msipheader3e56b697 align="Right" style="margin:0">
	    <span style='font-size:10.0pt;font-family:Calibri;color:#000000'>$MessageSensitivity</span>
	</p>
	<br />
    <div class=WordSection1>
      <p class=MsoNormal>$MessageBody<o:p></o:p>
      </p>
      <p class=MsoNormal>
        <o:p>&nbsp;</o:p>
      </p>
      <p class=MsoNormal>$MessageSignature <o:p></o:p>
      </p>
      <p class=MsoNormal>
        <o:p>&nbsp;</o:p>
      </p>
      <p class=MsoNormal>-- <o:p></o:p>
      </p>
      <p class=MsoNormal>
        <o:p>&nbsp;</o:p>
      </p>
      <p class=MsoNormal>
        <b>
          <span style='font-size:12.0pt;color:#3EB489;mso-fareast-language:JA'>$MessageName<o:p></o:p>
          </span>
        </b>
      </p>
      <p class=MsoNormal>
        <span style='color:#333333;mso-fareast-language:JA'>$MessageTitle<o:p></o:p>
        </span>
      </p>
      <!-- *****PLEASE NOTE - IMAGES MUST MATCH THE IMAGE FILENAME*****  -->
      <p class=MsoNormal>
        <span style='color:#333333;mso-fareast-language:JA'>
          <img width=157 height=98 style='width:1.6354in;height:1.0208in' id="Picture_x0020_6" src="cid:CompanyLogo.png">
          <o:p></o:p>
        </span>
      </p>
      <p class=MsoNormal>
        <span style='color:#333333;mso-fareast-language:JA'>$MessageAddress<o:p></o:p>
        </span>
      </p>
      <p class=MsoNormal>
        <span style='color:#333333;mso-fareast-language:JA'>Mobile: $MessagePhone<o:p></o:p>
        </span>
      </p>
      <p class=MsoNormal>
        <span style='color:#333333;mso-fareast-language:JA'>
          <a href="mailto:$MessageEmail">$MessageEmail</a> | <a href="https://blog.thickmints.dev/">blog.thickmints.dev</a>
          <o:p></o:p>
        </span>
      </p>
      <p class=MsoNormal>
        <span style='color:#333333;mso-fareast-language:JA'>
          <o:p>&nbsp;</o:p>
        </span>
      </p>
      <!-- *****PLEASE NOTE - IMAGES MUST MATCH THE IMAGE FILENAMES*****  -->
      <p class=MsoNormal>
        <span style='color:#333333;mso-fareast-language:JA'>
          <img border=0 width=132 height=129 style='width:1.375in;height:1.3437in' id="Picture_x0020_8" src="cid:Award1.png" alt="Girl Scout Badge 1">&nbsp; <img border=0 width=132 height=129 style='width:1.375in;height:1.3437in' id="Picture_x0020_7" src="cid:Award2.jpg" alt="Girl Scout Badge 2">&nbsp;&nbsp; <img border=0 width=132 height=129 style='width:1.375in;height:1.3437in' id="Picture_x0020_9" src="cid:Award3.jpg" alt="Girl Scout Badge 3">&nbsp;&nbsp; <img border=0 width=132 height=129 style='width:1.375in;height:1.3437in' id="Picture_x0020_10" src="cid:Award4.jpg" alt="Girl Scout Badge 4">
          <o:p></o:p>
        </span>
      </p>
      <p class=MsoNormal>
        <span style='color:#333333;mso-fareast-language:JA'>Images and names are all copyrights of Girl Scouts of the United States of America. <o:p></o:p>
        </span>
      </p>
      <p class=MsoNormal>
        <span style='color:#333333;mso-fareast-language:JA'>&copy; 2016-2021 Girl Scouts of the United States of America. <o:p></o:p>
        </span>
      </p>
    </div>
  </body>
</html>
"@

# SECTION 4 - SENDING MESSAGE
$Smtp = New-Object Net.Mail.SmtpClient($SmtpServer,$SmtpPort)
$Smtp.EnableSsl = $true
$Smtp.Credentials = $creds
$Smtp.Send($Message)
