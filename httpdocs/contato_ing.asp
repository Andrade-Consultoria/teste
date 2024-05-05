<%

 if request.form("envio") = "ok" then

    If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
        Dim recaptcha_secret, sendstring, objXML
        '' Secret key
        recaptcha_secret = "6LeAPbEeAAAAANAZqoR6rYTLqneqTTJt1rdJQtaB"

        sendstring = "https://www.google.com/recaptcha/api/siteverify?secret=" & recaptcha_secret & "&response=" & Request.form("g-recaptcha-response")

        Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP")
        objXML.Open "GET", sendstring, False

        objXML.Send
        
        
        if mid(objXML.responseText,16,4) = "true" then
        
        

	name = request.form("name")
	email = request.form("email")
	mensagem = request.form("message")
	phone = request.form("phone")
	assunto = request.form("subject")


 f0 = "Formulário Contato - Site" & "<br><br>"
 f1 = "Nome: "  & name & "<br>"
 f2 = "Email: "  & email  & "<br>"
 f3 = "Telefone: "  & phone  & "<br>"
 f4 = "Mensagem: "  & mensagem  & "<br>"
 corpo = f0 & f1 & f2 & f3 & f4
 
 

Dim objCDOSYSCon
''CRIA A INSTÂNCIA COM O OBJETO CDOSYS

Set objCDOSYSMail = Server.CreateObject("CDO.Message")

Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration")

	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.valenteadvogados.adv.br"

	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 587

	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2

	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30 

	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = "autentica@valenteadvogados.adv.br"
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "Novo2022@"

objCDOSYSCon.Fields.update 

Set objCDOSYSMail.Configuration = objCDOSYSCon

objCDOSYSMail.From = "secretaria@valenteadvogados.adv.br" 


objCDOSYSMail.To = "secretaria@valenteadvogados.adv.br"
''objCDOSYSMail.To = "ricardo@rcweb.com.br"

objCDOSYSMail.Subject = assunto & " - Email de quem enviou no corpo da mensagem"
objCDOSYSMail.HtmlBody = corpo


objCDOSYSMail.Send 

Set objCDOSYSMail = Nothing 
Set objCDOSYSCon = Nothing 

else

Response.Write "<script>alert('É obrigado marcar a caixa do ReCaptcha')</script>"


end if

end if

end if


%>
<%
menu5="current"   
%>

<!DOCTYPE HTML>
<!--[if lt IE 7]> <html class="no-js ie6" lang="en"> <![endif]-->
<!--[if IE 7]>    <html class="no-js ie7" lang="en"> <![endif]-->
<!--[if IE 8]>    <html class="no-js ie8" lang="en"> <![endif]-->
<!--[if gt IE 8]><!--> <html class="no-js" lang="en"> <!--<![endif]-->

	
<!-- Mirrored from demo.artureanec.com/html/saulsattorneys/style_1/contacts.html by HTTrack Website Copier/3.x [XR&CO'2014], Thu, 20 Dec 2018 11:49:48 GMT -->
<head>
		<title>Valente Lawyers</title>
		<meta http-equiv="Content-Language" content="pt-br">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252"> 

		<meta name="viewport" content="initial-scale=1.0, user-scalable=no" />
		<meta name="viewport" content="user-scalable=no, width=device-width, height=device-height, initial-scale=1, maximum-scale=1, minimum-scale=1, minimal-ui" />

		<meta name="theme-color" content="#C1AA81" />
		<meta name="msapplication-navbutton-color" content="#C1AA81" />
		<meta name="apple-mobile-web-app-status-bar-style" content="#C1AA81" />

		<!-- Favicons
		================================================== -->
    <link rel="apple-touch-icon" sizes="180x180" href="images/apple-touch-icon.png">
    <link rel="icon" type="image/png" href="images/favicon-32x32.png" sizes="32x32">
    

		<!-- CSS
		================================================== -->
		<link rel="stylesheet" href="css/style.min.css" type="text/css">

		<!-- / color -->
		<link class="colors_style" rel="stylesheet" href="css/color_scheme/color_1.css" type="text/css"/>

		<!-- / settings_box -->
		<link rel="stylesheet" href="settings_box/settings_box.css" type="text/css">

		<!-- / google font -->
		<link href='https://fonts.googleapis.com/css?family=Oxygen:400,700' rel='stylesheet' type='text/css'>
		<link href='https://fonts.googleapis.com/css?family=Rufina:400,700' rel='stylesheet' type='text/css'>

		<!-- Load jQuery
		================================================== -->
		<script type="text/javascript" src="js/device.min.js"></script>
		<script type="text/javascript" src="js/modernizr.custom.min.js"></script>
	</head>

	<body>

<script type="text/javascript" language="javascript">
function validateform(){
var captcha_response = grecaptcha.getResponse();
if(captcha_response.length == 0)
{
    alert('É obrigado marcar a caixa do ReCaptcha');
    return false;
}
else
{
    // Captcha is Passed
    return true;
}
}
// ]]></script>
  
<script src='https://www.google.com/recaptcha/api.js'></script>
		
		<!--#include file="menu_ing.asp"-->
		

		<section id="headline">
			<div class="container">
				<div class="row">
					<div class="col-xs-12 col-md-5">
						<p id="page-title" class="h1">Contact</p>
					</div>

					
				</div>
			</div>
		</section>
        
          <iframe src="https://www.google.com/maps/embed?pb=!1m18!1m12!1m3!1d1943.9526877463197!2d-38.45700709179269!3d-12.977903684014999!2m3!1f0!2f0!3f0!3m2!1i1024!2i768!4f13.1!3m3!1m2!1s0x7161b1c8afe2267%3A0x29d716d7c545636d!2sCondom%C3%ADnio%20Salvador%20Shopping%20Business!5e0!3m2!1spt-BR!2sbr!4v1646326434212!5m2!1spt-BR!2sbr" width="100%" height="450" frameborder="0" style="border:0" allowfullscreen></iframe>
        
        

		<main role="main">
			<section id="s-contact" class="section transparent">
				<div class="container">
					<div class="row">
						<div class="col-xs-12 col-md-5 col-lg-4">
							<div class="contact-item">
								<h2>Contact information</h2>

								<div class="contact-item_info">
									<article>
										<h4>Address</h4>

										Alameda Salvador, 1057<br>
                                        Sala 902, Torre América, Salvador Shopping Business<br>
                                        Caminho das Árvores<br>
                                    Salvador - BA, CEP 41.820-790
									</article>

									<article>
										<h4>Phone</h4>

										<strong><a href="tel:557135612360" style="text-decoration: none;">+55 71 3561-2360</a></strong>
									</article>

									<article>
										<h4>Email</h4>

										<a href="mailto:secretaria@valenteadvogados.adv.br" style="text-decoration: none;">secretaria@valenteadvogados.adv.br</a>
									</article>
								</div>
							</div>
						</div>

						<div class="col-xs-12 col-md-7 col-lg-7 col-lg-offset-1">
							<div class="contact-item">
								
								<% if request.form("envio")="ok" then %>
							<div class="alert alert-success" role="alert">
								
								<strong>Success!</strong> Your message has been sent, wait for our return.
							</div>
							<% end if %>
								
								<h2>Contact Form</h2>

								<form method="post" action="contato.asp" onsubmit="return validateform();">
									<label class="input-wrp">
										<input type="text" name="name" placeholder="Name" required/>
										<span></span>
									</label>
									
									<label class="input-wrp">
										<input type="text" name="subject" placeholder="Subject" />
										<span></span>
									</label>

									<label class="input-wrp">
										<input type="text" name="email" placeholder="E-Mail" required/>
										<span></span>
									</label>

									<label class="input-wrp">
										<input type="text" name="phone" placeholder="Phone" />
										<span></span>
									</label>

									<label class="input-wrp">
										<textarea placeholder="Message" name="message"></textarea>
										<span></span>
									</label>
									
									
										<div class="g-recaptcha" data-sitekey="6LeAPbEeAAAAAEPGbdF8SYbKlZ__zOat_fqyY75l"></div> 
										<input type="hidden" name="envio" value="ok">

									<button class="custom-btn small dark-color" type="submit" data-text="Submit"><span>Send</span></button>
								</form>
							</div>
						</div>
					</div>

					
				</div>
			</section>

			
		</main>
		
		

				<!--#include file="footer_ing.asp"-->

		<div id="btn-to-top-wrap">
			<a id="btn-to-top" class="circled" href="javascript:void(0);" data-visible-offset="1500"></a>
		</div>

		<script type="text/javascript" src="js/main.min.js"></script>

		<!-- / settings_box -->
		<script type="text/javascript" src="settings_box/settings_box.js"></script>
	</body>

<!-- Mirrored from demo.artureanec.com/html/saulsattorneys/style_1/contacts.html by HTTrack Website Copier/3.x [XR&CO'2014], Thu, 20 Dec 2018 11:49:49 GMT -->
</html>