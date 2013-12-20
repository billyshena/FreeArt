<%session.lcid=1036%>
<!--#include file="base.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta http-equiv="content-language" content="FR">
<meta name="identifier-url" content="http://www.voyance-immediate.fr">
<meta name="Keywords" lang="fr" content="voyance par t�l�phone, Voyance, telephone , tarot, astrologie, voyante, t�l, voyants, consultation, avenir, tel, t�l�phone, astrodirect, astro direct, voyance par tarot, consultation voyance telephone, voyante s�rieuse, voyance au t�l�phone, connaitre son avenir, astrologie par telephone, voyante t�l�phone, voyance en directe, voyance tel, voyance t�l�phone, voyance pas cher, t�l�phone voyance, voyance par t�l, voyance par t�l�phone sans attente, voyance par telephone serieuse, voyance par telephone sans cb, voyant paris, minutes gratuites, voyance en ligne, voyance amour">
<meta name="dc.keywords" content="voyance, voyance immediate, voyance gratuite, medium pure, astrologie, astrologie gratuite, horoscope, horoscope gratuit, tarot, tarot gratuit, magie blanche, magie blanche gratuite, voyance gratuite en ligne, voyance gratuite en direct, voyance gratuite par t�l�phone, voyance gratuite par email, voyance gratuite par telephone, voyance gratuite par tel, voyance gratuite par mail, voyance gratuite par webcam, voyance par webcam, rune, runes, guide voyance, astrologie couples, voyance gratuite immediate, voyance populaire, voyance amour">
<meta name="Description" content="Voyance immediate : voyance par t�l�phone ou voyance en ligne, sans attente et en tarif local. Consultation voyance en priv�e avec de vrais m�diums de qualit�. Voyance imm�diate, 1er site de voyance en ligne solidaire. Domaine voyance amour et voyance sur le professionnel.">
<meta name="author" content="Voyance Immediate">
<meta name="Subject"  content="voyance">
<meta name="copyright" content="www.voyance-immediate.fr">
<meta name="robots" content="index, follow, noydir, noodp">
<meta http-equiv="cache-control" content="no-cache">
<meta http-equiv="pragma" content="no-cache">
<link rel="shortcut icon" href="favicon.ico">
<script src="js/jquery.js"></script>
<link href="site.css" rel="stylesheet" type="text/css" media="all">
<title>VOYANCE IMMEDIATE : Le site de la voyance en ligne, par t&eacute;l&eacute;phone ou par CHAT.</title>
<script type="text/javascript">

  var _gaq = _gaq || [];
  _gaq.push(['_setAccount', 'UA-1905162-2']);
  _gaq.push(['_trackPageview']);

  (function() {
    var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
	ga.src = ('https:' == document.location.protocol ? 'https://' : 'http://') + 'stats.g.doubleclick.net/dc.js';
    var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
  })();

</script>
</head>
<body>
<table width="960" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="591" align="center" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="70" colspan="2">&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td width="20%" align="center" valign="middle"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td align="right" valign="middle"><img src="images/conception/espace_client.jpg" width="23" height="15" alt="espace client"></td>
            <td align="center" valign="middle"><a href="espace_client/default.asp" class="lien_gris_12">Acc&egrave;s membre</a></td>
          </tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="62" align="center" valign="middle"><table width="960" border="0" cellpadding="0" cellspacing="0" id="contenu_partie_haute">
      <tr>
        <td align="center" valign="bottom">
        <!--#include file="haut.asp" -->
        </td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td align="center" valign="middle"><table width="960" border="0" cellpadding="0" cellspacing="0" id="contenu_partie_centre">
      <tr>
        <td width="40%" align="center" valign="top"><table width="98%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td align="center" valign="middle"><br>
              <table width="100%" border="0" cellpadding="0" cellspacing="0" id="contenu_menu_edito">
                <tr>
                  <td align="center" valign="middle"><br>
                    <br>
                    <table width="90%" border="0" cellspacing="8" cellpadding="0">
                      <tr>
                      <td align="left" valign="top"><div align="justify"><span class="mauve_12_italique">
                        <%
			  		  
			  intRandom = 59
			  
			  SQL = "SELECT * FROM reseau where id_prod = " & intRandom & ""
			  Set rs = Server.CreateObject("ADODB.Recordset")
			  rs.Open sql, conn, 3, 3
			  '
			  do while not rs.eof
			  edito = rs.fields("edito")
			  rs.movenext
			  loop
              
			  rs.close 
              set rs=nothing
			  
			  response.write edito
			  
			  %>
                      </span></div></td>
                      </tr>
                </table></td>
                  </tr>
                </table>
              <br>
              <table width="98%" border="0" cellpadding="0" cellspacing="0" id="contenu_menu_blog">
                <tr>
                  <td align="center" valign="middle"><br>
                    <br>
                    <table width="85%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                      <td width="60" align="left" valign="top"><a href="http://www.voyance-immediate.fr/blog/default.asp" class="ex3">
                        <%        
    
	la_une = "oui"
	
	sql = "SELECT * FROM blog WHERE une = '" & la_une & "'"
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sql, conn, 3, 3				
	do while not rs.eof
    titre = rs.fields("titre")
	auteur = rs.fields("auteur")
	datos = rs.fields("datos")
	sujet = rs.fields("id_prod")
 	rs.MoveNext				
	Loop		  
	rs.close
	set rs = nothing
	
	titros = Ucase(titre)
	nom_image = auteur & ".jpg"
	image = "http://www.voyance-immediate.fr/admin/data/" & nom_image
	
response.write ("<img src='" & image & "' border = ""0""/>")
%>
                      </a></td>
                      <td align="left" valign="top" class="mauve_12_italique"><div align="justify">
          <%		  
		  response.write "<u>sujet du jour :</u><br><br><i><a href=""http://www.voyance-immediate.fr/blog/default.asp?sujet=" & titre & """>" & titros & "</a></i>       (par " & auteur & ")"	
		  %>
                      </div></td>
                    </tr>
                </table></td>
                </tr>
            </table>
              <br>
              <table width="100%" border="0" cellpadding="0" cellspacing="0" id="contenu_menu_mail">
                <tr>
                  <td align="center" valign="middle"><form action="voyance_par_mail.asp" method="post">
                    <table width="85%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="50%" height="40">&nbsp;</td>
                        <td width="50%" height="40" colspan="2">&nbsp;</td>
                        </tr>
                      <tr>
                        <td align="left" valign="middle" class="mauve_12_italique">CHOIX DU VOYANT &gt;</td>
                        <td align="center" valign="middle"><select name="voyant" size="1" id="voyant">
                          <% 
Set rs = conn.Execute("SELECT * FROM voyant order by voyant") 
%>
                          <% 
rs.movefirst
Do While Not rs.EOF 
if rs("promo_04") = "Question Mail" Then
if rs("etat") <>"OFF" Then
%>
                          <option value="<%= rs("voyant") %>"><%= rs("voyant") %></option>
                          <%
end if
end if
rs.MoveNext
Loop 
'
rs.close
set rs=nothing
'
%>
                          </select></td>
                        <td align="center" valign="middle"><label>
                          <input type="submit" name="button2" id="button2" value="ok">
                          </label></td>
                        </tr>
                      <tr>
                        <td height="50" align="center" valign="middle" class="mauve_12_italique">(r&eacute;ponse sous 24h00)</td>
                        <td height="50" colspan="2" valign="bottom"><input type="submit" name="button" id="button" value="Envoyer"></td>
                        </tr>
                      </table>
                    </form></td>
                  </tr>
                </table>
              <br>
              <table width="90%" border="0" cellpadding="0" cellspacing="0" id="contenu_menu_rapide">
                <tr>
                  <td align="center" valign="middle"><form action="voyance_oui_non.asp" method="post">
                    <br>
                    <br>
                    <br>
                    <br>
                    <table width="85%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="39%" height="30" align="left" valign="top" class="mauve_10_italique">PRENOM</td>
                        <td width="61%" height="30" align="left" valign="top"><label for="naissance"></label>
                          <input name="prenom" type="text" id="textfield3" size="20"></td>
                      </tr>
                      <tr>
                        <td height="30" align="left" valign="top" class="mauve_10_italique">E-MAIL</td>
                        <td height="30" align="left" valign="top"><input name="email" type="text" id="textfield" size="20"></td>
                      </tr>
                      <tr>
                        <td height="30" align="left" valign="top" class="mauve_10_italique">DATE DE NAISSANCE</td>
                        <td height="30" align="left" valign="top" class="mauve_12_italique"><input name="naissance" type="text" id="textfield2" size="8" maxlength="8"> 
                          (jjmmaaaa)</td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" class="mauve_10_italique">QUESTION</td>
                        <td align="left" valign="top"><textarea name="question" cols="20" rows="3" id="question"></textarea></td>
                      </tr>
                      <tr>
                        <td height="40" align="center" valign="middle" class="mauve_10_italique">CODE VALIDATION<br>
                        <%
									'CREATION D UN PASSWORD

									session("password") = makePassword(3)
									session("password2") = "REF-" & session("password") & "-" & date

									function makePassword(byVal maxLen)

									Dim strNewPass
									Dim whatsNext, upper, lower, intCounter
									Randomize

									For intCounter = 1 To maxLen
        							upper = 90
        							lower = 65
        							strNewPass = strNewPass & Chr(Int((upper - lower + 1) * Rnd + lower))
									Next
        							makePassword = strNewPass	
									makePassword = lCase(makePassword)

									end function
									
									response.write "<BR><span class=""mauve_16_bold"">" &  session("password") & "</span>"

						%>
                        </td>
                        <td height="40" align="left" valign="middle"><span class="mauve_10_italique">
                          <input name="verif" type="text" id="verif" size="3" maxlength="3">
                        (inscrivez le code de gauche)</span></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top">&nbsp;</td>
                        <td align="left" valign="top"><input type="submit" name="button3" id="button3" value="Envoyer"></td>
                      </tr>
              </table>
                    <br>
                  </form></td>
                  </tr>
                </table>
              <br>
              <table width="90%" border="0" cellpadding="0" cellspacing="0" id="contenu_menu_temoignage">
                <tr>
                  <td align="center" valign="middle"><table width="90%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="5%">&nbsp;</td>
                      <td align="left" valign="middle">&nbsp;</td>
                      <td width="6%">&nbsp;</td>
                      </tr>
                    <tr>
                      <td width="5%">&nbsp;</td>
                      <td align="left" valign="middle" class="mauve_12_italique"><div align="justify" class="mauve_10_italique">
                        <%
			  x = 1
			  
			  SQL = "SELECT * FROM commentaire"
			  Set rs = Server.CreateObject("ADODB.Recordset")
			  rs.Open sql, conn, 3, 3
			  '
			  do while not rs.eof
			  rs.movenext
			  x = x + 1
			  loop
              
			  rs.close 
              set rs=nothing
			  
			  Randomize
			  intLength = x
			  intRandom = CInt((Rnd * 1000)Mod intLength) + 1
			  		  
			  SQL = "SELECT * FROM commentaire where id_prod = " & intRandom & ""
			  Set rs = Server.CreateObject("ADODB.Recordset")
			  rs.Open sql, conn, 3, 3
			  '
			  do while not rs.eof
			  commentaire = rs.fields("commentaire")
			  auteur = rs.fields("prenom")
			  voyante = rs.fields("voyante")
			  ville = rs.fields("ville")
			  rs.movenext
			  loop
              
			  rs.close 
              set rs=nothing
			  
			  response.write "<b>" & Ucase(auteur) & " </b>, de " & ville & " pour " & Ucase(voyante) & " : " & "<BR><BR>"
			  response.write commentaire & "<BR>"
			  
			  %>
                        </div></td>
                      <td width="6%">&nbsp;</td>
                      </tr>
                    <tr>
                      <td width="5%">&nbsp;</td>
                      <td height="20" align="center" valign="bottom"><a href="temoignage.asp" class="lien_gris_fonce_10">voir l'ensemble de nos commentaires</a></td>
                      <td width="6%">&nbsp;</td>
                      </tr>
                    </table></td>
                  </tr>
                </table>
              <br>
              <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0" id="contenu_menu_reseaux">
                <tr>
                  <td align="center" valign="middle"><table width="98%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="50%">&nbsp;</td>
                      <td width="50%">&nbsp;</td>
                    </tr>
                    <tr>
                      <td width="50%">&nbsp;</td>
                      <td width="50%" height="180" align="center" valign="middle"><br>
                        <br>
                        <br>
                        <br>
                        <br>
                        <table width="98%" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                          <td height="40" align="center" valign="middle"><a href="https://www.facebook.com/voyancesolidaire" target="_blank"><img src="images/conception/facebook.png" alt="voyance imm&eacute;diate sur facebook" width="59" height="24" border="0"></a></td>
                        </tr>
                        <tr>
                          <td height="40" align="center" valign="middle"><a href="http://www.youtube.com/channel/UCH4CyZQE2BEM1Pqwg_3FAzQ" target="_blank"><img src="images/conception/logo_youtube.jpg" alt="voyance imm&eacute;diate sur youtube" width="53" height="23" border="0"></a></td>
                        </tr>
                        <tr>
                          <td height="40" align="center" valign="middle"><a href="https://twitter.com/lignevoyance" target="_blank"><img src="images/conception/twitter.jpg" alt="voyance imm&eacute;diate sur twitter" width="16" height="16" border="0"></a></td>
                        </tr>
              </table></td>
                    </tr>
                  </table></td>
                </tr>
              </table>
              <br></td>
            </tr>
          </table>
          <br>
          <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td width="3%">&nbsp;</td>
              <td align="center" valign="middle">
                <a class="twitter-timeline"  href="https://twitter.com/voyanceimmediat"  data-widget-id="388727116683939840">Tweets de @voyanceimmediat</a>
  <script>!function(d,s,id){var js,fjs=d.getElementsByTagName(s)[0],p=/^http:/.test(d.location)?'http':'https';if(!d.getElementById(id)){js=d.createElement(s);js.id=id;js.src=p+"://platform.twitter.com/widgets.js";fjs.parentNode.insertBefore(js,fjs);}}(document,"script","twitter-wjs");</script>
                </td>
              </tr>
            </table>
          <br></td>
        <td width="60%" align="left" valign="top"><br>
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td align="left" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" id="contenu_menu_post_it">
              <tr>
                <td align="center" valign="middle"><table width="98%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="4%"><br>                      <br></td>
                    <td width="50%" align="center" valign="bottom"><br>
                      <br>
                      <table width="100%" border="0" cellspacing="4" cellpadding="0">
                      <tr>
                        <td height="30" align="left" valign="middle" class="mauve_12_italique"><div align="justify">- De v&eacute;ritables professionnels confirm&eacute;s depuis plusieurs ann&eacute;es et situ&eacute;s en France.</div></td>
                      </tr>
                      <tr>
                        <td height="30" align="left" valign="middle" class="mauve_12_italique"><div align="justify">- Un site de voyance en ligne totalement ind&eacute;pendant, non soumis aux Groupes Audiotels.</div></td>
                      </tr>
                      <tr>
                        <td height="30" align="left" valign="middle" class="mauve_12_italique"><div align="justify">- V&eacute;ritable tarification &agrave; 1 euro la minute, tarif local et sans surtaxe de prix. (ex. 12mn = 12&euro; ou 22mn = 22&euro; etc. ...)</div></td>
                      </tr>
                      <tr>
                        <td height="30" align="left" valign="middle" class="mauve_12_italique"><div align="justify">- Un site qui ne cautionne pas l'audiotel et qui n'en proposera jamais.</div></td>
                      </tr>
                      <tr>
                        <td height="30" align="left" valign="middle" class="mauve_12_italique"><div align="justify">- Le site de la voyance &quot;solidaire&quot;, avec une tarification qui n'a jamais augment&eacute;e depuis sa cr&eacute;ation.</div></td>
                      </tr>
                    </table></td>
                    <td width="46%" align="center" valign="middle">
                                 <param name="quality" value="high">
            <param name="wmode" value="opaque">
            <param name="scale" value="noscale">
            <param name="salign" value="lt">
            <param name="FlashVars" value="&amp;MM_ComponentVersion=1&amp;skinName=Clear_Skin_1&amp;streamName=video/pr�sentation5;autoPlay=false&amp;autoRewind=true" />
<param name="quality" value="high">
              <param name="wmode" value="opaque">
              <param name="scale" value="noscale">
              <param name="salign" value="lt">
              <param name="FlashVars" value="&amp;MM_ComponentVersion=1&amp;skinName=Clear_Skin_1&amp;streamName=video/pr�sentation5;autoPlay=false&amp;autoRewind=true" />
              <br>
              <embed src="FLVPlayer_Progressive.swf" flashvars="&MM_ComponentVersion=1&skinName=Clear_Skin_1&streamName=video/pr�sentation5&autoPlay=false&autoRewind=true" quality="high" wmode="transparent" scale="noscale" width="232" height="132" name="FLVPlayer" salign="lt" type="application/x-shockwave-flash" pluginspage="http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash"></embed>

                    </td>
                    </tr>
                </table></td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td align="center" valign="middle"><img src="images/conception/notre_equipe_voyance.jpg" width="517" height="68" alt="voyance en ligne - notre &eacute;quipe de voyants et m&eacute;diums"></td>
          </tr>
          <tr>
            <td align="left" valign="top"><table width="98%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="20" align="left" valign="middle" class="mauve_10_italique"><img src="images/rubriques/bouton_vert.jpg" width="18" height="17" title = "voyant disponible"></td>
                <td align="left" valign="middle" class="mauve_10_italique">voyant disponible</td>
                <td width="20" align="left" valign="middle"><img src="images/rubriques/bouton_rouge.jpg" width="18" height="17" title = "voyant non pr&eacute;sent"></td>
                <td align="left" valign="middle" class="mauve_10_italique">voyant non pr&eacute;sent</td>
                <td width="20" align="left" valign="middle"><img src="images/rubriques/bouton_bleu.jpg" width="18" height="17" title = "voyant en cours de consultation"></td>
                <td align="left" valign="middle" class="mauve_10_italique">voyant en consultation</td>
                <td width="20" align="left" valign="middle"><img src="images/rubriques/bouton_gris.jpg" width="18" height="17" title = "voyant temporairement indisponible"></td>
                <td align="left" valign="middle" class="mauve_10_italique">voyant en pause.</td>
                </tr>
            </table>
              <br></td>
          </tr>
          <tr>
            <td align="left" valign="top"><%
 
sql = "SELECT * from voyant where etat <> '" & "OFF" & "' order by dispo desc"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql, conn, 3, 3

i= 1
nbchamp = 2

response.write "<table width=""90%"" border=""0"" cellspacing=""5"" cellpadding=""5"">" ' TABLEAU PRINCIPAL

do while not rs.eof ' loop
image = "http://www.voyance-immediate.fr/admin/data/" & rs.fields("visuel2")

couleur = "http://www.voyance-immediate.fr/admin/data/bouton_vert.jpg"
couleur2 = "http://www.voyance-immediate.fr/admin/data/bouton_rouge.jpg"
couleur3 = "http://www.voyance-immediate.fr/admin/data/bouton_gris.jpg"
couleur4 = "http://www.voyance-immediate.fr/admin/data/bouton_bleu.jpg"

symbol_01 = "http://www.voyance-immediate.fr/admin/data/spe_chat1.png"
symbol_02 = "http://www.voyance-immediate.fr/admin/data/spe_tel1.png"
symbol_03 = "http://www.voyance-immediate.fr/admin/data/spe_webcam1.png"

media_url = "voyance_telephone.asp?voyant=" & rs.fields("voyant")
voyant = rs.fields("voyant")
conges = rs.fields("conges")


 if i= 1 then ' le compteur commence
  
  response.write("<tr>")
  response.write ("<td td align=""left"" valign=""top"">")
  
  response.write ("<table width=""280"" border=""0"" cellpadding=""2"" cellspacing=""2""  id=""contenu_sous_fiche"">")
  response.write ("<td td align=""left"" valign=""top"">")
  
  ' cr&eacute;ation du tableau pour l'affichage du profil
  
  response.write ("<table width=""95%"" border=""0"" cellspacing=""5"" cellpadding=""5"">")
  response.write ("<tr>")
  response.write ("<td width=""80"" align=""center"" valign=""top"">")
  response.write ("<table width=""90%"" border=""0"" cellspacing=""0"" cellpadding=""0"">")
  response.write ("<tr>")
  response.write ("<td>")
  response.write ("<a  href=" & media_url & " title = " & voyant & " target=""_self""><img src='" & image & "' style=""border-color:#D2D7D7"" border = ""1""/></a>")
  response.write ("</td>")
  response.write ("<tr>")
  response.write ("</table>")
  response.write ("</td>")
  
  response.write ("<td align=""left"" valign=""top"">")
  response.write ("<table width=""90%"" border=""0"" cellspacing=""0"" cellpadding=""0"" >")
  response.write ("<tr>")
  response.write ("<td height=""18"" class=""mauve_16_bold"">")
  response.write Ucase(rs.fields("voyant"))
  response.write ("</td>")
  response.write ("</tr>") 
  response.write ("<tr>")
  response.write ("<td height=""18"" class=""mauve_12_italique"">")
  response.write rs.fields("profil")
  response.write ("</td>")
  response.write ("</tr>")   
  response.write ("<tr>")
  response.write ("<td height=""18"" class=""mauve_16_bold"">")
  response.write ("<b>04.20.10.01.83</b>")
  response.write ("</td>")
  response.write ("</tr>")   
  response.write ("<tr>")
  response.write ("<td height=""18"">")
  response.write ("<a href=" & media_url & " title = ""consulter les disponibilit�s du voyant ainsi que son profil."" target=""_self"" class=""lien_gris_12"">PROFIL / PLANNING</a>")
  response.write ("</td>")
  response.write ("</tr>")   
  response.write ("<tr>")
  response.write ("<td class=" &rs.fields("voyant")& " height=""18"" class=""mauve_16_bold"">")
  
  if rs.fields("dispo") = "ON" then ' si on est disponible
  response.write ("<BR><img src='" & couleur & "' style=""border-color:#D2D7D7"" border = ""0""/>")
  end if
  
  if rs.fields("dispo") = "OFF" then ' si on est hors disponibilit&eacute;
  response.write ("<BR><img src='" & couleur2 & "' style=""border-color:#D2D7D7"" border = ""0""/>")
  if conges <> "" then
  response.write "<span class=""rouge_10_italique""> " & conges & "</span>"
  end if
  end if
  
  if rs.fields("dispo") = "OCCUPE" then ' si on est temporairement indisponible
  response.write ("<BR><img src='" & couleur3 & "' style=""border-color:#D2D7D7"" border = ""0""/>")
  end if 
   
  if rs.fields("dispo") = "ZEN" then ' si on est en consultation
  response.write ("<BR><img src='" & couleur4 & "' style=""border-color:#D2D7D7"" border = ""0""/>")
  end if   
  
  response.write ("</td>")
  response.write ("</tr>")   
  response.write ("</table>")
  response.write ("</td>")
  response.write ("</tr>")
  response.write ("</table>")  

  '
  
  response.write ("</td>")
  response.write ("</table>")
  
  response.write ("</td>")
 
 elseif i = nbchamp then ' si le nombre de colonne est atteint on ferme la ligne
  
  response.write ("<td align=""left"" valign=""top"">")
  
  response.write ("<table width=""280"" border=""0"" cellpadding=""5"" cellspacing=""5"" id=""contenu_sous_fiche"">")
  response.write ("<td align=""left"" valign=""top"">")
  
 ' cr&eacute;ation du tableau pour l'affichage du profil
  
  response.write ("<table width=""98%"" border=""0"" cellspacing=""5"" cellpadding=""5"">")
  response.write ("<tr>")
  response.write ("<td width=""80"" align=""center"" valign=""top"">")
  response.write ("<table width=""98%"" border=""0"" cellspacing=""0"" cellpadding=""0"">")
  response.write ("<tr>")
  response.write ("<td>")
  response.write ("<a href=" & media_url & " title = " & voyant & " target=""_self""><img src='" & image & "' style=""border-color:#D2D7D7"" border = ""1""/></a>") ' partie droite
  response.write ("</td>")
  response.write ("<tr>")
  response.write ("</table>")
  response.write ("</td>")
  
  response.write ("<td align=""left"" valign=""top"">")
  response.write ("<table width=""90%"" border=""0"" cellspacing=""0"" cellpadding=""0"">")
  response.write ("<tr>")
  response.write ("<td height=""18"" class=""mauve_16_bold"">")
  response.write Ucase(rs.fields("voyant"))
  response.write ("</td>")
  response.write ("</tr>") 
  response.write ("<tr>")
  response.write ("<td height=""18"" class=""mauve_12_italique"">")
  response.write rs.fields("profil")
  response.write ("</td>")
  response.write ("</tr>")   
  response.write ("<tr>")
  response.write ("<td height=""18"" class=""mauve_16_bold"">")
  response.write ("<b>04.20.10.01.83</b>")
  response.write ("</td>")
  response.write ("</tr>")   
  response.write ("<tr>")
  response.write ("<td height=""18"">")
  response.write ("<a href=" & media_url & " title = ""consulter les disponibilit�s du voyant ainsi que son profil."" target=""_self"" class=""lien_gris_12"">PROFIL / PLANNING</a>")
  response.write ("</td>")
  response.write ("</tr>")   
  response.write ("<tr>")
  response.write ("<td class=" &rs.fields("voyant")& " height=""18"" class=""mauve_16_bold"" valign=""middle"">")

  if rs.fields("dispo") = "ON" then
  response.write ("<BR><img src='" & couleur & "' style=""border-color:#D2D7D7"" border = ""0""/>")
  end if

  if rs.fields("dispo") = "OFF" then
  response.write ("<BR><img src='" & couleur2 & "' style=""border-color:#D2D7D7"" border = ""0""/>")
  if conges <> "" then
  response.write "<span class=""rouge_10_italique""> " & conges & "</span>"
  end if
  end if
  
  if rs.fields("dispo") = "OCCUPE" then ' si on est temporairement indisponible
  response.write ("<BR><img src='" & couleur3 & "' style=""border-color:#D2D7D7"" border = ""0""/>")
  end if 
   
  if rs.fields("dispo") = "ZEN" then ' si on est en consultation
  response.write ("<BR><img src='" & couleur4 & "' style=""border-color:#D2D7D7"" border = ""0""/>")
  end if     

  response.write ("</td>")
  response.write ("</tr>")   
  response.write ("</table>")
  response.write ("</td>")
  response.write ("</tr>")
  response.write ("</table>")  

  '
  
  response.write ("</td>")
  response.write ("</table>")
  
  response.write ("</td>")
  response.write ("</tr>")
  
  i = 0 ' initialisation de la routine
  
 else ' sinon on peu inscrire une deuxieme case
  
  response.write ("<td align=""left"" valign=""top"">")
  
  response.write ("<table width=""280"" border=""0"" cellpadding=""5"" cellspacing=""5"" id=""contenu_sous_fiche"">")
  response.write ("<td align=""left"" valign=""top"">")
  
 ' cr&eacute;ation du tableau pour l'affichage du profil
  
  response.write ("<table width=""95%"" border=""0"" cellspacing=""5"" cellpadding=""5"">")
  response.write ("<tr>")
  response.write ("<td width=""80"" align=""center"" valign=""top"">")
  response.write ("<table width=""90%"" border=""0"" cellspacing=""0"" cellpadding=""0"">")
  response.write ("<tr>")
  response.write ("<td>")
  response.write ("<a href=" & media_url & " title = " & voyant & " target=""_self""><img src='" & image & "' style=""border-color:#D2D7D7"" border = ""1""/></a>")
  response.write ("</td>")
  response.write ("<tr>")
  response.write ("</table>")
  response.write ("</td>")
  
  response.write ("<td align=""left"" valign=""top"">")
  response.write ("<table width=""90%"" border=""0"" cellspacing=""0"" cellpadding=""0"">")
  response.write ("<tr>")
  response.write ("<td height=""18"" class=""mauve_16_bold"">")
  response.write Ucase(rs.fields("voyant"))
  response.write ("</td>")
  response.write ("</tr>") 
  response.write ("<tr>")
  response.write ("<td height=""18"" class=""mauve_12_italique"">")
  response.write rs.fields("profil")
  response.write ("</td>")
  response.write ("</tr>")   
  response.write ("<tr>")
  response.write ("<td height=""18"">")
  response.write ("<b>04.20.10.01.83</b>")
  response.write ("</td>")
  response.write ("</tr>")   
  response.write ("<tr>")
  response.write ("<td height=""18"">")
  response.write ("<a href=" & media_url & " title = ""consulter les disponibilit�s du voyant ainsi que son profil."" target=""_self"" class=""lien_gris_12"">VOIR PROFIL</a>")
  response.write ("</td>")
  response.write ("</tr>")   
  response.write ("<tr>")
  response.write ("<td height=""18"" class=""mauve_16_bold"">")
  response.write "<b>" & rs.fields("dispo") & "</b>"
  response.write ("</td>")
  response.write ("</tr>")   
  response.write ("</table>")
  response.write ("</td>")
  response.write ("</tr>")
  response.write ("</table>")  

  '
  
  response.write ("</td>")
  response.write ("</table>")
  
  response.write ("</td>")
 end if
 
 i = i +1
 
 rs.movenext

loop

response.write "</table>" ' FERMETURE DU TABLEAU PRINCIPAL
%>
              <br>
              <br>
              <table width="90%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td>

                  </td>
                </tr>
            </table></td>
          </tr>
          </table></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="120" align="center" valign="middle"><table width="960" border="0" cellpadding="0" cellspacing="0" id="contenu_partie_bas">
      <tr>
        <td align="center" valign="middle"><!--#include file="bas.asp" -->
          <br>
          <table width="90%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td height="30" align="center" valign="middle" class="blanc_12_normal"><hr></td>
            </tr>
            <tr>
              <td height="30" align="center" valign="middle"><span class="blanc_12_normal">&copy; 2013 / 2014 - Tous droits r&eacute;serv&eacute;s - Voyance Imm&eacute;diate - n&deg; de SIRET : 413 125 048 000 12<br>
                1er site de voyance en ligne et de voyance par t&eacute;l&eacute;phone avec un v&eacute;ritable forfait &agrave; 1 euro la minute.
              </span></td>
            </tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>


<script>

setInterval(function(){

    $.ajax({
        url: 'statutVoyant.asp',
        type:'GET',
        dataType:'json',
        success:function(data){
           $.each(data,function(i,item){
           var className = '.' + item[0];
           var statut = item[1];

           $(className).empty();

            if(statut == "ON"){
                $(className).empty();
                $(className).append('<br/><img src="images/rubriques/bouton_vert.jpg">');
            } else if(statut == "OFF"){

                $(className).empty();
                $(className).append('<br/><img src="images/rubriques/bouton_rouge.jpg"> ');
            } else if(statut == "ZEN"){

                $(className).empty();
                $(className).append('<br><img src="images/rubriques/bouton_bleu.jpg"> ');

            } else if(statut == "OCCUPE"){

                $(className).empty();
                $(className).append('<br><img src="images/rubriques/bouton_gris.jpg"> ');

            }
           });

        }


    });

},1000);



</script>