<%
Dim CepOrigem, CepDestino, TipoFrete, Peso, Comprimento, Altura, Largura, Diametro, MaoPropria, ValorDeclarado, AvisoRecebimento, Url, Param, Valor, Prazo, MsgErro
MsgErro = ""

if Request("hdnBusca") = "1" then
	Call CalculaFrete()
end if

Sub CalculaFrete()
	CepOrigem  = Request("CepOrigem")
	CepDestino = Request("CepDestino")

	if isNull(CepOrigem) or trim(CepOrigem) = "" then
		'CepOrigem = "81730230"
		MsgErro = "Informe a Origem."
	elseIf isNull(CepDestino) or trim(CepDestino) = "" then
		'CepDestino = "81610200"
		MsgErro = "Informe o Destino."
	end if

	if isNull(TipoFrete) or trim(TipoFrete) = "" then
		' |-----------------------------
		' | Código do tipo de frete:
		' |-----------------------------
		' | 40010 SEDEX sem contrato
		' | 41106 PAC sem contrato
		' |	41211 PAC com contrato
		' |	41068 PAC com contrato
		' | 40215 SEDEX 10 sem contrato
		' | 40290 SEDEX Hoje sem contrato
		' | 04804 SEDEX Hoje à vista 
		' |-----------------------------
		TipoFrete 			= "41106" 'Código de cada tipo de frete
	end if

	' |-----------------------------
	' | ****** OS VALORES ABAIXO SÃO OS VALORES MÍNIMOS DE MEDIDAS.
	' | ****** SE FOREM INFORMADOS VALORES MENORES, O VALOR DO FRETE E PRAZO RETORNARÃO ZERADOS.
	' |-----------------------------
	if isNull(Peso) or trim(Peso) = "" then
		Peso 				= "1"	'Peso em Kg
	end if
	if isNull(Comprimento) or trim(Comprimento) = "" then
		Comprimento			= "15"	'Em cm
	end if
	if isNull(Altura) or trim(Altura) = "" then
		Altura 				= "5"	'Em cm
	end if
	if isNull(Largura) or trim(Largura) = "" then
		Largura 			= "10"	'Em cm
	end if
	if isNull(Diametro) or trim(Diametro) = "" then
		Diametro 			= "0"	'Em cm
	end if
	if isNull(MaoPropria) or trim(MaoPropria) = "" then
		MaoPropria			= "s"	's para sim n para não
	end if
	if isNull(ValorDeclarado) or trim(ValorDeclarado) = "" then
		ValorDeclarado		= "0" 'Em Reais
	end if
	if isNull(AvisoRecebimento) or trim(AvisoRecebimento) = "" then
		AvisoRecebimento	= "n"	's para sim n para não
	end if

	Url = "http://ws.correios.com.br/calculador/CalcPrecoPrazo.asmx/CalcPrecoPrazo"
	Param = "?" &_
		"nCdEmpresa="&_
		"&sDsSenha="&_
		"&nCdServico="&TipoFrete&_
		"&sCepOrigem="&CepOrigem&_
		"&sCepDestino="&CepDestino&_
		"&nVlPeso="&Peso&_
		"&nCdFormato=1"&_
		"&nVlComprimento="&Comprimento&_
		"&nVlAltura="&Altura&_
		"&nVlLargura="&Largura&_
		"&nVlDiametro="&Diametro&_
		"&sCdMaoPropria="&MaoPropria&_
		"&nVlValorDeclarado="&ValorDeclarado&_
		"&sCdAvisoRecebimento="&AvisoRecebimento
	if MsgErro = "" then
		Call ChamaAPICorreios()
	end if
End Sub

Sub ChamaAPICorreios()
	'Abrindo XML
	Dim XML
	Set XML = Server.CreateObject("MSXML2.XMLHTTP")
	XML.open "GET", (url & Param), false
	XML.setRequestHeader "Content-Type", "text/XML"
	XML.Send

	'Tratando XML
	set xmlRss = Server.CreateObject("Microsoft.XMLDOM")
	xmlRss.async = false
	xmlRss.loadXml(XML.ResponseText)
	set xmlValor = xmlRss.getElementsByTagName("Valor")
	set xmlPrazo = xmlRss.getElementsByTagName("PrazoEntrega")
	
	for i = 0 to xmlValor.length-1
		Valor =  xmlValor.item(i).childNodes.item(0).text
	next
	for i = 0 to xmlPrazo.length-1
		Prazo = xmlPrazo.item(i).childNodes.item(0).text
	next
	set xmlRss = nothing
End Sub
%>

<html>
	<body>
		<form name="formCep" id="formCep" action="#" method="post">
			<input type="hidden" id="hdnBusca" name="hdnBusca" value="1" />
			<div style="width:240px;margin:5px;padding:5px;border:1px solid #ddd;-webkit-border-radius: 5px;-moz-border-radius: 5px;border-radius: 5px;">
				<table>
					<tr>
						<td>CEP Origem: </td>
						<td><input type="text" id="CepOrigem" name="CepOrigem" value="" size="8" onkeydown="limpaMsg()"/></td>
						<td rowspan="2"><img src="caminhao_correios.png" style="width:40px;" onkeydown="limpaMsg()"/></td>
					</tr>
					<tr>
						<td>CEP Destino:</td>
						<td><input type="text" id="CepDestino" name="CepDestino" value="" size="8"/></td>
					</tr>
					<tr>
						<td></td>
						<td colspan="2" style="text-align:left;"><input type="submit" id="btnCalcula" name="btnCalcula" value="Calcular Frete"/></td>
					</tr>
				</table>
				<%if Request("hdnBusca") = "1" and MsgErro = "" then%>
					<br />
					<div>
						<span>Valor: R$ <%=Valor%></span>
						<br/>
						<span>Prazo: <%=Prazo%> dia(s)</span>
					</div>
				<%end if%>
				<%if MsgErro <> "" then%>
					<span style="color:#F00;font-weight:bold;" id="msgErro" name="msgErro"><%=MsgErro %></span>
				<%end if%>
			</div>
		</form>
		<script type="text/javascript">
			function limpaMsg() {
                document.getElementById("msgErro").innerHTML = "";
            }
		</script>
	</body>
</html>