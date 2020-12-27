Option Compare Database

'html


Public Function html_header() As String
Dim HTML As String

HTML = ""
HTML = HTML & "<!DOCTYPE html>" & vbCr
HTML = HTML & "<html lang=" & Chr(34) & "nl" & Chr(34) & ">" & vbCr

HTML = HTML & "<head>" & vbCr
HTML = HTML & "    <meta charset=" & Chr(34) & "utf-8" & Chr(34) & ">" & vbCr
HTML = HTML & "    <meta http-equiv=" & Chr(34) & "X-UA-Compatible" & Chr(34) & " content=" & Chr(34) & "IE=edge" & Chr(34) & ">" & vbCr
HTML = HTML & "    <meta name=" & Chr(34) & "viewport" & Chr(34) & " content=" & Chr(34) & "width=device-width, initial-scale=1" & Chr(34) & ">" & vbCr
HTML = HTML & "    <!-- The above 3 meta tags *must* come first in the head  any other head content must come *after* these tags --> " & vbCr

HTML = HTML & "    <!-- favicon -->"

HTML = HTML & "    <link rel=" & Chr(34) & "apple-touch-icon" & Chr(34) & " sizes=" & Chr(34) & "180x180" & Chr(34) & " href=" & Chr(34) & "/apple-touch-icon.png" & Chr(34) & ">" & vbCr
HTML = HTML & "    <link rel=" & Chr(34) & "icon" & Chr(34) & " type=" & Chr(34) & "image/png" & Chr(34) & " sizes=" & Chr(34) & "32x32" & Chr(34) & " href=" & Chr(34) & "/favicon-32x32.png" & Chr(34) & ">" & vbCr
HTML = HTML & "    <link rel=" & Chr(34) & "icon" & Chr(34) & " type=" & Chr(34) & "image/png" & Chr(34) & " sizes=" & Chr(34) & "16x16" & Chr(34) & " href=" & Chr(34) & "/favicon-16x16.png" & Chr(34) & ">" & vbCr
HTML = HTML & "    <link rel=" & Chr(34) & "manifest" & Chr(34) & " href=" & Chr(34) & "/site.webmanifest" & Chr(34) & ">" & vbCr
HTML = HTML & "    <link rel=" & Chr(34) & "mask-icon" & Chr(34) & " href=" & Chr(34) & "/safari-pinned-tab.svg" & Chr(34) & " color=" & Chr(34) & "#5bbad5" & Chr(34) & ">" & vbCr
HTML = HTML & "   <meta name=" & Chr(34) & "msapplication-TileColor" & Chr(34) & " content=" & Chr(34) & "#da532c" & Chr(34) & ">" & vbCr
HTML = HTML & "   <meta name=" & Chr(34) & "theme-color" & Chr(34) & " content=" & Chr(34) & "#ffffff" & Chr(34) & ">" & vbCr
HTML = HTML & "    <!-- /favicon -->" & vbCr

HTML = HTML & "    <title>BC70 Step Portaal</title>" & vbCr

HTML = HTML & "    <style> " & vbCr
        
HTML = HTML & "        td, p, body {" & vbCr
HTML = HTML & "    font-size: 10pt;" & vbCr
HTML = HTML & "    font-family: Verdana, Arial, Helvetica, sans-serif;" & vbCr
HTML = HTML & "    color: #000000;" & vbCr
HTML = HTML & "}" & vbCr
HTML = HTML & vbCr
HTML = HTML & " h1 {" & vbCr
HTML = HTML & "     font-size: 14pt;" & vbCr
HTML = HTML & "     font-family: Verdana, Arial, Helvetica, sans-serif;" & vbCr
HTML = HTML & "     font-variant: small-caps;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " h2 {" & vbCr
HTML = HTML & "     font-size: 12pt;" & vbCr
HTML = HTML & "     font-family: Verdana, Arial, Helvetica, sans-serif;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " h3 {" & vbCr
HTML = HTML & "     font-size: 10pt;" & vbCr
HTML = HTML & "     font-family: Verdana, Arial, Helvetica, sans-serif;" & vbCr
HTML = HTML & "     font-weight: bold;" & vbCr
HTML = HTML & "     color: #003399;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " h3 {" & vbCr
HTML = HTML & "     margin-bottom: 0;" & vbCr
HTML = HTML & " }" & vbCr & vbCr
HTML = HTML & vbCr
HTML = HTML & " th, h1, h2, h3 {" & vbCr
HTML = HTML & "     color: #003399;" & vbCr
HTML = HTML & "     background-color: #FFFFFF;" & vbCr
HTML = HTML & " }" & vbCr & vbCr
HTML = HTML & vbCr
HTML = HTML & " p {" & vbCr
HTML = HTML & "     margin-top: 0;" & vbCr
HTML = HTML & " }" & vbCr & vbCr
HTML = HTML & vbCr
HTML = HTML & " th {" & vbCr
HTML = HTML & "     font-size: 20pt;" & vbCr
HTML = HTML & "     font-family: Verdana, Arial, Helvetica, sans-serif;" & vbCr
HTML = HTML & "     font-weight: normal;" & vbCr
HTML = HTML & "     text-align: center;" & vbCr
HTML = HTML & "     border: 0 solid Black;" & vbCr
HTML = HTML & "     color: #003399;" & vbCr
HTML = HTML & "     background-color: #FFFFFF;" & vbCr
HTML = HTML & " }" & vbCr & vbCr
HTML = HTML & vbCr
HTML = HTML & " table.data {" & vbCr
HTML = HTML & "     border: 1px #6699CC solid;" & vbCr
HTML = HTML & "     border-collapse: collapse;" & vbCr
HTML = HTML & "     border-spacing: 0;" & vbCr
HTML = HTML & "     background-color: #FAFAFA;" & vbCr
HTML = HTML & "     width:100%;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " table.data td {" & vbCr
HTML = HTML & "     border-bottom: 1px solid #99CCFF;" & vbCr
HTML = HTML & "     border-top: 0;" & vbCr
HTML = HTML & "     border-left: 1px solid #9CF;" & vbCr
HTML = HTML & "     border-right: 0;" & vbCr
HTML = HTML & "     background-color: #FAFAFA;" & vbCr
HTML = HTML & "     vertical-align: top;" & vbCr
HTML = HTML & "     padding: 2px;" & vbCr
HTML = HTML & " }" & vbCr & vbCr

HTML = HTML & " table.data th {" & vbCr
HTML = HTML & "     font-size: 13px;" & vbCr
HTML = HTML & "     border-bottom: 2px solid #6699CC;" & vbCr
HTML = HTML & "     border-left: 1px solid #6699CC;" & vbCr
HTML = HTML & "     font-weight: bold;" & vbCr
HTML = HTML & "     background-color: #BEC8D1;" & vbCr
HTML = HTML & "     vertical-align: top;" & vbCr
HTML = HTML & "     padding: 2px;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " table.data-rotate {" & vbCr
HTML = HTML & "     border: 1px #6699CC solid;" & vbCr
HTML = HTML & "     border-collapse: collapse;" & vbCr
HTML = HTML & "     font-size: 7px;" & vbCr
HTML = HTML & "     border-spacing: 0;" & vbCr

HTML = HTML & "     background-color: #FAFAFA;" & vbCr

HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " table.data-rotate td {" & vbCr
HTML = HTML & "     font-size: 9px;" & vbCr
HTML = HTML & "     border-bottom: 1px solid #99CCFF;" & vbCr
HTML = HTML & "     border-top: 0;" & vbCr
HTML = HTML & "     border-left: 1px solid #9CF;" & vbCr
HTML = HTML & "     border-right: 0;" & vbCr
HTML = HTML & "     background-color: #FAFAFA;" & vbCr
HTML = HTML & "     vertical-align: top;" & vbCr
HTML = HTML & "     padding: 2px;" & vbCr
HTML = HTML & "     white-space:nowrap;" & vbCr
HTML = HTML & " }" & vbCr & vbCr
HTML = HTML & vbCr
HTML = HTML & " table.data-rotate th {" & vbCr
HTML = HTML & "     font-size: 13px;" & vbCr
HTML = HTML & "     border-bottom: 2px solid #6699CC;" & vbCr
HTML = HTML & "     border-left: 1px solid #6699CC;" & vbCr
HTML = HTML & "     font-weight: bold;" & vbCr
HTML = HTML & "     background-color: #BEC8D1;" & vbCr
HTML = HTML & "     vertical-align: bottom;" & vbCr
HTML = HTML & "     padding: 2px;" & vbCr
HTML = HTML & "     white-space:nowrap;" & vbCr
HTML = HTML & "     text-align: center;" & vbCr
HTML = HTML & "     padding: 0 0 8px 0;" & vbCr
HTML = HTML & "     height: 120px;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr

HTML = HTML & " .data-rotate-header-container {" & vbCr
HTML = HTML & "     width: 10px;" & vbCr
HTML = HTML & "     transform-origin: bottom left;" & vbCr
HTML = HTML & "     transform: translateX(30px) rotate(-90deg);" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr

HTML = HTML & " tr.data-3  {" & vbCr
HTML = HTML & "     border-top: 3px solid #99CCFF;" & vbCr
'HTML = HTML & "     border-top: 0;" & vbCr
'HTML = HTML & "     border-left: 1px solid #9CF;" & vbCr
'HTML = HTML & "     border-right: 0;" & vbCr
'HTML = HTML & "     font-weight: bold;" & vbCr
'HTML = HTML & "     background-color: #FAFAFA;" & vbCr
'HTML = HTML & "     vertical-align: top;" & vbCr
'HTML = HTML & "     text-align: left;" & vbCr
'HTML = HTML & "     padding:15px;" & vbCr
HTML = HTML & " }" & vbCr & vbCr
HTML = HTML & " tr.data-2 {" & vbCr
HTML = HTML & "     border-bottom: 2px solid #99CCFF;" & vbCr
'HTML = HTML & "     border-top: 0;" & vbCr
'HTML = HTML & "     border-left: 1px solid #9CF;" & vbCr
'HTML = HTML & "     border-right: 0;" & vbCr
'HTML = HTML & "     font-weight: bold;" & vbCr
'HTML = HTML & "     background-color: #FAFAFA;" & vbCr
'HTML = HTML & "     vertical-align: top;" & vbCr
'HTML = HTML & "     text-align: right;" & vbCr
'HTML = HTML & "     padding: 0;" & vbCr
HTML = HTML & " }" & vbCr & vbCr


HTML = HTML & vbCr
HTML = HTML & " .data-rotate-header-content {" & vbCr
HTML = HTML & "     width: 10px;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr


HTML = HTML & " table.error {" & vbCr
HTML = HTML & "     border: 2pt solid #FF0000;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " td.error {" & vbCr
HTML = HTML & "     font-weight: bold;" & vbCr
HTML = HTML & "     font-size: 11pt;" & vbCr
HTML = HTML & "     font-family: Verdana, Arial, sans-serif;" & vbCr
HTML = HTML & "     color: #000000;" & vbCr
HTML = HTML & "     background-color: #FFE0E0;" & vbCr
HTML = HTML & "     padding: 2pt 4pt 2pt 4pt;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " li.error {" & vbCr
HTML = HTML & "     font-size: 11pt;" & vbCr
HTML = HTML & "     font-weight: bold;" & vbCr
HTML = HTML & "     font-family: Verdana, Arial, sans-serif;" & vbCr
HTML = HTML & "     text-decoration: none;" & vbCr
HTML = HTML & "     padding: 2pt 4pt 2pt 4pt;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " table.pageframe {" & vbCr
HTML = HTML & "     border: 1px solid #999999;" & vbCr
HTML = HTML & "     background-color: #FFFFFF;" & vbCr
HTML = HTML & "     width: 690px;" & vbCr
HTML = HTML & "     border-spacing: 0;" & vbCr
HTML = HTML & "     padding: 0;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " table.resultsframe {" & vbCr
HTML = HTML & "     border: 0;" & vbCr
HTML = HTML & "     width: 630px;" & vbCr
HTML = HTML & "     border-spacing: 20px;" & vbCr
HTML = HTML & "     padding: 0;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " table.resultsframe_inner {" & vbCr
HTML = HTML & "     border: 0;" & vbCr
HTML = HTML & "     width: 600px;" & vbCr
HTML = HTML & "     border-spacing: 0;" & vbCr
HTML = HTML & "     padding: 0;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " td.header {" & vbCr
HTML = HTML & "     text-align: right;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " td.sectionheader {" & vbCr
HTML = HTML & "     word-wrap: break-word;" & vbCr
HTML = HTML & "     font-size: 10pt;" & vbCr
HTML = HTML & "     font-family: Verdana, Arial, sans-serif;" & vbCr
HTML = HTML & "     font-weight: bold;" & vbCr
HTML = HTML & "     text-align: center;" & vbCr
HTML = HTML & "     color: #FFFFFF;" & vbCr
HTML = HTML & "     background-color: #009D57;" & vbCr
HTML = HTML & "     padding: 2px 4px 2px 4px;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " td.footer {" & vbCr
HTML = HTML & "     font-size: 10pt;" & vbCr
HTML = HTML & "     font-family: Verdana, Arial, sans-serif;" & vbCr
HTML = HTML & "     font-weight: bold;" & vbCr
HTML = HTML & "     text-align: right;" & vbCr
HTML = HTML & "     color: #FFFFFF;" & vbCr
HTML = HTML & "     background-color: #009D57;" & vbCr
HTML = HTML & "     padding: 2px 4px 2px 4px;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " a.footer {" & vbCr
HTML = HTML & "     color: #FFFFFF;" & vbCr
HTML = HTML & "    text-decoration: none;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " a.footer:hover {" & vbCr
HTML = HTML & "     text-decoration: underline;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " div.centered {" & vbCr
HTML = HTML & "     text-align: center;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " div.centered table {" & vbCr
HTML = HTML & "     margin: 0 auto;" & vbCr
HTML = HTML & "     text-align: left;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " body {" & vbCr
HTML = HTML & "     color: #000000;" & vbCr
HTML = HTML & "     background-color: #EFF6EF;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " @media print {" & vbCr
HTML = HTML & "     body {" & vbCr
HTML = HTML & "         color: #000000;" & vbCr
 HTML = HTML & "        background-color: #FFFFFF;" & vbCr
 HTML = HTML & "    }"
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " th.resultsheader {" & vbCr
HTML = HTML & "     font-size: 7pt;" & vbCr
HTML = HTML & "     border-width: 0 0 1pt 0;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & " th.resultsheaderright {" & vbCr
HTML = HTML & "     font-size: 7pt;" & vbCr
HTML = HTML & "     border-width: 0 1pt 1pt 0;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " th.resultsheaderleft {" & vbCr
HTML = HTML & "     font-size: 7pt;" & vbCr
HTML = HTML & "     border-width: 0 0 1pt 1pt;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " th.boardheaderleft {" & vbCr
HTML = HTML & "     color: #003399;" & vbCr
HTML = HTML & "     background-color: #FFFFFF;" & vbCr
HTML = HTML & "     font-size: 8pt;" & vbCr
HTML = HTML & "     text-align: left;" & vbCr
HTML = HTML & "     font-weight: bold;" & vbCr
HTML = HTML & "     border-color: #000000;" & vbCr
HTML = HTML & "     border-width: 1pt 0 1pt 0;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " th.boardheaderright {" & vbCr
HTML = HTML & "     color: #003399;" & vbCr
HTML = HTML & "     background-color: #FFFFFF;" & vbCr
HTML = HTML & "     font-size: 8pt;" & vbCr
HTML = HTML & "     text-align: right;" & vbCr
HTML = HTML & "     vertical-align: middle;" & vbCr
HTML = HTML & "     font-weight: bold;" & vbCr
HTML = HTML & "     border-color: #000000;" & vbCr
HTML = HTML & "     border-width: 1pt 0 1pt 0;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " td.boardchairlabel {" & vbCr
HTML = HTML & "     color: #003399;" & vbCr
HTML = HTML & "     background-color: #FFFFFF;" & vbCr
HTML = HTML & "     font-size: 8pt;" & vbCr
HTML = HTML & "     font-weight: normal;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " th.biddingchairlabel {" & vbCr
HTML = HTML & "     font-size: 8pt;" & vbCr
HTML = HTML & "     font-weight: bold;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " th.biddingchairplayer {" & vbCr
HTML = HTML & "     font-size: 8pt;" & vbCr
HTML = HTML & "     border-width: 0 0 1pt 0;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " td.fieldrecreationalrowleft {" & vbCr
HTML = HTML & "     border: 1pt solid #999999;" & vbCr
HTML = HTML & "     border-right-width: 0;" & vbCr
HTML = HTML & "     background-color: #F5F5F5;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " td.fieldrecreationalrowmiddle {" & vbCr
HTML = HTML & "    border: 1pt solid #999999;" & vbCr
HTML = HTML & "    border-left-width: 0;" & vbCr
HTML = HTML & "     border-right-width: 0;" & vbCr
HTML = HTML & "     background-color: #F5F5F5;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " td.fieldrecreationalrowright {" & vbCr
HTML = HTML & "     border: 1pt solid #999999;" & vbCr
HTML = HTML & "     border-left-width: 0;" & vbCr
HTML = HTML & "     background-color: #F5F5F5;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " table.results {" & vbCr
'html = html & "     border-style: solid;" & vbCr
'html = html & "     border-width:  0 0 1px 0;" & vbCr
HTML = HTML & "     border-spacing: 0;" & vbCr
HTML = HTML & "     padding: 1px;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " th.resultsheader {" & vbCr
HTML = HTML & "     white-space:nowrap;" & vbCr
HTML = HTML & "     font-size: 7pt;" & vbCr
HTML = HTML & "     text-transform: lowercase;" & vbCr
HTML = HTML & "     border-width: 0 0 1px 0;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " td.detailalignleftborderright {" & vbCr
HTML = HTML & "     white-space:nowrap;" & vbCr
HTML = HTML & "     text-align: left;" & vbCr
HTML = HTML & "     border-style: solid;"
HTML = HTML & "     border-width: 0 1px 0 0;" & vbCr
HTML = HTML & "     padding: 0 15px;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " td.detailalignleftborderrightbottom {" & vbCr
HTML = HTML & "     white-space:nowrap;" & vbCr
HTML = HTML & "     text-align: left;" & vbCr
HTML = HTML & "     border-style: solid;"
HTML = HTML & "     border-width: 0 1px 1px 0;" & vbCr
HTML = HTML & "     padding: 0 15px;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " td.detailalignrightborderright {" & vbCr
HTML = HTML & "     white-space:nowrap;" & vbCr
HTML = HTML & "     text-align: right;" & vbCr
HTML = HTML & "     border-style: solid;"
HTML = HTML & "     border-width: 0 1px 0 0;" & vbCr
HTML = HTML & "     padding: 0 15px;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " td.detailalignrightborderrightbottom {" & vbCr
HTML = HTML & "     white-space:nowrap;" & vbCr
HTML = HTML & "     text-align: right;" & vbCr
HTML = HTML & "     border-style: solid;"
HTML = HTML & "     border-width: 0 1px 1px 0;" & vbCr
HTML = HTML & "     padding: 0 15px;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " td.detailalignleftborderleft {" & vbCr
HTML = HTML & "     white-space:nowrap;" & vbCr
HTML = HTML & "     text-align: left;" & vbCr
HTML = HTML & "     border-style: solid;"
HTML = HTML & "     border-width: 0 0 0 1px;" & vbCr
HTML = HTML & "     padding: 0 15px;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " td.detailalignleftborderleftbottom {" & vbCr
HTML = HTML & "     white-space:nowrap;" & vbCr
HTML = HTML & "     text-align: left;" & vbCr
HTML = HTML & "     border-style: solid;"
HTML = HTML & "     border-width: 0 0 1px 1px;" & vbCr
HTML = HTML & "     padding: 0 15px;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " td.detailalignleftborderbottom {" & vbCr
HTML = HTML & "     white-space:nowrap;" & vbCr
HTML = HTML & "     text-align: left;" & vbCr
HTML = HTML & "     border-style: solid;"
HTML = HTML & "     border-width: 0 0 1px 0;" & vbCr
HTML = HTML & "     padding: 0 15px;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " td.detailalignrightborderleft {" & vbCr
HTML = HTML & "     white-space:nowrap;" & vbCr
HTML = HTML & "     text-align: right;" & vbCr
HTML = HTML & "     border-style: solid;"
HTML = HTML & "     border-width: 0 0 0 1px;" & vbCr
HTML = HTML & "     padding: 0 15px;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " td.detailalignrightborderleftbottom {" & vbCr
HTML = HTML & "     white-space:nowrap;" & vbCr
HTML = HTML & "     text-align: right;" & vbCr
HTML = HTML & "     border-style: solid;"
HTML = HTML & "     border-width: 0 0 1px 1px;" & vbCr
HTML = HTML & "     padding: 0 15px;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " td.detailalignrightborderbottom {" & vbCr
HTML = HTML & "     white-space:nowrap;" & vbCr
HTML = HTML & "     text-align: right;" & vbCr
HTML = HTML & "     border-style: solid;"
HTML = HTML & "     border-width: 0 0 1px 0;" & vbCr
HTML = HTML & "     padding: 0 15px;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " td.detailaligncenterborderbottom {" & vbCr
HTML = HTML & "     white-space:nowrap;" & vbCr
HTML = HTML & "     text-align: center;" & vbCr
HTML = HTML & "     border-style: solid;"
HTML = HTML & "     border-width: 0 0 1px 0;" & vbCr
HTML = HTML & "     padding: 0 15px;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " td.detailaligncenterborderbottomnopadding {" & vbCr
HTML = HTML & "     white-space:nowrap;" & vbCr
HTML = HTML & "     text-align: center;" & vbCr
HTML = HTML & "     border-style: solid;"
HTML = HTML & "     border-width: 0 0 1px 0;" & vbCr
HTML = HTML & "     padding: 0;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " table.field {" & vbCr
HTML = HTML & "     border-spacing: 0;" & vbCr
HTML = HTML & "     padding: 1px;" & vbCr
HTML = HTML & "     width: 100%;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " table.board {" & vbCr
HTML = HTML & "     border-width: 0;" & vbCr
HTML = HTML & "     border-spacing: 0;" & vbCr
HTML = HTML & "     padding: 1px;" & vbCr
HTML = HTML & "     width: 100%;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " th.fieldheader {" & vbCr
HTML = HTML & "     font-size: 7pt;" & vbCr
HTML = HTML & "     text-transform: lowercase;" & vbCr
HTML = HTML & "     border-width: 0 0 1px 0;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " tr.fieldrow {" & vbCr
HTML = HTML & "     font-size: 11pt;" & vbCr
HTML = HTML & "     white-space: nowrap;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " tr.fieldrowselected {" & vbCr
HTML = HTML & "     font-size: 11pt;" & vbCr
HTML = HTML & "     outline: thin solid red;" & vbCr
HTML = HTML & "     background-color: #FFF4F4;" & vbCr
HTML = HTML & "     white-space: nowrap;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " table.overview {" & vbCr
HTML = HTML & "     border-spacing: 0;" & vbCr
HTML = HTML & "     padding: 1px;" & vbCr
HTML = HTML & "     width: 100%;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " th.overviewheader {" & vbCr
HTML = HTML & "     font-size: 7pt;" & vbCr
HTML = HTML & "     text-transform: lowercase;" & vbCr
HTML = HTML & "     border-width: 0 0 1px 0;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " tr.overviewrow {" & vbCr
HTML = HTML & "     font-size: 11pt;" & vbCr
HTML = HTML & "     white-space: nowrap;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " tr.overviewrowselected {" & vbCr
HTML = HTML & "     font-size: 11pt;" & vbCr
HTML = HTML & "     outline: thin solid red;" & vbCr
HTML = HTML & "     background-color: #FFF4F4;" & vbCr
HTML = HTML & "     white-space: nowrap;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " a {" & vbCr
HTML = HTML & "     color: blue;" & vbCr
HTML = HTML & "     text-decoration: none;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " a.external {" & vbCr
HTML = HTML & "     font-size: 8pt;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & " a.board {" & vbCr
HTML = HTML & "     color: blue;" & vbCr
HTML = HTML & "     text-decoration: underline;" & vbCr
HTML = HTML & " }" & vbCr
HTML = HTML & vbCr
HTML = HTML & "     </style>" & vbCr
HTML = HTML & "<script language=" & Chr(34) & "JavaScript" & Chr(34) & "type=" & Chr(34) & "text/javascript" & Chr(34) & ">" & vbCr

HTML = HTML & "function goBack() { " & vbCr
HTML = HTML & "  window.history.back();" & vbCr
HTML = HTML & "}" & vbCr
HTML = HTML & "</script>" & vbCr



HTML = HTML & " </head>" & vbCr
HTML = HTML & vbCr

html_header = HTML


End Function



Public Function html_Begin_Body() As String
Dim HTML As String

    HTML = ""
    HTML = HTML & "<body>" & vbCr
    HTML = HTML & " <div class=" & Chr(34) & "centered" & Chr(34) & ">" & vbCr
    HTML = HTML & "     <table class=" & Chr(34) & "pageframe" & Chr(34) & ">" & vbCr
    
 ' eerste rij
    
     HTML = HTML & "    <tr>" & vbCr
     HTML = HTML & "        <td style=" & Chr(34) & "background-color:#269f59" & Chr(34) & ">" & vbCr
     HTML = HTML & "             <div style=" & Chr(34) & "height: 142px; overflow:hidden;" & Chr(34) & ">" & vbCr
     HTML = HTML & "                   <img style=" & Chr(34) & "height: 142px" & Chr(34) & " "
     HTML = HTML & " src=" & Chr(34) & "https://portal.stepbridge.nl/images/logo_nl.png" & Chr(34) & " "
     HTML = HTML & " alt=" & Chr(34) & "Logo" & Chr(34) & " > " & vbCr
     HTML = HTML & "             </div>" & vbCr
     HTML = HTML & "        </td>" & vbCr
     HTML = HTML & "    </tr>" & vbCr
  'tweede rij een lege regel
  
     HTML = HTML & rij_Lege_regel_met_back()
     
     html_Begin_Body = HTML
End Function

Public Function rij_sectionheader(Titel As Variant) As String
Dim HTML As String

    HTML = ""
    HTML = HTML & "   <tr>" & vbCr
    HTML = HTML & "      <td class=" & Chr(34) & "sectionheader" & Chr(34) & "><a id=" & Chr(34) & "begin" & Chr(34) & ">" & Titel & "</a></td> " & vbCr
    HTML = HTML & "  </tr>" & vbCr
    
    rij_sectionheader = HTML
End Function

Public Function rij_teamheader(Team As Variant, refTeam As Variant) As String
Dim HTML As String

    HTML = ""
    HTML = HTML & "   <tr>" & vbCr
    HTML = HTML & "      <td class=" & Chr(34) & "sectionheader" & Chr(34) & "><a href=" & Chr(34) & refTeam & Chr(34) & ">" & Team & "</a></td> " & vbCr
    HTML = HTML & "  </tr>" & vbCr
    rij_teamheader = HTML
End Function

Public Function rij_header(Title As Variant) As String
Dim HTML As String

    HTML = ""
    HTML = HTML & "   <tr>" & vbCr
    HTML = HTML & "      <td class=" & Chr(34) & "sectionheader" & Chr(34) & ">" & Title & "</td> " & vbCr
    HTML = HTML & "  </tr>" & vbCr
    rij_header = HTML
End Function


Public Function rij_Paren(Paar1 As Variant, Paar2 As Variant, RefPaar1 As Variant, refPaar2 As Variant) As String
Dim HTML As String

    HTML = ""
    HTML = HTML & "   <tr style=" & Chr(34) & "vertical-align: top;" & Chr(34) & ">" & vbCr
    HTML = HTML & nAlign("td", "center", "") & "<a href=" & Chr(34) & RefPaar1 & Chr(34) & ">" & Paar1 & "</a></td> " & vbCr
    HTML = HTML & nAlign("td", "center", "") & "<a href=" & Chr(34) & refPaar2 & Chr(34) & ">" & Paar2 & "</a></td> " & vbCr
    HTML = HTML & "     <td >&nbsp;</td> " & vbCr
    HTML = HTML & "  </tr>" & vbCr
    rij_Paren = HTML
End Function
Public Function begin_kolommen() As String
  Dim HTML As String

    HTML = ""
    HTML = HTML & "   <tr>" & vbCr
    HTML = HTML & "   <td>" & vbCr
    HTML = HTML & "<table id=" & Chr(34) & "Scoresheet" & Chr(34) & ">" & vbCr
    HTML = HTML & "  <tr style=" & Chr(34) & "vertical-align: top;" & Chr(34) & ">"
    begin_kolommen = HTML
    
    
End Function

Public Function eind_kolommen() As String
Dim HTML As String
HTML = ""
HTML = HTML & "</table>" & vbCr
HTML = HTML & "</td>" & vbCr
HTML = HTML & "</tr>" & vbCr
eind_kolommen = HTML

End Function

Public Function eind_kolom() As String
Dim HTML As String
HTML = ""
HTML = HTML & "</table>" & vbCr
HTML = HTML & "</td>" & vbCr

eind_kolom = HTML

End Function

Public Function html_Einde_Body() As String
    html_Einde_Body = "<body>" & vbCr & "</html>" & vbCr
End Function
Public Function html_Einde_Body_scoresheet() As String
Dim HTML As String

HTML = HTML & "</table>" & vbCr
HTML = HTML & "</div>" & vbCr
HTML = HTML & "</html>" & vbCr

    html_Einde_Body_scoresheet = HTML
End Function
Function rij_Lege_regel() As String

Dim HTML As String
    HTML = ""
     HTML = HTML & "    <tr>" & vbCr
     HTML = HTML & "       <td>&nbsp;</td>" & vbCr
     HTML = HTML & "    </tr>" & vbCr
     
     rij_Lege_regel = HTML
End Function

'<a href="javascript: history.go(-1)">Go Back</a>
Function rij_Lege_regel_met_back() As String

Dim HTML As String
    HTML = ""
     HTML = HTML & "    <tr>" & vbCr
     HTML = HTML & "       <td style=" & Chr(34) & "text-align: right" & Chr(34) & "><a href=" & Chr(34) & "javascript: history.go(-1)" & Chr(34) & ">&nbsp;&larr;</a></td>" & vbCr
     HTML = HTML & "    </tr>" & vbCr
     
     rij_Lege_regel_met_back = HTML
End Function

Public Function Align(TableTag As Variant, KindAlign As Variant, Padding As Variant) As String
Dim HTML As String
 HTML = ""
 HTML = HTML & "<" & TableTag & " style=" & Chr(34) & "text-align: " & KindAlign
 
 If Padding = "" Then
 HTML = HTML & Chr(34) & ">"
 Else
  HTML = HTML & "; padding: " & Padding & Chr(34) & ">"
 End If
 Align = HTML
End Function
Public Function nAlign(TableTag As Variant, KindAlign As Variant, Padding As Variant) As String
Dim HTML As String
 HTML = ""
 HTML = HTML & "<" & TableTag & " style=" & Chr(34) & "text-align: " & KindAlign & "; white-space:nowrap;"
 
 If Padding = "" Then
 HTML = HTML & Chr(34) & ">"
 Else
  HTML = HTML & "; padding: " & Padding & Chr(34) & ">"
 End If
 nAlign = HTML
End Function

'white-space:nowrap;


Public Function ImageCard(Card As Variant) As String
 Dim HTML As String
 Dim cardnumber As Integer
 HTML = ""
 HTML = HTML & "<img src=" & Chr(34) & "/images/suit"
 
 Select Case Card
 Case "S"
 cardnumber = 411
 Case "H"
 cardnumber = 311
 Case "D"
 cardnumber = 211
 Case "C"
 cardnumber = 111
 Case Else
    HTML = "SA"
 End Select
 
 If HTML <> "" And HTML <> "SA" Then
    HTML = HTML & cardnumber & ".gif" & Chr(34) & " alt=" & Chr(34) & Card & Chr(34) & ">"
 End If
 
 ImageCard = HTML

End Function

'Public Function ScoresheetRow(Boardnr As Variant, Contract As Variant, Resultaat As Variant, Door As Variant, Score As Variant, ButlerScore As Variant) As String
Public Function ScoresheetRow(Boardnr As Variant, Contract As Variant, resultaat As Variant, Door As Variant, score As Variant) As String


Dim HTML, contracthoogte, kleur, strcontract As String
Dim hoogte As Integer
HTML = ""

        'test op niet gespeeld of kunstmatige score
        
        
        HTML = HTML & "   <tr>" & vbCr
        HTML = HTML & "<td class=" & Chr(34) & "detailalignrightborderleft" & Chr(34) & " >" & Boardnr & "</td>" & vbCr
        'ontleed contact
        
        If Contract = "" Then
            strcontract = "&nbsp;"
        Else
            hoogte = Val(Contract)
            If hoogte = 0 Then
                contracthoogte = "pass"
                kleur = ""
            Else
                contracthoogte = Left(Contract, 1)
                kleur = Mid(Contract, 2)
            End If
            '
            If InStr(kleur, "SA") > 0 Then
                    strcontract = contracthoogte & kleur
                 Else
                strcontract = contracthoogte & ImageCard(Left(kleur, 1))
                If Len(kleur) > 1 Then
                strcontract = strcontract & Mid(kleur, 2)
                End If
            End If
        End If
        If Door = "" Then Door = "&nbsp;"
        
        resultaat = plus_add(resultaat)
        
        If score = "" Then score = "&nbsp;"
        If ButlerScore = "" Then ButlerScore = "&nbsp;"
        HTML = HTML & nAlign("td", "center", "0") & strcontract & "</td>" & vbCr
        HTML = HTML & Align("td", "right", "0 15px") & resultaat & "</td>" & vbCr
        HTML = HTML & Align("td", "center", "0 15px") & Door & "</td>" & vbCr
        HTML = HTML & Align("td class=" & Chr(34) & "detailalignrightborderright" & Chr(34), "right", "0 15px") & score & "</td>" & vbCr
        'html = html & Align("td", "right", "0 15px") & ButlerScore & "</td>" & vbCr
        'HTML = HTML & "<td class=" & Chr(34) & "detailalignrightborderright" & Chr(34) & " >" & ButlerScore & "</td>" & vbCr
      
        HTML = HTML & "   </tr>" & vbCr
        ScoresheetRow = HTML
End Function

'Public Function ScoresheetLastRow(Boardnr As Variant, Contract As Variant, Resultaat As Variant, Door As Variant, Score As Variant, ButlerScore As Variant) As String
Public Function ScoresheetLastRow(Boardnr As Variant, Contract As Variant, resultaat As Variant, Door As Variant, score As Variant) As String

Dim HTML, contracthoogte, kleur, strcontract As String
Dim hoogte As Integer
HTML = ""

        'test op niet gespeeld of kunstmatige score
        
        
        HTML = HTML & "   <tr>" & vbCr
        HTML = HTML & "<td class=" & Chr(34) & "detailalignrightborderleftbottom" & Chr(34) & " >" & Boardnr & "</td>" & vbCr
        'ontleed contact
        
        If Contract = "" Then
            strcontract = "&nbsp;"
        Else
            hoogte = Val(Contract)
            If hoogte = 0 Then
                contracthoogte = "pass"
                kleur = ""
            Else
                contracthoogte = Left(Contract, 1)
                kleur = Mid(Contract, 2)
            End If
            '
            If InStr(kleur, "SA") > 0 Then
                    strcontract = contracthoogte & kleur
                 Else
                strcontract = contracthoogte & ImageCard(Left(kleur, 1))
                If Len(kleur) > 1 Then
                strcontract = strcontract & Mid(kleur, 2)
                End If
            End If
        End If
        If Door = "" Then Door = "&nbsp;"
        resultaat = plus_add(resultaat)
        
        If score = "" Then score = "&nbsp;"
        If ButlerScore = "" Then ButlerScore = "&nbsp;"
            HTML = HTML & "<td class=" & Chr(34) & "detailaligncenterborderbottom" & Chr(34) & " >" & strcontract & "</td>" & vbCr
           ' html = html & nAlign("td", "center", "0") & strcontract & "</td>" & vbCr
        'html = html & Align("td", "right", "0 15px") & Resultaat & "</td>" & vbCr
        HTML = HTML & "<td class=" & Chr(34) & "detailaligncenterborderbottom" & Chr(34) & " >" & resultaat & "</td>" & vbCr
 
       '  html = html & Align("td", "center", "0 15px") & Door & "</td>" & vbCr
        HTML = HTML & "<td class=" & Chr(34) & "detailaligncenterborderbottom" & Chr(34) & " >" & Door & "</td>" & vbCr
        'html = html & Align("td", "right", "0 15px") & Score & "</td>" & vbCr
       HTML = HTML & "<td class=" & Chr(34) & "detailalignrightborderrightbottom" & Chr(34) & " >" & score & "</td>" & vbCr
        
        'html = html & Align("td", "right", "0 15px") & ButlerScore & "</td>" & vbCr
        'HTML = HTML & "<td class=" & Chr(34) & "detailalignrightborderrightbottom" & Chr(34) & " >" & ButlerScore & "</td>" & vbCr
      
        HTML = HTML & "   </tr>" & vbCr
        ScoresheetLastRow = HTML
End Function

Public Function TeamResultRow(UitTeam As Variant, refUitTeam As Variant, ImpsWij As Variant, ImpsZij As Variant, VPsWij As Variant, VPsZij As Variant) As String
HTML = ""
HTML = HTML & " <tr>" & vbCr
'html = html & Align("td", "left", "0") & "<a href=" & Chr(34) & refThuisTeam & Chr(34) & ">" & ThuisTeam & "</a></td>" & vbCr
HTML = HTML & Align("td", "left", "0 15px") & "<a href=" & Chr(34) & refUitTeam & Chr(34) & ">" & UitTeam & "</a></td>" & vbCr
HTML = HTML & Align("td", "right", "0") & ImpsWij & "</td>" & vbCr
HTML = HTML & Align("td", "right", "0") & ImpsZij & "</td>" & vbCr
If VPsWij <> "&nbsp;" Then VPsWij = Format(VPsWij, "#0.00")
If VPsZij <> "&nbsp;" Then VPsZij = Format(VPsZij, "#0.00")
HTML = HTML & Align("td", "right", "0") & VPsWij & "</td>" & vbCr
HTML = HTML & Align("td", "right", "0") & VPsZij & "</td>" & vbCr
HTML = HTML & " </tr>" & vbCr
TeamResultRow = HTML
End Function

Public Function TeamNoResultRow(UitTeam As Variant, refUitTeam As Variant) As String
HTML = ""
HTML = HTML & " <tr>" & vbCr
'html = html & Align("td", "left", "0") & "<a href=" & Chr(34) & refThuisTeam & Chr(34) & ">" & ThuisTeam & "</a></td>" & vbCr
HTML = HTML & Align("td", "left", "0 15px") & "<a href=" & Chr(34) & refUitTeam & Chr(34) & ">" & UitTeam & "</a></td>" & vbCr
HTML = HTML & Align("td", "right", "0") & "&nbsp;" & "</td>" & vbCr
HTML = HTML & Align("td", "right", "0") & "&nbsp;" & "</td>" & vbCr

HTML = HTML & Align("td", "right", "0") & "&nbsp;" & "</td>" & vbCr
HTML = HTML & Align("td", "right", "0") & "&nbsp;" & "</td>" & vbCr
HTML = HTML & " </tr>" & vbCr
TeamNoResultRow = HTML
End Function


Public Function TeamUitslagenResultRow(ThuisTeam As Variant, refThuisTeam As Variant, UitTeam As Variant, refUitTeam As Variant, ImpsThuis As Variant, ImpsUit As Variant, VpsThuis As Variant, VpsUit As Variant) As String
HTML = ""
HTML = HTML & " <tr>" & vbCr
HTML = HTML & Align("td", "left", "0 15px") & "<a href=" & Chr(34) & refThuisTeam & Chr(34) & ">" & ThuisTeam & "</a></td>" & vbCr
HTML = HTML & Align("td", "left", "0 15px") & "<a href=" & Chr(34) & refUitTeam & Chr(34) & ">" & UitTeam & "</a></td>" & vbCr
HTML = HTML & Align("td", "right", "0") & ImpsThuis & "</td>" & vbCr
HTML = HTML & Align("td", "right", "0") & ImpsUit & "</td>" & vbCr
If VpsThuis <> "&nbsp;" Then VpsThuis = Format(VpsThuis, "#0.00")
If VpsUit <> "&nbsp;" Then VpsUit = Format(VpsUit, "#0.00")
HTML = HTML & Align("td", "right", "0") & VpsThuis & "</td>" & vbCr
HTML = HTML & Align("td", "right", "0") & VpsUit & "</td>" & vbCr
HTML = HTML & " </tr>" & vbCr
TeamUitslagenResultRow = HTML
End Function

Public Function TeamUnderlineUitslagenResultRow(ThuisTeam As Variant, refThuisTeam As Variant, UitTeam As Variant, refUitTeam As Variant, ImpsThuis As Variant, ImpsUit As Variant, VpsThuis As Variant, VpsUit As Variant) As String
HTML = ""
HTML = HTML & " <tr class=" & "data-3" & ">" & vbCr
HTML = HTML & Align("td", "left", "0 15px") & "<a href=" & Chr(34) & refThuisTeam & Chr(34) & ">" & ThuisTeam & "</a></td>" & vbCr
HTML = HTML & Align("td", "left", "0 15px") & "<a href=" & Chr(34) & refUitTeam & Chr(34) & ">" & UitTeam & "</a></td>" & vbCr
HTML = HTML & Align("td", "right", "0") & ImpsThuis & "</td>" & vbCr
HTML = HTML & Align("td", "right", "0") & ImpsUit & "</td>" & vbCr
If VpsThuis <> "&nbsp;" Then VpsThuis = Format(VpsThuis, "#0.00")
If VpsUit <> "&nbsp;" Then VpsUit = Format(VpsUit, "#0.00")
HTML = HTML & Align("td", "right", "0") & VpsThuis & "</td>" & vbCr
HTML = HTML & Align("td", "right", "0") & VpsUit & "</td>" & vbCr
HTML = HTML & " </tr>" & vbCr
TeamUnderlineUitslagenResultRow = HTML
End Function

Public Function TeamUitslagenNoResultRow(ThuisTeam As Variant, refThuisTeam As Variant, UitTeam As Variant, refUitTeam As Variant) As String
HTML = ""
HTML = HTML & " <tr>" & vbCr
HTML = HTML & Align("td", "left", "0 15px") & "<a href=" & Chr(34) & refThuisTeam & Chr(34) & ">" & ThuisTeam & "</a></td>" & vbCr
HTML = HTML & Align("td", "left", "0 15px") & "<a href=" & Chr(34) & refUitTeam & Chr(34) & ">" & UitTeam & "</a></td>" & vbCr
HTML = HTML & Align("td", "right", "0") & "&nbsp;" & "</td>" & vbCr
HTML = HTML & Align("td", "right", "0") & "&nbsp;" & "</td>" & vbCr
HTML = HTML & Align("td", "right", "0") & "&nbsp;" & "</td>" & vbCr
HTML = HTML & Align("td", "right", "0") & "&nbsp;" & "</td>" & vbCr
HTML = HTML & " </tr>" & vbCr
TeamUitslagenNoResultRow = HTML
End Function
Public Function SaldoRow(saldo As Variant, imps As Variant, ImpsWe As Variant, ImpsThey As Variant) As String
Dim HTML As String
        HTML = ""
        HTML = HTML & "   <tr>" & vbCr

        If saldo = "" Then saldo = "&nbsp;"
        If imps = "" Then imps = "&nbsp;"
        If ImpsWe = "" Then ImpsWe = "&nbsp;"
        If ImpsThey = "" Then ImpsThey = "&nbsp;"
        
        HTML = HTML & "<td class=" & Chr(34) & "detailalignrightborderleft" & Chr(34) & " >" & saldo & "</td>" & vbCr
       ' html = html & Align("td", "right", "0 15px") & Saldo & "</td>" & vbCr
        HTML = HTML & Align("td", "right", "0 15px") & imps & "</td>" & vbCr
       
        HTML = HTML & Align("td", "right", "0 15px") & ImpsWe & "</td>" & vbCr
       ' html = html & Align("td", "right", "0 15px") & ImpsThey & "</td>" & vbCr
         HTML = HTML & "<td class=" & Chr(34) & "detailalignrightborderright" & Chr(34) & " >" & ImpsThey & "</td>" & vbCr
      
        HTML = HTML & "   </tr>" & vbCr
        SaldoRow = HTML
        
End Function
Public Function SaldoLastRow(saldo As Variant, imps As Variant, ImpsWe As Variant, ImpsThey As Variant) As String
Dim HTML As String
        HTML = ""
        HTML = HTML & "   <tr>" & vbCr

        If saldo = "" Then saldo = "&nbsp;"
        If imps = "" Then imps = "&nbsp;"
        If ImpsWe = "" Then ImpsWe = "&nbsp;"
        If ImpsThey = "" Then ImpsThey = "&nbsp;"
        
        HTML = HTML & "<td class=" & Chr(34) & "detailalignrightborderleftbottom" & Chr(34) & " >" & saldo & "</td>" & vbCr
       ' html = html & Align("td", "right", "0 15px") & Saldo & "</td>" & vbCr
        'html = html & Align("td", "right", "0 15px") & Imps & "</td>" & vbCr
        HTML = HTML & "<td class=" & Chr(34) & "detailalignrightborderbottom" & Chr(34) & " >" & imps & "</td>" & vbCr
        'html = html & Align("td", "right", "0 15px") & ImpsWe & "</td>" & vbCr
        HTML = HTML & "<td class=" & Chr(34) & "detailalignrightborderbottom" & Chr(34) & " >" & ImpsWe & "</td>" & vbCr
       ' html = html & Align("td", "right", "0 15px") & ImpsThey & "</td>" & vbCr
         HTML = HTML & "<td class=" & Chr(34) & "detailalignrightborderrightbottom" & Chr(34) & " >" & ImpsThey & "</td>" & vbCr
      
        HTML = HTML & "   </tr>" & vbCr
        SaldoLastRow = HTML
        
End Function


Public Function Scoresheetheader(Paar1 As Variant, RefPaar1 As Variant) As String
Dim HTML As String
    HTML = ""
    HTML = HTML & Align("td", "left", "") & vbCr
    HTML = HTML & "<table class=" & Chr(34) & "results" & Chr(34) & ">" & vbCr
    HTML = HTML & " <caption> " & "<a href=" & Chr(34) & RefPaar1 & Chr(34) & ">" & Paar1 & "</a></caption> " & vbCr
    
    HTML = HTML & "  <thead>" & vbCr
    HTML = HTML & "  <tr>" & vbCr
    HTML = HTML & "  <th class=" & Chr(34) & "resultsheader" & Chr(34) & ">Spel</th>" & vbCr
   ' html = html & "  <th class=" & Chr(34) & "resultsheader" & Chr(34) & " colspan=" & Chr(34) & "2" & Chr(34) & ">Contract</th>" & vbCr
    HTML = HTML & "  <th class=" & Chr(34) & "resultsheader" & Chr(34) & ">Contract</th>" & vbCr
    HTML = HTML & "  <th class=" & Chr(34) & "resultsheader" & Chr(34) & ">+/-</th>" & vbCr
    HTML = HTML & "  <th class=" & Chr(34) & "resultsheader" & Chr(34) & ">Door</th>" & vbCr
    HTML = HTML & "  <th class=" & Chr(34) & "resultsheader" & Chr(34) & ">Score</th>" & vbCr
    'HTML = HTML & "  <th class=" & Chr(34) & "resultsheader" & Chr(34) & ">Resultaat</th>" & vbCr
    HTML = HTML & "  </tr>" & vbCr
    HTML = HTML & "  </thead>" & vbCr
    HTML = HTML & "  <tbody>" & vbCr
    Scoresheetheader = HTML
End Function


Public Function ScoreSaldoheader() As String
Dim HTML As String
    HTML = ""
    HTML = HTML & Align("td", "center", "") & vbCr
    HTML = HTML & "<table class=" & Chr(34) & "results" & Chr(34) & ">" & vbCr
    HTML = HTML & " <caption>M.P.</caption> " & vbCr
    HTML = HTML & "  <thead>" & vbCr
    HTML = HTML & "  <tr>" & vbCr
    HTML = HTML & "  <th class=" & Chr(34) & "resultsheader" & Chr(34) & ">Saldo</th>" & vbCr
    HTML = HTML & "  <th class=" & Chr(34) & "resultsheader" & Chr(34) & ">imps</th>" & vbCr
    HTML = HTML & "  <th class=" & Chr(34) & "resultsheader" & Chr(34) & ">imps wij</th>" & vbCr
    HTML = HTML & "  <th class=" & Chr(34) & "resultsheader" & Chr(34) & ">imps zij</th>" & vbCr
    HTML = HTML & "  </tr>" & vbCr
    HTML = HTML & "  </thead>" & vbCr
    HTML = HTML & "  <tbody>" & vbCr
    ScoreSaldoheader = HTML
End Function



Public Function Scorefooter() As String
Dim HTML As String
     HTML = ""
     HTML = HTML & "  </tbody>" & vbCr
     HTML = HTML & " </table>" & vbCr
     HTML = HTML & "</td>" & vbCr
     Scorefooter = HTML
End Function

Public Function TeamResultheader() As String
    Dim hmtl As String
    HTML = ""
    HTML = HTML & "<table class=" & Chr(34) & "data" & Chr(34) & ">" & vbCr
    HTML = HTML & "<thead>" & vbCr
    HTML = HTML & " <tr>" & vbCr
    'html = html & "   " & Align("th class=" & Chr(34) & "data" & Chr(34), "left", "") & "Wij</th>" & vbCr
    HTML = HTML & "   " & Align("th class=" & Chr(34) & "data" & Chr(34), "left", "0 15px") & "Tegen</th>" & vbCr
    HTML = HTML & "   " & Align("th class=" & Chr(34) & "data" & Chr(34), "right", "") & "Wij Imps</th>" & vbCr
    HTML = HTML & "   " & Align("th class=" & Chr(34) & "data" & Chr(34), "right", "") & "Zij Imps</th>" & vbCr
    HTML = HTML & "   " & Align("th class=" & Chr(34) & "data" & Chr(34), "right", "") & "Wij VPs</th>" & vbCr
    HTML = HTML & "   " & Align("th class=" & Chr(34) & "data" & Chr(34), "right", "") & "Zij VPs</th>" & vbCr
    HTML = HTML & " </tr>" & vbCr
    HTML = HTML & "</thead>" & vbCr
    HTML = HTML & "<tbody>" & vbCr
    TeamResultheader = HTML
                              
End Function

Public Function TeamUitslagenResultheader() As String
    Dim hmtl As String
    HTML = "<tr><td>"
    HTML = HTML & "<table class=" & Chr(34) & "data" & Chr(34) & ">" & vbCr
    HTML = HTML & "<thead>" & vbCr
    HTML = HTML & " <tr>" & vbCr
    HTML = HTML & "   " & Align("th class=" & Chr(34) & "data" & Chr(34), "left", "") & "Thuis</th>" & vbCr
    HTML = HTML & "   " & Align("th class=" & Chr(34) & "data" & Chr(34), "left", "") & "Uit</th>" & vbCr
    HTML = HTML & "   " & Align("th class=" & Chr(34) & "data" & Chr(34), "right", "") & "Thuis Imps</th>" & vbCr
    HTML = HTML & "   " & Align("th class=" & Chr(34) & "data" & Chr(34), "right", "") & "Uit Imps</th>" & vbCr
    HTML = HTML & "   " & Align("th class=" & Chr(34) & "data" & Chr(34), "right", "") & "Thuis VPs</th>" & vbCr
    HTML = HTML & "   " & Align("th class=" & Chr(34) & "data" & Chr(34), "right", "") & "Uit VPs</th>" & vbCr
    HTML = HTML & " </tr>" & vbCr
    HTML = HTML & "</thead>" & vbCr
    HTML = HTML & "<tbody>" & vbCr
    TeamUitslagenResultheader = HTML
                              
End Function


Public Function Teamkruisheader(Kolom() As Variant) As String
Dim hmtl As String
Dim teller As Integer
    HTML = ""
    HTML = HTML & "   <tr>" & vbCr
    HTML = HTML & "   <td>" & vbCr
    HTML = HTML & "<table class=" & Chr(34) & "data-rotate " & Chr(34) & ">" & vbCr
    HTML = HTML & "<thead>" & vbCr
    HTML = HTML & " <tr>" & vbCr
    
   HTML = HTML & "  <th>" & Kolom(1) & "</th>" & vbCr
   For teller = 2 To AANTALTEAMS + 1
    HTML = HTML & "  <th><div class=" & Chr(34) & "data-rotate-header-container" & Chr(34) & ">"
    HTML = HTML & "<div class=" & Chr(34) & "data-rotate-header-content" & Chr(34) & ">" & Kolom(teller)
    HTML = HTML & "</div></div></th>" & vbCr
   Next
   For teller = AANTALTEAMS + 2 To AANTALTEAMS + 4
    HTML = HTML & "  <th>" & Kolom(teller) & "</th>" & vbCr
   Next
    HTML = HTML & " </tr>" & vbCr
    HTML = HTML & "</thead>" & vbCr
    HTML = HTML & "<tbody>" & vbCr
    Teamkruisheader = HTML
    
End Function
Public Function Teamkruisfooter() As String
    Dim HTML As String
    HTML = ""
    HTML = HTML & "   </td>" & vbCr
    HTML = HTML & "   </tr>" & vbCr
    HTML = HTML & "   </table>" & vbCr
   
    Teamkruisfooter = HTML
    
End Function
   
    
    
Public Function Teamkruisheaderrow(Kolom() As Variant, Teamnr As Variant) As String
Dim HTML As String
Dim teller As Integer
Dim b_href, e_href As String
Dim info As Integer
Dim Tegenstander As Integer
Dim avond As Integer
e_href = "</a>"
HTML = ""
HTML = HTML & " <tr>" & vbCr
HTML = HTML & Align("td style=" & Chr(34) & "background-color: #BEC8D1;" & Chr(34), "left", "") & Kolom(1) & "</td>" & vbCr
For teller = 2 To AANTALTEAMS + 1

    If Kolom(teller) = "" Then
    HTML = HTML & Align("td", "center", "") & "&nbsp;" & "</td>" & vbCr
    Else
     If Kolom(teller) = "xxx" Then
        HTML = HTML & Align("td style=" & Chr(34) & "background-color: #cfe692;" & Chr(34), "center", "") & Kolom(teller) & "</td>" & vbCr
     Else
        ' achtergrond
        Tegenstander = teller - 1
        avond = Team_Tegen_Avond(Teamnr, Tegenstander)
        info = WebInfo(avond)
        b_href = "<a href=" & Chr(34) & LOCALSITE & info & "/" & PREFIX & avond & "_Teamnr_" & Teamnr & ".html" & Chr(34) & ">"
        If TEAMBYE > 0 And teller = TEAMBYE + 1 Then
            HTML = HTML & Align("td style=" & Chr(34) & "background-color: #b7c5b7;" & Chr(34), "center", "") & b_href & "&nbsp;---" & e_href & "</td>" & vbCr
        Else
          ' Tgenstander = teller -1
          HTML = HTML & Align("td style=" & Chr(34) & "background-color: #b7c5b7;" & Chr(34), "center", "") & b_href & Format(Kolom(teller), "#0.00") & e_href & "</td>" & vbCr
        End If
      End If
    End If
Next

HTML = HTML & Align("td", "right", "") & Format(Kolom(AANTALTEAMS + 2), "#0.00") & "</td>" & vbCr
HTML = HTML & Align("td", "right", "") & Format(Kolom(AANTALTEAMS + 3), "#0.00") & "</td>" & vbCr
HTML = HTML & Align("td", "center", "") & Kolom(AANTALTEAMS + 4) & "</td>" & vbCr
HTML = HTML & "   </tr>" & vbCr
Teamkruisheaderrow = HTML
End Function
Public Function Byekruisheaderrow(Kolom() As Variant) As String
Dim HTML As String
Dim teller As Integer
HTML = ""
HTML = HTML & " <tr>" & vbCr
HTML = HTML & Align("td style=" & Chr(34) & "background-color: #BEC8D1;" & Chr(34), "left", "") & Kolom(1) & "</td>" & vbCr
For teller = 2 To AANTALTEAMS + 1
    If Kolom(teller) = "" Then
    HTML = HTML & Align("td", "center", "") & "&nbsp;" & "</td>" & vbCr
    Else
     If Kolom(teller) = "xxx" Then
        HTML = HTML & Align("td style=" & Chr(34) & "background-color: #cfe692;" & Chr(34), "center", "") & Kolom(teller) & "&nbsp;" & "&nbsp;" & "</td>" & vbCr
     Else
        ' achtergrond
        HTML = HTML & Align("td style=" & Chr(34) & "background-color: #b7c5b7;" & Chr(34), "center", "") & "&nbsp;---" & "</td>" & vbCr
       End If
    End If
Next

HTML = HTML & Align("td", "right", "") & Format(Kolom(AANTALTEAMS + 2), "#0.00") & "</td>" & vbCr
HTML = HTML & Align("td", "right", "") & Format(Kolom(AANTALTEAMS + 3), "#0.00") & "</td>" & vbCr
HTML = HTML & Align("td", "center", "") & Kolom(AANTALTEAMS + 4) & "</td>" & vbCr
HTML = HTML & "   </tr>" & vbCr
Byekruisheaderrow = HTML
End Function

Public Function TeamResultfooter() As String
Dim HTML, Voetje, Linkje As String
    HTML = ""
    If Voettekst = "" Then
        Voetje = "www.pwobridge.nl"
        Else
        Voetje = Voettekst
    End If
      If Voetlink = "" Then
        Linkje = "#"
        Else
        Linkje = Voetlink
      End If

    HTML = HTML & "</tbody>" & vbCr
    HTML = HTML & "</table>" & vbCr
    HTML = HTML & "       <tr>"
    HTML = HTML & "        <td class=" & Chr(34) & "footer" & Chr(34) & "><a class=" & Chr(34) & "footer" & Chr(34) & " href=" & Chr(34) & Linkje & Chr(34) & ">" & Voetje & "</a></td> "
    HTML = HTML & "    </tr>"
 
   
    TeamResultfooter = HTML
End Function
Public Function TeamUitslagenResultfooter() As String
Dim HTML As String
    HTML = ""
    HTML = HTML & "</tbody>" & vbCr
    HTML = HTML & "</table>"  'data
    HTML = HTML & "</td></tr>"
    HTML = HTML & "       <tr>"
    HTML = HTML & "        <td class=" & Chr(34) & "footer" & Chr(34) & "><a class=" & Chr(34) & "footer" & Chr(34) & " href=" & Chr(34) & "http://www.pwobridge.nl/" & Chr(34) & ">www.pwobridge.nl</a></td> "
    HTML = HTML & "    </tr>"
    HTML = HTML & "</table>" & vbCr 'pagefram
    HTML = HTML & "</div>"
    TeamUitslagenResultfooter = HTML
End Function