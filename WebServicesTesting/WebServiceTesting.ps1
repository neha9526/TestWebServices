$cls
$scriptPath= split-path -parent $MyInvocation.MyCommand.Definition
. $scriptPath\utilities.ps1

[string]$CSS= @'
    <style>
    h1, h5, th { text-align: center; }
table { margin: auto; font-family: Segoe UI; box-shadow: 10px 10px 5px #888; border: thin ridge grey; }
th { background: #0046c3; color: #fff; max-width: 400px; padding: 5px 10px; }
td { font-size: 11px; padding: 5px 20px; color: #000; }
tr { background: #b8d1f3; }
tr:nth-child(even) { background: #dae5f4; }
tr:nth-child(odd) { background: #b8d1f3; }
  
 </style>
'@

#------------------Function to edit soap message accordingly with the datafile------------
function XmlEdit ([xml]$SoapMessage, [array]$datarow)
{  
try{  
foreach ($d in $array ){
            if($datarow.$d -ne ""){
  
                    $xpathfilter = "//*[local-name()='$d']"
                    $node = Select-Xml -Xml $SoapMessage -XPath $xpathfilter
                    $node.Node.InnerXml = $datarow.$d   
                    
                }
            }
  
  

  return ,$SoapMessage.InnerXml
 }
  catch [System.Exception]
  {
  Write-Host "Error while Editing Xml"

  }

}


#----------------Function to send webrequest with edited Soap Envelope------------------------

function WebService([string]$URl, [string]$ContentType, [xml]$body ,[array]$data , $path, $headers )
{
try{
$resultXml = iwr $URl -ContentType  $ContentType -Body $body -OutFile $path -Method Post -Headers $headers -TimeoutSec 1800
$resultXml.content
Write-Host $data.Type  $data.Web_Service_Version  "is Succesfull"
$Global:testStatus = "Success" 
$Global:HttpStatusCode=200
}
catch [System.Net.WebException] {
        
        $Global:testStatus = "Unsuccessful "
        Write-Host  $data.Type  $data.Web_Service_Version  "is UnSuccesfull"

        Write-Host $_.Exception
        $Global:HttpStatusCode=[int]$_.Exception.Response.StatusCode
        $ErrorMessage1=$_.Exception.Message 
        $request = $_.Exception.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($request)
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $responseBody = $reader.ReadToEnd();
        $responseBody
        

        $xmlException = [xml]$responseBody
        $Error1 = Select-Xml -xml $xmlException -XPath "//*[local-name()='faultcode']"
        $ErrorMessage2 = $Error1.Node.InnerText
        $Error1= Select-Xml -xml $xmlException -XPath "//*[local-name()='faultstring']"
        $ErrorMessage3 = ($Error1.Node.InnerXml).Replace("`r`n","")
        
        $Global:ErrorMessage =$ErrorMessage1 
        #--- + "`b" +$ErrorMessage2+"`b"+$ErrorMessage3[1] + "`b"      
        $Global:ErrorMessage
        

}


}



cls
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$dataFile= Import-Csv -Path ($scriptPath + '\DataStore\dataFileAdd.csv')

$contentType="text/xml"

$Global:HttpStatusCode
$Global:testStatus
$Global:ErrorMessage 
$header = "Web Service Action"+ "," +"intA"+"," + "intB"+ "," +"HttpStatusCode"+ "," +"TestResult" +","+"Error Message"

$array= ("intA","intB")

[xml] $wsdl = Get-Content  -Path (Join-Path $scriptPath '\lib\AddService.xml' )

$LogTime = Get-Date -Format "yyyy-MM-dd-HH-mm-ss"
$reportLogcsv =$scriptPath + "\Reports\report"+"_"+ $LogTime +".csv"
$reportLoghtml =$scriptPath + "\Reports\report"+"_"+ $LogTime +".html"
New-Item $reportLogcsv -type file
Add-Content -Path $reportLogcsv -Value $header

ForEach ($data in $dataFile ) 
{   
         if($data.Skip -ne ""){
            Continue  #Skipping if Skip Column has Some Value
         }

         $URI_Service="http://www.dneonline.com/calculator.asmx"
         $LogTime = Get-Date -Format "yyyy-MM-dd-HH-mm-ss"
         $OutputPath =$scriptPath +"\ResponseEnvelopes\" +$data.Web_Service_Action + "_" + $LogTime +$random + ".xml"
         New-Item $OutputPath -type file
         
         [xml]  $currentSoapMessage =  ( XmlEdit -SoapMessage $wsdl -datarow $data) 
         WebService -URl $URI_Service -ContentType $contentType -body $currentSoapMessage -data $data -path $OutputPath

        
         #Reporting
        $result = $data.Web_Service_Action + "," +$data.intA +","+$data.intB+ "," +$HttpStatusCode +"," + $testStatus+","+$Global:ErrorMessage
        Add-Content -Path $reportLogcsv -Value $result
        $Resultshtml = Import-Csv $reportLogcsv | ConvertTo-Html -Title "Reports"  -Head $css -body "<H2 align=""center"">WebService Test Reports</H2>" |foreach {
        $PSItem -replace "<td>Fail</td>", "<td style='background-color:red'>Fail</td>"}
        $Resultshtml =(($Resultshtml.Replace("&lt;","<")).Replace("&gt;",">")).Replace("&quot;",'"') >  "$reportLoghtml"
  
   # --- reset all the Values
        $Global:HttpStatusCode=""
        $Global:testStatus=""
        $Global:ErrorMessage=""
        

}




 