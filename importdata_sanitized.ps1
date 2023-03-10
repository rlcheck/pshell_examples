# 6/16/2018 by RLH
#


#check to see if powershell 4+ is installed.
$build=$PSVersionTable.PSVersion
if(($build -match 4) -or ($build -match 5))
	{
	Write-Host powershell $build is installed, proceeding..
	}
else {exit}




########################FUNCTIONS#############################
##############################################################

$functions = {

#set debug
$debuglvl = 0

#This contains all the functions we use for jobs
	Function LogWrite
	{
	   Param ([string]$logstring)
	   $Logfile = "D:\Program Files\PS\$(gc env:computername).log"
	   $tm = get-date
	   $totalstring = "$tm"+"-$logstring"
	   Add-content $Logfile -value $totalstring
	}

	
	function Invoke-MySQL {
	#Using mysql ado.net connector to provide fastest link to local mariadb, 32 bit connector installed because i will need to run from 32 bit ps because of office connectors
	Param(
	  [Parameter(
	  Mandatory = $true,
	  ParameterSetName = '',
	  ValueFromPipeline = $true)]
	  [string]$Query
	  )
	$MySQLAdminUserName = 'stuff'
	$MySQLAdminPassword = 'stuff'
	$MySQLDatabase = 'data'
	$MySQLHost = '127.0.0.1'
	$ConnectionString = "server=" + $MySQLHost + "; port=3306; uid=" + $MySQLAdminUserName + "; pwd=" + $MySQLAdminPassword + "; database="+$MySQLDatabase
	Try {
	  [void][System.Reflection.Assembly]::LoadWithPartialName("MySql.Data")
	  $Connection = New-Object MySql.Data.MySqlClient.MySqlConnection
	  $Connection.ConnectionString = $ConnectionString
	  $Connection.Open()
	  $Command = New-Object MySql.Data.MySqlClient.MySqlCommand($Query, $Connection)
	  $DataAdapter = New-Object MySql.Data.MySqlClient.MySqlDataAdapter($Command)
	  $DataSet = New-Object System.Data.DataSet
	  $RecordCount = $dataAdapter.Fill($dataSet, "data")
	  $DataSet.Tables[0]
	  }
	Catch {
	  Logwrite "ERROR : UNABLE TO RUN MYSQL-QUERY : $query `n$Error[0]"
	  Write-host "ERROR : UNABLE TO RUN MYSQL-QUERY : $query `n$Error[0]"
	  $Error.Clear()
	 }
	Finally {
	  $Connection.Close()
	  }
	 }

	 
	function Invoke-Sybase {
	#Using ODBC connector setup as a DSN because sybase doesnt support ado.net at version 10
	Param(
	  [Parameter(
	  Mandatory = $true,
	  ParameterSetName = '',
	  ValueFromPipeline = $true)]
	  [string]$Query
	  )
	$ConnectionString ="DSN=SERVER;Uid=stuff;Pwd=stuff;"
	Try {
	  $Connection = New-Object System.Data.Odbc.OdbcConnection
	  $Connection.ConnectionString = $ConnectionString
	  $Connection.Open()
	  $Command = New-Object System.Data.Odbc.OdbcCommand($Query, $Connection)
	  $DataAdapter = New-Object system.Data.odbc.odbcDataAdapter($Command)
	  $DataSet = New-Object System.Data.DataSet
	  [void]$DataAdapter.fill($DataSet)
	  $DataSet.Tables[0]
	  }
	Catch {
	  LogWrite "ERROR : UNABLE TO RUN SYBASE-QUERY : $query `n$Error[0]"
	  Write-host "ERROR : UNABLE TO RUN SYBASE-QUERY : $query `n$Error[0]"
	  $Error.Clear()
	 }
	Finally {
	  $Connection.Close()
	  }
	 }

	 
	function Invoke-MSSQL {
	#Using MSSQL built in connector
	Param(
	  [Parameter(
	  Mandatory = $true,
	  ParameterSetName = '',
	  ValueFromPipeline = $true)]
	  [string]$Query
	  )
	$ConnectionString = "Server=server.domain.local;Database=stuff;User Id=script;Password=stuff;"
	Try {
	  $Connection = New-Object System.Data.SqlClient.SqlConnection
	  $Connection.ConnectionString = $ConnectionString
	  $Connection.Open()
	  $Command = New-Object System.Data.SqlClient.SqlCommand($Query, $Connection)
      $Command.CommandTimeout = 90
	  $DataAdapter = New-Object System.Data.SqlClient.SqlDataAdapter($Command)
	  $DataSet = New-Object System.Data.DataSet
	  $RecordCount = $dataAdapter.Fill($dataSet, "data")
	  $DataSet.Tables[0]
	  }
	Catch {
	  Logwrite "ERROR : UNABLE TO RUN MSSQL-QUERY : $query `n$Error[0]"
	  Write-host "ERROR : UNABLE TO RUN MSSQL-QUERY : $query `n$Error[0]"
	  $Error.Clear()
	 }
	Finally {
	  $Connection.Close()
	  }
	 }

	 
	function Invoke-ACCESS {  
	  param(  
	  [Parameter(Mandatory=$true)]  
	  $sql = "select * from ProductReport",  
	  [Parameter(Mandatory=$true)]  
	  $connectionstring
	  ) 
	Try{
	  $db = New-Object -comObject ADODB.Connection  
	  $db.Open($connectionstring)  
	  $rs = $db.Execute($sql)  
	  while (!$rs.EOF) {  
		$hash = @{}  
		foreach ($field in $rs.Fields) {  
		  $hash.$($field.Name) = $field.Value  
		}  
		$rs.MoveNext()      
		New-Object PSObject -property $hash  
	  }
	  $rs.Close()  
	  $db.Close()  
	}
	Catch {
	  Logwrite "ERROR : UNABLE TO RUN ACCESS-QUERY : $query `n$Error[0]"
	  Write-host "ERROR : UNABLE TO RUN ACCESS-QUERY : $query `n$Error[0]"
	  $Error.Clear()
	 }
	Finally {
	  #Logwrite "The Invoke Access function tried to run"
      $rs.Close()  
	  $db.Close()  
	  }
	  
	} 


	function Import-EXCEL 
	{ 
	
	Param( 
			[parameter( 
				mandatory=$true,  
				position=1,  
				ValueFromPipeline=$true,  
				ValueFromPipelineByPropertyName=$true)] 
			[String[]] 
			$FilePath, 
			[parameter(mandatory=$false)] 
			$SheetName = 1, 			 
			[parameter(mandatory=$false)] 
			$StartRow = 1, 
			[parameter(mandatory=$false)] 
			$EndRow = 2, 
			[parameter(mandatory=$false)] 
			$ColOne = "", 
			[parameter(mandatory=$false)] 
			$ColTwo = "", 			
			[parameter(mandatory=$false)] 
			$ColThree ="", 
			[parameter(mandatory=$false)] 
			$ColFour = ""
		) 
	try{
	#Create an Object Excel.Application using Com interface we have to sort out what excel applications are running beforehand to kill it properly once done
	$before = @(Get-Process [e]xcel | %{$_.Id})
	$objExcel = New-Object -ComObject Excel.Application
	$ExcelId = Get-Process excel | %{$_.Id} | ?{$before -notcontains $_}
	# Disable the 'visible' property so the document won't open in excel
	$objExcel.Visible = $false
	# Open the Excel file and save it in $WorkBook
	$WorkBook = $objExcel.Workbooks.Open($FilePath)
	# Load the WorkSheet 'whatever'
	$WorkSheet = $WorkBook.sheets.item($SheetName)
	$CurrRow = $StartRow
	$oneresults = @()
	$tworesults = @()
	$threeresults = @()
	$fourresults = @()
	While ($CurrRow -le $Endrow){
	if($ColOne -ne ""){$oneresults += $worksheet.Range("$ColOne$CurrRow").Text}
	if($ColTwo -ne ""){$tworesults += $worksheet.Range("$ColTwo$CurrRow").Text}
	if($ColThree -ne ""){$threeresults += $worksheet.Range("$ColThree$CurrRow").Text}
	if($ColFour -ne ""){$fourresults += $worksheet.Range("$ColFour$CurrRow").Text}
	$CurrRow = $CurrRow + 1
	}
	return $oneresults, $tworesults, $threeresults, $fourresults
	}
	Catch {
	  Logwrite "ERROR : UNABLE TO RUN IMPORT-EXCEL :  `n$Error[0]"
	  Write-host "ERROR : UNABLE TO RUN IMPORT-EXCEL : `n$Error[0]"
	  $Error.Clear()
	 }
	 Finally {
	 #all this is necesary to kill opened com objects
	$objExcel.quit()
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel)
	Stop-Process -Id $ExcelId -Force -ErrorAction SilentlyContinue
	 }
	}	


}

########################END OF FUNCTIONS######################
##############################################################



#########################################################
#########  START OF SCRIPTBLOCK TWO######################
#########################################################

$scriptblockTwo = {
$sblock=Measure-Command {

######FOSS DATA #####

# BENTLEY DATA GATHERING

$sw=Measure-Command {
try{
$vatslist = get-childitem -path "\\bentleypc\Plant Data\VATS*.csv" | where-object {$_.LastWriteTime -gt (get-date).AddDays(-7)}
$plant1list = get-childitem -path "\\bentleypc\Plant Data\PLANT1*.csv" | where-object {$_.LastWriteTime -gt (get-date).AddDays(-7)}
$wheyfatlist = get-childitem -path "\\bentleypc\Plant Data\WHEYFAT*.csv" | where-object {$_.LastWriteTime -gt (get-date).AddDays(-7)}
Remove-Item  "D:\Program Files\DATA\IMPORT\bentvat.csv" -Force
Remove-Item  "D:\Program Files\DATA\IMPORT\bentplant1.csv" -Force
Remove-Item  "D:\Program Files\DATA\IMPORT\bentwheyfat.csv" -Force
Foreach ($v in $vatslist)
	{
	    
        #write-host $v
		Import-CSV $v | export-csv -append "D:\Program Files\DATA\IMPORT\bentvat.csv" -NoTypeInformation
		
	}

Foreach ($v in $plant1list)
	{
	    
        #write-host $v
		Import-CSV $v | export-csv -append "D:\Program Files\DATA\IMPORT\bentplant1.csv" -NoTypeInformation
		
	}

Foreach ($v in $wheyfatlist)
	{
	    
        #write-host $v
		Import-CSV $v | export-csv -append "D:\Program Files\DATA\IMPORT\bentwheyfat.csv" -NoTypeInformation
		
	}

}


Catch {
  Logwrite "ERROR : UNABLE TO RUN SECTION BENTLEY UPLOAD : $query `n$Error[0]"
  Write-host "ERROR : UNABLE TO RUN SECTION BENTLEY UPLOAD: $query `n$Error[0]"
  $Error.Clear()
 }
Finally {
  if($debuglvl -gt 0){Logwrite "BENTLEY upload update ran.."  }
  }
}
Write-host "Total time for bentley transform and feed is $($sw.Totalseconds)"
if($debuglvl -gt 0){LogWrite "Total time for bentley transform and feed is $($sw.Totalseconds)"}



# Pull and feed dsi lab results for last 14 days

##try{
##$lr = Invoke-MSSQL -Query "SELECT DISTINCT pim.CompanyNumber, pim.MfgSequenceNo, pim.ProductionStartDate, rslt.TestScoreSgt, rslt.TestResult, rslt.LabTestSgt, rslt.LastUpdateTime, lab.ColumnHeading FROM dbo.NBPProductionWithItemsMade AS pim LEFT OUTER JOIN dbo.NIFTestScoreResult AS rslt ON pim.CompanyNumber = rslt.CompanyNumber AND pim.TestScoreSgt = rslt.TestScoreSgt LEFT OUTER JOIN dbo.NIFLabTest AS lab ON rslt.LabTestSgt = lab.LabTestSgt WHERE (ProductionStartDate > DATEADD(day, -30, GETDATE()) AND rslt.LabTestSgt = 1) OR (ProductionStartDate > DATEADD(day, -30, GETDATE()) AND rslt.LabTestSgt = 3) OR (ProductionStartDate > DATEADD(day, -30, GETDATE()) AND rslt.LabTestSgt = 29) OR (ProductionStartDate > DATEADD(day, -30, GETDATE()) AND rslt.LabTestSgt = 4) OR (ProductionStartDate > DATEADD(day, -30, GETDATE()) AND rslt.LabTestSgt = 5) OR (ProductionStartDate > DATEADD(day, -30, GETDATE()) AND rslt.LabTestSgt = 10) ORDER BY ProductionStartDate ASC, MfgSequenceNo ASC, ColumnHeading ASC, rslt.LastUpdateTime DESC"
##Invoke-Mysql -Query "drop table if exists dsilabresults"
##Invoke-Mysql -Query "CREATE TABLE IF NOT EXISTS dsilabresults (
##  LABID INT NOT NULL AUTO_INCREMENT,
##  CompanyNumber INT(2) DEFAULT NULL,
##  VAT VARCHAR (8) DEFAULT NULL,
##  ProductionStartDate DATE DEFAULT NULL,
##  TestScoreSgt INT (10) DEFAULT NULL,
##  TestResult DECIMAL(4,2) unsigned DEFAULT NULL,
##  LabTestSgt INT (3) DEFAULT NULL,
##  LastUpdateTime DATETIME DEFAULT NULL,
##  ColumnHeading VARCHAR (25) DEFAULT NULL,
##  FDA_FOSSONE_ID VARCHAR (15) DEFAULT NULL,
##  UNIQUE KEY LABID (LABID)
##) ENGINE=InnoDB DEFAULT CHARSET=latin1 ROW_FORMAT=COMPRESSED"
##
##$z=0
##$proddate=[datetime]$($lr[$z].ProductionStartDate)
##$prodval = "$($proddate.Year)-$($proddate.Month)-$($proddate.Day)"
##$lastdate=[datetime]$($lr[$z].LastUpdateTime)
##$lastval = "$($lastdate.Year)-$($lastdate.Month)-$($lastdate.Day) $($lastdate.Hour):$($lastdate.Minute):$($lastdate.Second)"
##$FDA_FOSSONE_ID="$($proddate.Month.ToString("00"))$($proddate.Day.ToString("00"))$($proddate.Year)$($lr[$z].MfgSequenceNo)"
##$QS="INSERT INTO dsilabresults (FDA_FOSSONE_ID, CompanyNumber, VAT, ProductionStartDate, TestScoreSgt, TestResult, LabTestSgt, LastUpdateTime, ColumnHeading) VALUES (`'$FDA_FOSSONE_ID`', `'$($lr[$z].CompanyNumber)`', `'$($lr[$z].MfgSequenceNo)`', `'$prodval', `'$($lr[$z].TestScoreSgt)`', `'$($lr[$z].TestResult)`', `'$($lr[$z].LabTestSgt)`', `'$lastval`', `'$($lr[$z].ColumnHeading)`')"
##$stringbuilder = New-Object System.Text.StringBuilder
##$null = $stringbuilder.Append($QS)
##
##$lrct=(($lr.Count)-1)
##for ($z=1; $z -le $lrct; $z++){
##$proddate=[datetime]$($lr[$z].ProductionStartDate)
##$prodval = "$($proddate.Year)-$($proddate.Month)-$($proddate.Day)"
##$lastdate=[datetime]$($lr[$z].LastUpdateTime)
##$lastval = "$($lastdate.Year)-$($lastdate.Month)-$($lastdate.Day) $($lastdate.Hour):$($lastdate.Minute):$($lastdate.Second)"
##$FDA_FOSSONE_ID="$($proddate.Month.ToString("00"))$($proddate.Day.ToString("00"))$($proddate.Year)$($lr[$z].MfgSequenceNo)"
##$null = $stringbuilder.Append(", (`'$FDA_FOSSONE_ID`', `'$($lr[$z].CompanyNumber)`', `'$($lr[$z].MfgSequenceNo)`', `'$prodval', `'$($lr[$z].TestScoreSgt)`', `'$($lr[$z].TestResult)`', `'$($lr[$z].LabTestSgt)`', `'$lastval`', `'$($lr[$z].ColumnHeading)`')")
##}
##$outs= $stringbuilder.ToString()
##Invoke-Mysql -Query "$outs"
##
###Adding data to dsilabreultshist
##Invoke-Mysql -Query "DELETE FROM `dsilabresults` WHERE `VAT` IS NULL OR VAT > 100"
##Invoke-Mysql -Query "REPLACE INTO dsilabresultshist (FDA_FOSSONE_ID) SELECT FDA_FOSSONE_ID FROM dsilabresults"
##Invoke-Mysql -Query "update dsilabresults,dsilabresultshist set dsilabresultshist.VAT = dsilabresults.VAT where dsilabresultshist.VAT IS NULL and dsilabresults.FDA_FOSSONE_ID=dsilabresultshist.FDA_FOSSONE_ID"
##Invoke-Mysql -Query "update dsilabresults,dsilabresultshist set dsilabresultshist.ProductionStartDate = dsilabresults.ProductionStartDate where dsilabresultshist.ProductionStartDate IS NULL and dsilabresults.FDA_FOSSONE_ID=dsilabresultshist.FDA_FOSSONE_ID"
##Invoke-Mysql -Query "update dsilabresults,dsilabresultshist set dsilabresultshist.TestScoreSgt = dsilabresults.TestScoreSgt where dsilabresultshist.TestScoreSgt IS NULL and dsilabresults.FDA_FOSSONE_ID=dsilabresultshist.FDA_FOSSONE_ID"
##Invoke-Mysql -Query "update dsilabresults,dsilabresultshist set dsilabresultshist.PROTEIN = dsilabresults.TestResult where dsilabresults.FDA_FOSSONE_ID=dsilabresultshist.FDA_FOSSONE_ID AND dsilabresults.ColumnHeading = 'PROTEIN'"
##Invoke-Mysql -Query "update dsilabresults,dsilabresultshist set dsilabresultshist.SALT = dsilabresults.TestResult where dsilabresults.FDA_FOSSONE_ID=dsilabresultshist.FDA_FOSSONE_ID AND dsilabresults.ColumnHeading = 'SALT'"
##Invoke-Mysql -Query "update dsilabresults,dsilabresultshist set dsilabresultshist.FAT = dsilabresults.TestResult where dsilabresults.FDA_FOSSONE_ID=dsilabresultshist.FDA_FOSSONE_ID AND dsilabresults.ColumnHeading = 'FAT'"
##Invoke-Mysql -Query "update dsilabresults,dsilabresultshist set dsilabresultshist.FDB = dsilabresults.TestResult where dsilabresults.FDA_FOSSONE_ID=dsilabresultshist.FDA_FOSSONE_ID AND dsilabresults.ColumnHeading = 'FDB'"
##Invoke-Mysql -Query "update dsilabresults,dsilabresultshist set dsilabresultshist.MOIST = dsilabresults.TestResult where dsilabresults.FDA_FOSSONE_ID=dsilabresultshist.FDA_FOSSONE_ID AND dsilabresults.ColumnHeading = 'MOISTURE'"
##Invoke-Mysql -Query "update dsilabresults,dsilabresultshist set dsilabresultshist.pH = dsilabresults.TestResult where dsilabresults.FDA_FOSSONE_ID=dsilabresultshist.FDA_FOSSONE_ID AND dsilabresults.ColumnHeading = 'pH'"
##Invoke-Mysql -Query "update dsilabresults,dsilabresultshist set dsilabresultshist.LastUpdateTime = dsilabresults.LastUpdateTime where dsilabresults.FDA_FOSSONE_ID=dsilabresultshist.FDA_FOSSONE_ID"
##
##
##}
##Catch {
##  Logwrite "ERROR : UNABLE TO RUN SECTION DSI LAB RESULTS: $query `n$Error[0]"
##  Write-host "ERROR : UNABLE TO RUN SECTION DSI LAB RESULTS: $query `n$Error[0]"
##  $Error.Clear()
## }
##Finally {
##  if($debuglvl -gt 0){Logwrite "DSI LAB RESULTS update ran.."  }
##  }

Write-host "Total time for dsilabresults transform and feed is $($sw.Totalseconds)"
if($debuglvl -gt 0){LogWrite "Total time for dsilabresults transform and feed is $($sw.Totalseconds)"}

}
if($debuglvl -gt 0){Logwrite "2nd Scriptblock ran in $sblock"}
}

#########################################################
#########  END OF SCRIPTBLOCK TWO########################
#########################################################


#########################################################
#########  START OF SCRIPTBLOCK THREE####################
#########################################################

$scriptblockThree = {
$sblock=Measure-Command {

############  SALES SECTION  #################

#Plling and storing open orders, pulling DSI shipping data and updating OE with it, pulling DSI refnum and updating.

$sw=Measure-Command {
try{
$oe = Invoke-MSSQL -Query "select * from dbo.stuff"
$rn = Invoke-MSSQL -Query "SELECT DISTINCT OrderDate, OrderNumber, PurchaseOrder, OrderRefCode, ShipVia FROM dbo.stuff WHERE OrderDate > DATEADD(day, -150, GETDATE())"
$si = Invoke-MSSQL -Query "select Ordernumber,Itemcode,PurchaseOrder,ShipDate, BOLNumber from [DSIPAR].[dbo].[v_nof_getBillOfLading] WHERE ShipDate > DATEADD(day, -90, GETDATE())"
##massive string builder for insert to add new rows here
$z=0
$orderdate=[datetime]$($oe[$z].ORDERDATE)
$odval = "$($orderdate.Year)-$($orderdate.Month)-$($orderdate.Day)"
$rsdate=[datetime]$($oe[$z].RequestedShipDate)
$rsdval = "$($rsdate.Year)-$($rsdate.Month)-$($rsdate.Day)" 
$oqty=$($oe[$z].OrderedQuantity)
$oqty=([math]::Round($oqty))
$QtyPicked=$($oe[$z].QtyPicked)
$QtyPicked=([math]::Round($QtyPicked))
$QtyShipped=$($oe[$z].QtyShipped)
$QtyShipped=([math]::Round($QtyShipped))
$QtyBackordered=$($oe[$z].QtyBackordered)
$QtyBackordered=([math]::Round($QtyBackordered))
$QtyCancelled=$($oe[$z].QtyCancelled)
$QtyCancelled=([math]::Round($QtyCancelled))
$QtyInvoiced=$($oe[$z].QtyInvoiced)
$QtyInvoiced=([math]::Round($QtyInvoiced))
$oeid =  "$($oe[$z].OrderNumber)$($oe[$z].ItemCode)$oqty"
$shpto = $($oe[$z].ShipTo)
$shpto = $shpto -replace "`'", "\`'"
$QS="INSERT IGNORE INTO oehist (oeid, CompanyNumber, CompanyNameAndNumber, DivisionNumber, WarehouseSgt, Warehouse, OrderNum, PurchaseOrder, OrderDate, ReqSDate, CustomerClass, BillToSgt, CustomerID, CustomerName, ShipToSgt, ShipTo, ProductGroup, ProductType, ICode, ItemDescription, UOM, OrdedQty, QtyPicked, QtyShipped, QtyBackordered, QtyCancelled, QtyInvoiced, OrderedQtyInventory, InventoryUOM, OrdStat, ActiveStatus) VALUES (`'$oeid`', `'$($oe[0].CompanyNumber)`',`'$($oe[0].CompanyNameAndNumber)`',`'$($oe[0].DivisionNumber)`',`'$($oe[0].WarehouseSgt)`',`'$($oe[0].Warehouse)`',`'$($oe[0].OrderNumber)`',`'$($oe[0].PurchaseOrder)`',`'$odval`',`'$rsdval`',`'$($oe[0].CustomerClass)`',`'$($oe[0].BillToSgt)`',`'$($oe[0].CustomerID)`',`'$($oe[0].CustomerName)`',`'$($oe[0].ShipToSgt)`',`'$shpto',`'$($oe[0].ProductGroup)`',`'$($oe[0].ProductType)`',`'$($oe[0].ItemCode)`',`'$($oe[0].ItemDescription)`',`'$($oe[0].UOM)`',`'$oqty`',`'$QtyPicked`',`'$QtyShipped`',`'$QtyBackordered`',`'$QtyCancelled`',`'$QtyInvoiced`',`'$($oe[0].OrderedQtyInventory)`',`'$($oe[0].InventoryUOM)`',`'$($oe[0].OrderStatus)`',`'$($oe[0].ActiveStatus)`')"
$stringbuilder = New-Object System.Text.StringBuilder
$null = $stringbuilder.Append($QS)


$oect=(($oe.Count)-1)
for ($z=1; $z -le $oect; $z++){
$orderdate=[datetime]$($oe[$z].ORDERDATE)
$odval = "$($orderdate.Year)-$($orderdate.Month)-$($orderdate.Day)"
$rsdate=[datetime]$($oe[$z].RequestedShipDate)
$rsdval = "$($rsdate.Year)-$($rsdate.Month)-$($rsdate.Day)" 
$oqty=$($oe[$z].OrderedQuantity)
$oqty=([math]::Round($oqty))
$QtyPicked=$($oe[$z].QtyPicked)
$QtyPicked=([math]::Round($QtyPicked))
$QtyShipped=$($oe[$z].QtyShipped)
$QtyShipped=([math]::Round($QtyShipped))
$QtyBackordered=$($oe[$z].QtyBackordered)
$QtyBackordered=([math]::Round($QtyBackordered))
$QtyCancelled=$($oe[$z].QtyCancelled)
$QtyCancelled=([math]::Round($QtyCancelled))
$QtyInvoiced=$($oe[$z].QtyInvoiced)
$QtyInvoiced=([math]::Round($QtyInvoiced))
$oeid =  "$($oe[$z].OrderNumber)$($oe[$z].ItemCode)$oqty"
$shpto = $($oe[$z].ShipTo)
$shpto = $shpto -replace "`'", "\`'"
$null = $stringbuilder.Append(", (`'$oeid`', `'$($oe[$z].CompanyNumber)`',`'$($oe[$z].CompanyNameAndNumber)`',`'$($oe[$z].DivisionNumber)`',`'$($oe[$z].WarehouseSgt)`',`'$($oe[$z].Warehouse)`',`'$($oe[$z].OrderNumber)`',`'$($oe[$z].PurchaseOrder)`',`'$odval`',`'$rsdval`',`'$($oe[$z].CustomerClass)`',`'$($oe[$z].BillToSgt)`',`'$($oe[$z].CustomerID)`',`'$($oe[$z].CustomerName)`',`'$($oe[$z].ShipToSgt)`',`'$shpto',`'$($oe[$z].ProductGroup)`',`'$($oe[$z].ProductType)`',`'$($oe[$z].ItemCode)`',`'$($oe[$z].ItemDescription)`',`'$($oe[$z].UOM)`',`'$oqty`',`'$QtyPicked`',`'$QtyShipped`',`'$QtyBackordered`',`'$QtyCancelled`',`'$QtyInvoiced`',`'$($oe[$z].OrderedQtyInventory)`',`'$($oe[$z].InventoryUOM)`',`'$($oe[$z].OrderStatus)`',`'$($oe[$z].ActiveStatus)`')")
#Invoke-Mysql -Query "INSERT IGNORE INTO oehist (oeid, CompanyNumber, CompanyNameAndNumber, DivisionNumber, WarehouseSgt, Warehouse, OrderNum, PurchaseOrder, OrderDate, ReqSDate, CustomerClass, BillToSgt, CustomerID, CustomerName, ShipToSgt, ShipTo, ProductGroup, ProductType, ICode, ItemDescription, UOM, OrdedQty, QtyPicked, QtyShipped, QtyBackordered, QtyCancelled, QtyInvoiced, OrderedQtyInventory, InventoryUOM, OrdStat, ActiveStatus) VALUES (`'$oeid`', `'$($oe[$z].CompanyNumber)`',`'$($oe[$z].CompanyNameAndNumber)`',`'$($oe[$z].DivisionNumber)`',`'$($oe[$z].WarehouseSgt)`',`'$($oe[$z].Warehouse)`',`'$($oe[$z].OrderNumber)`',`'$($oe[$z].PurchaseOrder)`',`'$odval`',`'$rsdval`',`'$($oe[$z].CustomerClass)`',`'$($oe[$z].BillToSgt)`',`'$($oe[$z].CustomerID)`',`'$($oe[$z].CustomerName)`',`'$($oe[$z].ShipToSgt)`',`'$shpto',`'$($oe[$z].ProductGroup)`',`'$($oe[$z].ProductType)`',`'$($oe[$z].ItemCode)`',`'$($oe[$z].ItemDescription)`',`'$($oe[$z].UOM)`',`'$oqty`',`'$QtyPicked`',`'$QtyShipped`',`'$QtyBackordered`',`'$QtyCancelled`',`'$QtyInvoiced`',`'$($oe[$z].OrderedQtyInventory)`',`'$($oe[$z].InventoryUOM)`',`'$($oe[$z].OrderStatus)`',`'$($oe[$z].ActiveStatus)`')"
}
$outs= $stringbuilder.ToString()
Invoke-Mysql -Query "$outs"
## done with massive string builder one
## update string needs optimization yet but would require alot more time and reconstructing.
$oect=(($oe.Count)-1)
for ($z=0; $z -le $oect; $z++){
$orderdate=[datetime]$($oe[$z].ORDERDATE)
$odval = "$($orderdate.Year)-$($orderdate.Month)-$($orderdate.Day)"
$rsdate=[datetime]$($oe[$z].RequestedShipDate)
$rsdval = "$($rsdate.Year)-$($rsdate.Month)-$($rsdate.Day)" 
$oqty=$($oe[$z].OrderedQuantity)
$oqty=([math]::Round($oqty))
$QtyPicked=$($oe[$z].QtyPicked)
$QtyPicked=([math]::Round($QtyPicked))
$QtyShipped=$($oe[$z].QtyShipped)
$QtyShipped=([math]::Round($QtyShipped))
$QtyBackordered=$($oe[$z].QtyBackordered)
$QtyBackordered=([math]::Round($QtyBackordered))
$QtyCancelled=$($oe[$z].QtyCancelled)
$QtyCancelled=([math]::Round($QtyCancelled))
$QtyInvoiced=$($oe[$z].QtyInvoiced)
$QtyInvoiced=([math]::Round($QtyInvoiced))
$oeid =  "$($oe[$z].OrderNumber)$($oe[$z].ItemCode)$oqty"
$shpto = $($oe[$z].ShipTo)
$shpto = $shpto -replace "`'", "\`'"
Invoke-Mysql -Query "UPDATE oehist SET oehist.CompanyNumber = `'$($oe[$z].CompanyNumber)`', oehist.CompanyNameAndNumber = `'$($oe[$z].CompanyNameAndNumber)`', oehist.DivisionNumber = `'$($oe[$z].DivisionNumber)`', oehist.WarehouseSgt = `'$($oe[$z].WarehouseSgt)`',  oehist.Warehouse = `'$($oe[$z].Warehouse)`',  oehist.OrderNum =`'$($oe[$z].OrderNumber)`',  oehist.PurchaseOrder = `'$($oe[$z].PurchaseOrder)`', oehist.OrderDate = `'$odval`', oehist.ReqSdate = `'$rsdval`',  oehist.CustomerClass = `'$($oe[$z].CustomerClass)`',  oehist.BillToSgt = `'$($oe[$z].BillToSgt)`',  oehist.CustomerID =`'$($oe[$z].CustomerID)`',  oehist.CustomerName = `'$($oe[$z].CustomerName)`',  oehist.ShipToSgt = `'$($oe[$z].ShipToSgt)`', oehist.ShipTo=`'$shpto`',  oehist.ProductGroup = `'$($oe[$z].ProductGroup)`',  oehist.ProductType = `'$($oe[$z].ProductType)`',  oehist.Icode =`'$($oe[$z].ItemCode)`',  oehist.ItemDescription = `'$($oe[$z].ItemDescription)`',  oehist.UOM = `'$($oe[$z].UOM)`',  oehist.OrdedQTY = `'$oqty`',  oehist.QtyPicked =`'$QtyPicked`',  oehist.QtyShipped = `'$QtyShipped`',  oehist.QtyBackordered =`'$QtyBackordered`',  oehist.QtyCancelled = `'$QtyCancelled`',  oehist.QtyInvoiced = `'$QtyInvoiced`',  oehist.OrderedQtyInventory = `'$($oe[$z].OrderedQtyInventory)`',  oehist.InventoryUOM = `'$($oe[$z].InventoryUOM)`',  oehist.OrdStat = `'$($oe[$z].OrderStatus)`',  oehist.ActiveStatus = `'$($oe[$z].ActiveStatus)`' WHERE oehist.oeid LIKE $oeid"
}
##After the main loop apply efects for whole table
Invoke-Mysql -Query "update oehist, custdata
set oehist.email = CONCAT('mailto:', + custdata.email, '?CC=LHavemeier@firstdistrict.com', '&SUBJECT=Purchase Order ', + oehist.PurchaseOrder, ' has been released', '&body=Hello,%0D%0A', 'Purchase order number ', + oehist.PurchaseOrder, ' has been released for ', + oehist.CustomerName, ' and is available to be delivered from First District Association.%0D%0A%0D%0A', 'Customer requested ship date is ', + DATE_FORMAT(oehist.ReqSDate, '%W-%M-%d-%Y'), '. %0D%0A%0D%0A'  'Kind Regards,%0D%0A', 'First District Sales Staff')
WHERE oehist.CustomerID = custdata.CustomerID"

Invoke-Mysql -Query "update custdata
SET custdata.email = NULL
WHERE custdata.email = ''"

Invoke-Mysql -Query "update oehist, custdata
set oehist.email = NULL
WHERE oehist.CustomerID = custdata.CustomerID AND custdata.email IS NULL"

Invoke-Mysql -Query "update oehist, custdata
set oehist.customeragereq = custdata.CustomerAgeReq
WHERE oehist.customerID = custdata.CustomerID"

Invoke-Mysql -Query "update oehist
set oehist.EstRelDate = DATE_ADD(oehist.MfgSchDate, INTERVAL oehist.CustomerAgeReq DAY)
WHERE oehist.MfgSchDate IS NOT NULL"

Invoke-Mysql -Query "update oehist, custdata
set oehist.info = custdata.CustomerDetails
WHERE oehist.customerID = custdata.CustomerID"


#Create drop/create refnum/shipvia table and work it against oehist
Invoke-Mysql -Query "drop table if exists refnum"

Invoke-Mysql -Query "CREATE TABLE IF NOT EXISTS refnum (
  id INT UNSIGNED PRIMARY KEY AUTO_INCREMENT,
  OrderDate datetime DEFAULT NULL,
  OrderNumber int(10) DEFAULT NULL,
  PurchaseOrder text DEFAULT NULL,
  RefCode varchar(25) DEFAULT NULL,
  ShipVia text DEFAULT NULL,
  BOLNumber int(10) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1 ROW_FORMAT=COMPRESSED"


$rnct=(($rn.Count)-1)
for ($z=0; $z -le $rnct; $z++){
$orddate=[datetime]$($rn[$z].OrderDate)
$ordval = "$($orddate.Year)-$($orddate.Month)-$($orddate.Day)"
$shpvia = $($rn[$z].ShipVia)
$shpvia = $shpvia -replace "`'", "\`'"
Invoke-Mysql -Query "INSERT INTO refnum (OrderDate, OrderNumber, PurchaseOrder, RefCode, ShipVia) VALUES (`'$ordval`', `'$($rn[$z].OrderNumber)`', `'$($rn[$z].PurchaseOrder)`', `'$($rn[$z].OrderRefCode)`', `'$shpvia`')"
}
Invoke-Mysql -Query "DELETE FROM refnum WHERE OrderNumber = '' OR OrderNumber IS NULL"


Invoke-Mysql -Query "update oehist, refnum
set oehist.productnum = refnum.RefCode
WHERE oehist.ordernum = refnum.OrderNumber AND oehist.PurchaseOrder = refnum.PurchaseOrder"

Invoke-Mysql -Query "update oehist, refnum
set oehist.Xtra1 = refnum.ShipVia
WHERE oehist.ordernum = refnum.OrderNumber AND oehist.PurchaseOrder = refnum.PurchaseOrder"
#done with refnum

$sict=(($si.Count)-1)
for ($z=0; $z -le $sict; $z++){
$shipdate=[datetime]$($si[$z].ShipDate)
$shipval = "$($shipdate.Year)-$($shipdate.Month)-$($shipdate.Day)"
Invoke-Mysql -Query "UPDATE oehist SET oehist.ActShpDate = `'$shipval`' WHERE oehist.PurchaseOrder LIKE `'%$($si[$z].PurchaseOrder)%`' AND oehist.ordernum = `'$($si[$z].ordernumber)`' AND oehist.Icode LIKE `'%$($si[$z].ItemCode)%`'"
Invoke-Mysql -Query "UPDATE refnum SET refnum.BOLNumber = `'$($si[$z].BOLNumber)`' WHERE refnum.ordernumber = `'$($si[$z].ordernumber)`'"
}

Invoke-Mysql -Query "update oehist
set oehist.loadnum = '0'
WHERE oehist.loadnum IS NULL"

Invoke-Mysql -Query "update shipping, oehist
set shipping.Location = oehist.Warehouse
WHERE shipping.OrdNum = oehist.OrderNum"



}

Catch {
  Logwrite "ERROR : UNABLE TO RUN SECTION OE: $query `n$Error[0]"
  Write-host "ERROR : UNABLE TO RUN SECTION OE: $query `n$Error[0]"
  $Error.Clear()
 }
Finally {
  if($debuglvl -gt 0){Logwrite "OE update ran.."  }
  }
}
Write-host "Total time for oe transform and feed is $($sw.Totalseconds)"
if($debuglvl -gt 0){LogWrite "Total time for oe transform and feed is $($sw.Totalseconds)"}

##$lr = Invoke-MSSQL -Query "SELECT DISTINCT pim.CompanyNumber, pim.MfgSequenceNo, pim.ProductionStartDate, rslt.TestScoreSgt, rslt.TestResult, rslt.LabTestSgt, rslt.LastUpdateTime, lab.ColumnHeading FROM dbo.NBPProductionWithItemsMade AS pim LEFT OUTER JOIN dbo.NIFTestScoreResult AS rslt ON pim.CompanyNumber = rslt.CompanyNumber AND pim.TestScoreSgt = rslt.TestScoreSgt LEFT OUTER JOIN dbo.NIFLabTest AS lab ON rslt.LabTestSgt = lab.LabTestSgt WHERE (ProductionStartDate > DATEADD(day, -14, GETDATE()) AND rslt.LabTestSgt = 1) OR (ProductionStartDate > DATEADD(day, -14, GETDATE()) AND rslt.LabTestSgt = 3) OR (ProductionStartDate > DATEADD(day, -14, GETDATE()) AND rslt.LabTestSgt = 29) OR (ProductionStartDate > DATEADD(day, -14, GETDATE()) AND rslt.LabTestSgt = 4) OR (ProductionStartDate > DATEADD(day, -14, GETDATE()) AND rslt.LabTestSgt = 5) OR (ProductionStartDate > DATEADD(day, -14, GETDATE()) AND rslt.LabTestSgt = 10) ORDER BY ProductionStartDate ASC, MfgSequenceNo ASC, ColumnHeading ASC, rslt.LastUpdateTime DESC"

###########  IT SECTION  #################

#PC info is pushed to csv on server from local machines to capture logins and machine info

$sw=Measure-Command {
try{

$Header = "Timestamp", "UserID", "CompName", "IPv4", "WindowsVer", "OSArch", "SerialNum", "Model", "Processor", "Memory", "MacAddress"
Get-ChildItem "D:\Program Files\DATA\IMPORT\PUSHDATA" -Filter *.csv | 
Foreach-Object {
$csv = import-csv "D:\Program Files\DATA\IMPORT\PUSHDATA\$_" -Header $header
$csvct=$($csv.Count)
if ($csvct -lt 1){$csvct = 1}
for ($z=0; $z -lt $csvct; $z++){
$lastdate=[datetime]$($csv[$z].Timestamp)
$lastval = "$($lastdate.Year)-$($lastdate.Month)-$($lastdate.Day) $($lastdate.Hour):$($lastdate.Minute):$($lastdate.Second)"
$name = $($csv[$z].CompName)
$trimname = $name.Trim()
Invoke-Mysql -Query "REPLACE INTO pcinfo (Timestamp, UserID, CompName, IPv4, WindowsVer, OSArch, SerialNum, Model, Processor, Memory, MacAddress) VALUES (`'$lastval`', `'$($csv[$z].UserID)`', `'$trimname`', `'$($csv[$z].IPv4)`', `'$($csv[$z].WindowsVer)`', `'$($csv[$z].OSArch)`', `'$($csv[$z].SerialNum)`', `'$($csv[$z].Model)`', `'$($csv[$z].Processor)`', `'$($csv[$z].Memory)`', `'$($csv[$z].MacAddress)`')"

}
}
remove-item -path "D:\Program Files\DATA\IMPORT\PUSHDATA\*.*" -force
}

Catch {
  Logwrite "ERROR : UNABLE TO RUN SECTION IT : $query `n$Error[0]"
  Write-host "ERROR : UNABLE TO RUN SECTION IT: $query `n$Error[0]"
  $Error.Clear()
 }
Finally {
  if($debuglvl -gt 0){Logwrite "IT update ran.."  }
  }
}
Write-host "Total time for IT transform and feed is $($sw.Totalseconds)"
if($debuglvl -gt 0){LogWrite "Total time for IT transform and feed is $($sw.Totalseconds)"}


}
if($debuglvl -gt 0){Logwrite "3rd Scriptblock ran in $sblock"}
}

#########################################################
#########  END OF SCRIPTBLOCK THREE######################
#########################################################


Start-Job -Name One -InitializationScript $functions -ScriptBlock $ScriptblockOne
wait-job -name One -timeout 100
Start-Job -Name Two -InitializationScript $functions -ScriptBlock $ScriptblockTwo
wait-job -name Two -timeout 100
Start-Job -Name Three -InitializationScript $functions -ScriptBlock $ScriptblockThree
wait-job -name Three -timeout 100