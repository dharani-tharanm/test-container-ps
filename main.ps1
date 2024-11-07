#simple Advertisement banner
function WriteToLog {
    param (
        [string]$message,
        [string]$logFilePath
    )
    $timestamp = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
    $logMessage = "$timestamp $message"
    Add-Content -Path $logFilePath -Value $logMessage
    
}
function StoreBanner {
    param (
        [string]$title,
        [string]$slogan
    )
    $title = "Welcome To the Store !!!"
    $slogan = "Buy for your Loved Ones"
    $format_title = "    $title    "
    $format_slogan = "    $slogan     "
    $height = 3
    $width = $format_title.Length + 4
    for($i=0;$i -lt $height;$i++){
        if($i -eq 0 -or $i -eq ($height-1)){
            Write-Host ("*" * $width)
        }
        else{
            Write-Host("* "+"$format_title"+" *")
            Write-Host("*"*$width)
            Write-Host("* "+"     New Collection Arrived     "+" *")
            Write-Host("* "+"$format_slogan"+" *")
        }
    }
    
}
#Get Username and Password
function Login {
    param (
        [string]$username,
        [string]$paswordd
    )
    Write-Host "Entered Username $username and Password $paswordd waiting for Authentication"
    WriteToLog "Getting user and pass from user input" $logFilePath
}
#Compare the Username and Password with the Existing data_worksheet ->data_work ->excel
function Check_Credentials {
    param (
        [string]$username,
        [string]$paswordd,
        [Object]$data_worksheet
    )

    $check_user=$data_worksheet.Rows.Item(1).Find('Username').Column
    $check_pass=$data_worksheet.Rows.Item(1).Find('Password').Column
    $rowrange=$data_worksheet.UsedRange.Rows.Count
    
    for($row=2;$row -le $rowrange;$row++){
        $cell_user=$data_worksheet.Cells.Item($row,$check_user)
        $cell_pass=$data_worksheet.Cells.Item($row,$check_pass)
        if($username -eq $cell_user.Text -and $paswordd -eq $cell_pass.Text){
            #Write-Host "Welcome Back $username "
            $searchRange = $data_worksheet.UsedRange.Find($username).Row
            $searchCol = $data_worksheet.Rows.item(1).Find('Name').Column
            $cellValue = $data_worksheet.Cells.Item($searchRange, $searchCol).Text
            Write-Host "Welcome Back $cellValue "
            break
        }
        else {
            Write-Host "Enter Valid Credentials"
        }
    }

}
#Another Advertisement for Selection Item
function WelcomeuserBanner {
    $ad="*"+" MEN "+"*"+" WEMEN "+"*"+" KIDS "+"*"
    $format_ad = "* $ad *"
    $height=4
    $width=$ad.Length +4
    for($i=0;$i -lt $height;$i++){
        if($i -eq 0 -or $i -eq ($height-1)){
            Write-Host ("*"*$width)
        }
        else{
            Write-Host ($format_ad)
        }
    }
    Write-Host "1. MENSWARE"
    Write-Host "2. WEMENSWARE"
    Write-Host "3. KIDSWARE"
}
#Choose the Item Type
function CatagorySelection {
    param (
        [string]$select,
        [string]$username
    )
    if($select -eq 'MENSWARE'){
        Write-Host "MENSWARE IS A GREAT CHOICE"
    }
    elseif($select -eq 'WEMENSWARE'){
        Write-Host "WEMENSWARE IS A GREAT CHOICE"
    }
    elseif($select -eq 'KIDSWARE'){
        Write-Host "KIDSWARE IS A GREAT CHOICE"
    }
    else{
        Write-Host "No Input Recived"
    }
    WriteToLog "$username selected catagory is $select" $logFilePath 
    
}
#Get the Gender as Input and find relative item from Store_worksheet ->store_workbook ->excel
function Catagory {
    param (
        [string]$select,
        [Object]$store_worksheet
    )
    $list=@()
    $Find=$store_worksheet.Rows.Item(1).Find('Gender').Column
    $rowrange=$store_worksheet.UsedRange.Rows.Count
    $colrange=$store_worksheet.UsedRange.Columns.Count
    for($row=2;$row -le $rowrange;$row++){
        $cell=$store_worksheet.Cells.Item($row,$Find)
        if($select -eq $cell.Text){
            $store=[ordered]@{}
            for($col=1;$col -le $colrange;$col++){
                $cellName=$store_worksheet.Cells.Item(1,$col).Text
                $cellValue=$store_worksheet.Cells.Item($row,$col).Text
                $store[$cellName]=$cellValue
            }
            $obj=New-Object psobject -Property $store
            $list+=$obj
        }
    }

    if($list.Count -gt 0){
        $list | Format-Table -AutoSize
        WriteToLog "The selected Catagory is Shown" $logFilePath
    }
    else{
        Write-Host "No Data Found 404"
        WriteToLog "The selected catagory is not found" $logFilePath
    }
}

##Get the ID as Input and find relative item from Store_worksheet ->store_workbook ->excel
function Purchase {
    param (
        [string]$product_id,
        [Object]$store_worksheet
    )
    $search = $store_worksheet.Rows.Item(1).Find('purchaseitem_id').Column
    $rowrange = $store_worksheet.UsedRange.Rows.Count
    $colrange = $store_worksheet.UsedRange.Columns.Count
    $list=@()
    for($row = 2; $row -le $rowrange; $row++){
        $cell=$store_worksheet.Cells.Item($row,$search)
        if($product_id -eq $cell.Text){
            $product=[Ordered]@{}
            for($col=1;$col -le $colrange;$col++){
                $CellName=$store_worksheet.Cells.item(1,$col).Text
                $cellValue=$store_worksheet.Cells.item($row,$col).Text
                $product[$cellName]=$cellValue
            }
            $obj = New-Object psobject -Property $product
            $list+=$obj
        }
    }
    if($list.Count -gt 0){
        $list | Format-Table -AutoSize
        WriteToLog "The selected Id is shown " $logFilePath
    }
    else{
        "404 - Error in Product_Id"
        WriteToLog "The selected Id is Not found " $logFilePath
    }
}
#compare the item's price with amount 
function Order {
    param (
        [string]$amount,
        [string]$product_id,
        [Object]$store_worksheet
    )
    $wshell = New-Object -ComObject Wscript.Shell
    $searchRange = $store_worksheet.UsedRange
    $searchResult = $searchRange.Find($product_id)
    $searchCol = $store_worksheet.Rows.item(1).Find('Price').Column
    $rowIndex = $searchResult.Row
    $colIndex = $searchCol
    $cellValue = $store_worksheet.Cells.Item($rowIndex, $colIndex).Text
    if($cellValue -eq $amount){
        Write-Host "*****************************"
        Write-Host "The Payment paid sucessfully"
        Write-Host "*****************************"
        WriteToLog "The Payment paid successfully, amount paid: $amount" $logFilePath
        $Output = $wshell.Popup("The Payment paid sucessfully!", 0, "Ecommerce")

    }
    elseif($cellValue -gt $amount){
        Write-Host "************************************************"
        Write-Host "The Payment paid sucessfully due to Less Amount"
        Write-Host "**************************************************"
        WriteToLog "The Payment Failed due to Insufficient amount, amount paid: $amount" $logFilePath
        $Output = $wshell.Popup("The Payment Failed due to insufficient amount paid!", 0, "Ecommerce")
    }
    else{
        Write-Host "***************"
        Write-Host "Payment failed"
        Write-Host "***************" 
        WriteToLog "The Payment Failed" $logFilePath
        $Output = $wshell.Popup("Payment failed", 0, "Ecommerce")       
    }
}
#Generate bill of the purchased Item
function Bill {
    param (
        [string]$username,
        [string]$amount,
        [string]$product_id,
        [Object]$store_worksheet
    )
    $searchRange = $store_worksheet.UsedRange
    $searchResult = $searchRange.Find($product_id)
    $searchCol = $store_worksheet.Rows.item(1).Find('ProductName').Column
    $searchCol1 = $store_worksheet.Rows.item(1).Find('Color').Column
    $searchCol2 = $store_worksheet.Rows.item(1).Find('Gender').Column
    $searchCol3 = $store_worksheet.Rows.item(1).Find('MadeIn').Column
    $searchCol4 = $store_worksheet.Rows.item(1).Find('Price').Column
    $rowIndex = $searchResult.Row
    #$colIndex = $searchCol
    $cellValue = $store_worksheet.Cells.Item($rowIndex, $searchCol).Text
    $cellValue1 = $store_worksheet.Cells.Item($rowIndex, $searchCol1).Text
    $cellValue2 = $store_worksheet.Cells.Item($rowIndex, $searchCol2).Text
    $cellValue3 = $store_worksheet.Cells.Item($rowIndex, $searchCol3).Text
    $cellValue4 = $store_worksheet.Cells.Item($rowIndex, $searchCol4).Text
    
    $bill = [ordered]@{
        "Username" = $username
        "Product_id" =$product_id
        "ProductName" =$cellValue
        "Color" = $cellValue1
        "Gender" = $cellValue2
        "MadeIn" = $cellValue3
        "Price" = $cellValue4
        "Amount Paid" = $amount
    }
    Write-Host "Bill Summary of product $cellValue"
    foreach ($key in $bill.Keys) {
        Write-Host "   $key : $($bill[$key])   "
    }
    WriteToLog "Bill Generated to the user" $logFilePath
}
#Close the Excel and other garbage values
function CloseExcelFile {
    param (
        [string]$username,
        [Object]$data_workbook,
        [Object]$excel
    )
    WriteToLog "$username placed order successfully" $logFilePath
    $data_workbook.close($true)
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($data_workbook) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
}
function StoreWorkBook {
    param (
        [string]$storePath,
        [Object]$excel
    )
    
    $store_workbook=$excel.workbooks.Open($storePath)
    return $store_workbook   
}
function StoreWorksheet {
    param (
        [Object]$store_workbook
    )
    $store_worksheet=$store_workbook.Sheets.Item(1)
    return $store_worksheet  
}

#EcommerceFunction [Master Branch]
function Ecommerce {
    param (
        [string]$userPath,
        [string]$storePath,
        [string]$logFilePath
    )
    WriteToLog " Project Ecommercce get started" $logFilePath
    $excel= New-Object -ComObject Excel.Application
    $data_workbook = $excel.Workbooks.Open($userpath)
    $data_worksheet = $data_workbook.Sheets.Item(1)
    StoreBanner
    $username = Read-Host "Enter Your Username"
    $paswordd = Read-Host "Enter your Password"
    Login -username $username -paswordd $paswordd
    Check_Credentials -username $username -paswordd $paswordd -data_worksheet $data_worksheet
    WelcomeuserBanner
    #try using next excel data
    $select = Read-Host "Choose Your Catagory"
    CatagorySelection -select $select -username $username
    $store_workbook = StoreWorkBook -storePath $storePath -excel $excel
    $store_worksheet = StoreWorksheet -store_workbook $store_workbook
    Catagory -select $select -store_worksheet $store_worksheet
    $product_id = Read-Host "Enter the id Example #0001"
    Purchase -product_id $product_id -store_worksheet $store_worksheet
    $amount =Read-Host "Enter Amount to pay: "
    Order -amount $amount -product_id $product_id -store_worksheet $store_worksheet
    Bill -username $username -amount $amount -product_id $product_id -store_worksheet $store_worksheet
    CloseExcelFile -username $username -data_workbook $data_workbook -excel $data_workbook.Application 
}
#files and Global Call of Master Function
$userPath = 'C:\Users\z004yw6y\Git\Devops\Week4\Data\User.xlsx'
$storePath = 'C:\Users\z004yw6y\Git\Devops\Week4\Data\Store.xlsx'
$logFilePath = 'C:\Users\z004yw6y\Git\Devops\Week2\Log.txt'
Ecommerce -userPath $userPath -storePath $storePath -logFilePath $logFilePath
WriteToLog "user Logged out Successfully" $logFilePath