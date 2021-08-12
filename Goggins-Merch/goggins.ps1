
#*******************************************************************************************************************
#This script will iterate over all product items in the hash table every 10 seconds.                               *
#If an item is in stock it will trigger a MS Toast notifcation and send an email to your specified address.        *  
#Since the script already loops, no need to schedule it as a task.                                                 *
#                                                                                                                  *
#To execute the script, follow these steps:                                                                        *
#1. Allow PS scripts to run by updating the SetExecutionPolicy (as Admin) with: Set-ExecutionPolicy Unrestricted   *
#2. Launch Internet Explorer (if you haven't already) and then close it. The IE Engine needs to run at least       *
# once to complete the first-launch configuration. Otherwise the Invoke-WebRequest will fail.                      *
#3. From a cmd promt execute the command: Powershell.exe -windowstyle hidden -file "C:\Users\goggins.ps1"          *
#*******************************************************************************************************************

$smtpCredential = ConvertTo-SecureString –String "<your email password>" –AsPlainText -Force
$smtpUser = "youremail@gmail.com"
$Credential = New-Object –TypeName "System.Management.Automation.PSCredential" –ArgumentList $smtpUser, $smtpCredential

#1. Add/Remove these key/values with links to the products you wish to monitor: 
$product1 = 'https://shop.davidgoggins.com/collections/apparel/products/david_goggins_taking_souls_viscose_tee_black?variant=35981809615005'
$product2 = 'https://shop.davidgoggins.com/collections/second-layer/products/when-the-end-is-unknown-stay-hard-black-raglan-hooded-sweatshirt?variant=37092994646173'
$product3 = 'https://shop.davidgoggins.com/collections/headwear/products/david_goggins_whos_gonna_carry_the_boat_dad_hat' ##This is item in-stock currently. Use this to verify script behaviour for in-stock items
#$product4 = 'https://shop.davidgoggins.com/collections/path/to/other/item/you/want/to/watch/for>'

#2. Correspondingly, add/remove the hash table with the new product# and a name you can identify it by.
$outOfStock = @{
   $product1 =  "Viscose Black Tee"  
   $product2 =  "SH Hooded Sweatshirt" 
   $product3 = "Boat Dad Hat"
   #$product4 = "arbitrary name for item"
    }

do {

        Write-Host ""
        Write-Host "Current items to watch for: "
        Write-host ""
        Write-Host [+]----------------------[+]
        $outOfStock.Values
        Write-Host [+]----------------------[+]
        Write-host ""

    foreach ($product in $($outOfStock.keys)){        
        $WebResponse = Invoke-WebRequest -Uri $product -Method Get
        $button = $WebResponse.ParsedHtml.IHTMLDocument3_getElementsByTagName("div") | Where-Object {$_.className -eq 'productForm-buttons'}
        $status = $button.innerText
            if($status -match 'Add to Cart') {
                
                #[system.console]::beep(300,1000) <-- to make computer go "BEEEEEEEEP" but unnecessary since Toast notification has an audio queue too
                Write-Host "[+++] $($outOfStock[$product]) is IN STOCK!!!"
                
                # Code from https://den.dev/blog/powershell-windows-notification/
                $ToastTitle = 'ALERT: Goggins Merch'
                $ToastText = "[+++] $($outOfStock[$product]) is IN STOCK!!"

                [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] > $null
                $Template = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02)

                $RawXml = [xml] $Template.GetXml()
                ($RawXml.toast.visual.binding.text|where {$_.id -eq "1"}).AppendChild($RawXml.CreateTextNode($ToastTitle)) > $null
                ($RawXml.toast.visual.binding.text|where {$_.id -eq "2"}).AppendChild($RawXml.CreateTextNode($ToastText)) > $null

                $SerializedXml = New-Object Windows.Data.Xml.Dom.XmlDocument
                $SerializedXml.LoadXml($RawXml.OuterXml)

                $Toast = [Windows.UI.Notifications.ToastNotification]::new($SerializedXml)
                $Toast.Tag = "Target"
                $Toast.Group = "Target"
                $Toast.ExpirationTime = [DateTimeOffset]::Now.AddMinutes(1)
                $Toast.Priority = 1

                $Notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("Target")
                $Notifier.Show($Toast);

                $EmailFrom = "youremail@gmail.com"
                $EmailTo = "youremail@gmail.com"
                $Subject = "GOGGINS MERCHANDISE IS AVAILABLE NOW!!!!"
                $Body = "There is a $($outOfStock[$product]) available at online! Get it from here now!: $product"
                
                $SMTPServer = "smtp.gmail.com"
                $SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer, 587)   
                $SMTPClient.EnableSsl = $true
                $SMTPClient.Credentials = New-Object System.Net.NetworkCredential("youremail@gmail.com", $smtpCredential);       
                $SMTPClient.Send($EmailFrom, $EmailTo, $Subject, $Body)
      

                $outOfStock.Remove($product)
            } else {
                Write-Host "[---] $($outOfStock[$product]) is Sold Out :(. Checking next item(s)..."
                Start-Sleep -Seconds 5    
            }
        }
} while($outOfStock.Count -ne 0)