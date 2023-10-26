# AutomatePackagingandPatching

<h3>Overview</h3>
This is a new iteration of a script that leverages the Evergreen PowerShell module, Windows Package Manager, Chocolatey and the Numecent Cloudpager module to automatically package and patch applications using Numecent Cloudpaging Studio for packaging and Cloudpager for deployment. Combining these technologies can provide a solution to automatically package application and patch plus deploy to your early adopters automatically. Thank to Cloudpager's unique rapid rollback feature, you can also have confidence that any changes made that may need to be rolled back, can be rolled back quickly and easily. As the applications are packaged into application containers before deployment, the rollback and any removal should always be clean and net-new deployments and application updates are delivered dynamically and seamlessly to the users.

I put info at the bottom of this document to get a video of me walking through the script and explaining the main parts of the script.

<h3>Disclaimer</h3>

During my time at Numecent, I developed this project independently. It served as a trial for the Cloudpager API and an improvement to my existing automated application packaging and patching system. Although you are welcome to utilize this script as you wish, please note that it is not an official product of Numecent, and contacting Numecent support regarding this script is unnecessary.

<h3>Setting Up Evergreen Packaging Machine</h3>

I have only tested the script on Windows 10 Enterprise. It should work on a machine that is domain joined or non-domain joined. Typically, the preference for pacakging VMs is that they should be non-domain joined to avoid unncessary noise polluting packages. 

![image](https://github.com/Rorymon/AutomateEvergreenPackaging/assets/7652987/4697b8ec-1fa8-4c41-94b9-54004b3549d8)

On your Windows 10 VM, install Cloudpaging Studio. This will be used for packaging the applications and patches into application containers. You can find the latest version of Studio on the Numecent customer support portal.

In order to download applications found in the Evergreen PowerShell and upload completed packages using this script, you should launch the default browser on your machine at least once before taking your snapshot. If Internet Explorer is still on the VM, you should launch this once too. You should also install WinGet (App Installer App) and Chocolatey on the packaging VM.

![image](https://github.com/Rorymon/AutomateEvergreenPackaging/assets/7652987/7d2d2610-da11-43c8-bc29-cf161776c671)

To use the script, you will need to get the Cloudpager PowerShell module from Numecent and you will need to contact support to retrieve your tenant's unique API key. To ensure you can use the PowerShell module in the script as needed, I recommend placing the module in your Modules directory on your VM e.g. C:\Windows\System32\WindowsPowerShell\v1.0\Modules\Cloudpager.

![image](https://github.com/Rorymon/AutomateEvergreenPackaging/assets/7652987/1c52585e-4af7-4d55-8af5-8ec9e9b89591)

Next step, run the SetupScript.ps1 that is on my GitHub repository [(https://github.com/Rorymon/AutomateEvergreenPackaging/blob/main/SetupScript.ps1)](https://github.com/Rorymon/AutomatePackagingAndPatching/blob/main/SetupScript.ps1), this will create the directories for your packaging projects and it will download the relevant scripts.

![image](https://github.com/Rorymon/AutomateEvergreenPackaging/assets/7652987/678dedf0-cfa5-4c0f-8749-779015af7eb7)

Next, install the Evergreen PowerShell Module e.g. Install-Module Evergreen. Also installer the App Installer app from the Microsoft Store and Chocolatey (PowerShell cmdlet to install Chocolatey can be found on the Chocolatey site: https://chocolatey.org/install.

![image](https://github.com/Rorymon/AutomateEvergreenPackaging/assets/7652987/70fffbef-3251-4408-aee2-ff9e8aea1c7c)

I recommend following Numecent's guidance on prepping your packaging VM. [https://github.com/numecent](https://github.com/Numecent/Automated-Packaging/blob/master/CloudpagingStudio-prep.ps1) using the prep script on the Numecent github repo. 

![image](https://github.com/Rorymon/AutomateEvergreenPackaging/assets/7652987/775950de-9b34-4496-b230-d36e76a968d9)

With the pre-requisites all now on your VM, you should open the main script to populate some of the variables that are unique to you.

![image](https://github.com/Rorymon/AutomateEvergreenPackaging/assets/7652987/e9c89ee9-f60b-458b-b605-a63b13ff8a66)

You will need to input your API Key and you can optionally step through the script, uncomment some of the optional sections and put in other variables such as your OpenAI API key (if you have one) and a Teams Channel Webhook URI.

When you had edited the PowerShell script and input your API key. It is a good idea to reboot the VM at least once and ensure services were stopped correctly, UAC is disabled etc. When ready, shutdown and save a snapshot of your VM.

![image](https://github.com/Rorymon/AutomateEvergreenPackaging/assets/7652987/01863af7-9144-43c9-8209-9c7deadc37d2)

Power back on the VM and run your script for the first time.

![image](https://github.com/Rorymon/AutomateEvergreenPackaging/assets/7652987/38d5e773-40d7-4e84-b1ff-16f2136b2da0)

<h3>Using the Script</h3>

Here is a common example you may use:

```
.\AutomateEvergreenPackaging.ps1 -AppName "GoogleChrome" -Publisher "Google" -Sourcepackagetype "msi" -Sourcechannel "stable" -image_file_path "<PathToImageFile" -CommandLine "C:\Program Files\Google\Chrome\Application\chrome.exe" -Description "Google Chrome is the world's most popular web browser." -WorkpodID "<WorkdpodID>" -Chocolatey $false -Evergreen $true -WinGet $false
```

Before running this, you will want to get a decent icon image for the application. I tend to use .png files that are around 512 x 512 in size. The WorkdpodID is optional, this is only used when you wish to auto-deploy the new package to a Workpod. I always auto-deploy my applications and patches to an Early Adopters Workpod that contains a subset of my users. This way, if the vendor makes a big change to their application and it upsets users, it won't impact the entire organization. If you do not have a Workpod yet, you can create one via the Cloudpager Admin portal or via PowerShell. To retrieve your WorkpodID, use the Cloudpager PowerShell Module e.g. Get-CloudpagerWorkpod -SubscriptionKey "<CloudpagerAPIKey>" -Name "<WorkpodName>"

![image](https://github.com/Rorymon/AutomateEvergreenPackaging/assets/7652987/3dccd9da-67dc-4305-8d33-dffd3b4640ff)

The package uploads may take some time if the application is large. If you did not have Chrome in your tenant already, upon completion of the script you should now have the latest version available. If you already had Chrome but it was not the latest version, the latest version should be published. If the latest version is already in your tenant when you run the script, the script will detect the latest version is already available and exit.

<h3>For More Information</h3>

There are some additional tips you will need when working with the script such as the name currently must match the name listed in the Evergreen PowerShell Module and if the vendor media you select from the Evergreen PowerShell module is an msi, you do not require to pass -Arguments but if it is an exe, you will need to pass arguments to achieve a silent installation.

For a full explanation of this script and how it works, sign up to the Cloudpaging User Group and request access to the Slack Workspace for previous recordings. Join Now: https://www.meetup.com/cloudpaging-user-group/ 
