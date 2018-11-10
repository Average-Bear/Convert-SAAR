<#
.SYNOPSIS
    Parse DD Form 2875 for user account creation data, using iTextSharp library.

.DESCRIPTION
    Parse DD Form 2875 for user account creation data, using iTextSharp library.

.NOTES
    Author: JBear
    Date: 11/10/2018
    Version: 1.0

    Notes: PDF document(s) must not be (SECURED). Requires iTextSharp library.
#>

#Load iTextSharp .NET library
try {
    
    Add-Type -Path 'C:\Windows\System32\itextsharp.dll'
}

catch {

   Write-Error $_
}

#Retrieve files from specified SAAR directory
$Files = (Get-ChildItem 'C:\users\Averagebear\Desktop\SAAR\').FullName

foreach($File in $Files) {

    #Open PDF object
    $PDF = New-Object iTextSharp.text.pdf.PdfReader -ArgumentList $File

    #Retrieve required fields and data
    $Data = $PDF.AcroFields.XFA.DatasetsNode.Data.topmostSubform

    #Split name into first, middle, and last
    $Name = $Data.Name.Split(',').Split(" ") | where {$_ -ne ''}

    #Verify user signature
    $UserSignature = $PDF.AcroFields.VerifySignature("usersign")

    [PSCUstomObject] @{
    
        #User first name
        Firstname = $Name[0].Trim()

        #User last name
        LastName = $Name[1].Trim()

        #User middle initial
        MiddleIn = $Name[2].Trim()

        #User CAC/EDI number
        UserID = 
        
        if(!([String]::IsNullOrWhiteSpace($UserSignature.SignName))) {
            
            #Retrieve 10 consecutive numbers from string (Signature EDI)
            if($UserSignature.SignName -Match "([0-9]{10})"){
                
                $Matches[1].Trim()
            } 
                
            #If SAAR form is not signed with CAC
            else {
                
                $Data.UserID.Replace('EDIPI','').Replace(' ','').Replace('#','').Replace(':','').Trim()
            }
        }

        #If SAAR form is not signed by user
        else {
            
            $Data.UserID.Replace('EDIPI','').Replace(' ','').Replace('#','').Replace(':','').Trim()
        }

        #User email address
        Email = $Data.ReqEmail.Trim()

        #User phone number
        Phone = $Data.ReqPhone.Trim()

        #User organization
        Organization = $Data.ReqOrg.Trim()

        #User job title
        JobTitle = $Data.ReqTitle.Trim()
    }
}
