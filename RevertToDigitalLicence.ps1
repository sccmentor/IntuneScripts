$CheckForGVLK = Get-WmiObject SoftwareLicensingProduct -Filter "ApplicationID = '55c92734-d682-4d71-983e-d6ec3f16059f' and LicenseStatus = '5'"
$CheckForGVLK = $CheckForGVLK.ProductKeyChannel

if ($CheckForGVLK -eq 'Volume:GVLK'){

    $GetDigitalLicence = (Get-WmiObject -query 'select * from SoftwareLicensingService’).OA3xOriginalProductKey
    cscript c:\windows\system32\slmgr.vbs -ipk $GetDigitalLicence
}
