Function Get-ImageDetails2
{
    begin{        
         [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") |Out-Null 
    } 
     process{
          $fi=[System.IO.FileInfo]$_           
          if( $fi.Exists){
               $img = [System.Drawing.Image]::FromFile($_)
               $img.Clone()
               $img.Dispose()       
          }else{
               Write-Host "File not found: $_" -fore yellow       
          }   
     }    
    end{}
}

Function DirectoryToCreate($DirectoryToCreate) {
    if (-not (Test-Path -LiteralPath $DirectoryToCreate)) {
    
        try {
            New-Item -Path $DirectoryToCreate -ItemType Directory -ErrorAction Stop | Out-Null #-Force
        }
        catch {
            Write-Error -Message "Unable to create directory '$DirectoryToCreate'. Error was: $_" -ErrorAction Stop
        }
        "Successfully created directory '$DirectoryToCreate'."

    }
    else {
        <#"Directory $DirectoryToCreate already existed"#>
    }
}

DirectoryToCreate("$env:USERPROFILE\Pictures\Spotlight\")

Function Get-ImageDetails
{
    begin{        
         [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") |Out-Null 
    } 
     process{
        $file=[System.IO.FileInfo]$_           
        if( $file.Exists){
            $fs = New-Object System.IO.FileStream ($file.FullName, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
            $img=[System.Drawing.Image]::FromStream($fs)
            $fs.Dispose()
            $img | Add-Member `
                        -MemberType NoteProperty `
                        -Name Filename `
                        -Value $file.Fullname `
                        -PassThru
        }
     }    
    end{}
}


dir "$env:LOCALAPPDATA\Packages\Microsoft.Windows.ContentDeliveryManager_cw5n1h2txyewy\LocalState\Assets\" -Recurse | % {

    $image = $_
    $imagePath = $image.PSPath
    $imageName = $image.Name
    $imageDetails = $image| Get-ImageDetails
    
    $imHeight = $imageDetails.Height
    $imWidth = $imageDetails.Width
    if ($imHeight -le $imWidth ) {
        copy-item -path "$imagePath" -destination "$env:USERPROFILE\Pictures\Spotlight\$imageName.jpg"
    }

}