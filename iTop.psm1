<#

Requires MySql libraries to be imported first, see posh-mysql

#>



Function Get-SynchroDataSource
{
<#
    .Synopsis
    Find a synchronization data source

    .Description
    Find a synchronization data source

    .Parameter authName
    Logon for the iTop web service

    .Parameter authPwd
    Password for the iTop web service

    .Parameter uri
    uri for the iTop web service

    .Example
    Get-SynchroDataSource -authName 'user' -authPwd 'password' -uri 'https://webservice.edu'

    .Example
    Get-SynchroDataSource -authName 'user' -authPwd 'password' -uri 'https://webservice.edu' | Where {$_.Name -eq 'DEV VMware Source'}

#>
    Param(
        [Parameter(Mandatory=$True)][string]$authName,
        [Parameter(Mandatory=$True)][string]$authPwd,
        [Parameter(Mandatory=$True)][string]$uri
    )
    Get-iTopObject -objectClass 'SynchroDataSource' -ouputFields '*' -uri $uri -authName $authName -authPwd $authPwd
}

Function Get-Enclosure
{
<#
    .Synopsis
    Get an enclosure or collection of enclosures

    .Description
    Get an enclosure or collection of enclosures

    .Parameter authName
    Logon for the iTop web service

    .Parameter authPwd
    Password for the iTop web service

    .Parameter uri
    uri for the iTop web service

    .Example
    Get-Enclosure -authName 'user' -authPwd 'password' -uri 'https://webservice.edu'

    .Example
    Get-Enclosure -authName 'user' -authPwd 'password' -uri 'https://webservice.edu' | Where {$_.serialnumber -eq 'abc123'}
#>
    Param(
        [Parameter(Mandatory=$True)][string]$authName,
        [Parameter(Mandatory=$True)][string]$authPwd,
        [Parameter(Mandatory=$True)][string]$uri
    )

    Get-iTopObject -objectClass 'Enclosure' -ouputFields '*' -uri $uri -authName $authName -authPwd $authPwd
}

Function Get-StorageSystem
{
<#
    .Synopsis
    Get an enclosure or collection of storage systems

    .Description
    Get an enclosure or collection of storage systems

    .Parameter authName
    Logon for the iTop web service

    .Parameter authPwd
    Password for the iTop web service

    .Parameter uri
    uri for the iTop web service

    .Example
    Get-StorageSystem -authName 'user' -authPwd 'password' -uri 'https://webservice.edu'

    .Example
    Get-StorageSystem -authName 'user' -authPwd 'password' -uri 'https://webservice.edu' | Where {$_.name -eq 'abc123'}
#>
    Param(
        [Parameter(Mandatory=$True)][string]$authName,
        [Parameter(Mandatory=$True)][string]$authPwd,
        [Parameter(Mandatory=$True)][string]$uri
    )

    Get-iTopObject -objectClass 'StorageSystem' -ouputFields '*' -uri $uri -authName $authName -authPwd $authPwd
}

Function Get-LogicalVolume
{
<#
    .Synopsis
    Get an enclosure or collection of logical volumes

    .Description
    Get an enclosure or collection of logical volumes

    .Parameter authName
    Logon for the iTop web service

    .Parameter authPwd
    Password for the iTop web service

    .Parameter uri
    uri for the iTop web service

    .Example
    Get-LogicalVolume -authName 'user' -authPwd 'password' -uri 'https://webservice.edu'

    .Example
    Get-LogicalVolume -authName 'user' -authPwd 'password' -uri 'https://webservice.edu' | Where {$_.name -eq 'abc123'}
#>
    Param(
        [Parameter(Mandatory=$True)][string]$authName,
        [Parameter(Mandatory=$True)][string]$authPwd,
        [Parameter(Mandatory=$True)][string]$uri,
        [Parameter(Mandatory=$False)][string]$oqlFilter,
        [Parameter(Mandatory=$False)][string]$outputFields='*'
    )

    Get-iTopObject -objectClass 'LogicalVolume' -ouputFields $outputFields -oqlFilter $oqlFilter -uri $uri -authName $authName -authPwd $authPwd
}

Function Get-Brand
{
<#
    .Synopsis
    Get a brand typology or collection of brands

    .Description
    Get a brand typology or collection of brands

    .Parameter authName
    Logon for the iTop web service

    .Parameter authPwd
    Password for the iTop web service

    .Parameter uri
    uri for the iTop web service

    .Example
    Get-Brand -authName 'user' -authPwd 'password' -uri 'https://webservice.edu'

    .Example
    Get-Brand -authName 'user' -authPwd 'password' -uri 'https://webservice.edu' | Where {$_.Name -eq 'HP Inc.'}
#>

    Param(
        [Parameter(Mandatory=$True)][string]$authName,
        [Parameter(Mandatory=$True)][string]$authPwd,
        [Parameter(Mandatory=$True)][string]$uri
    )

    Get-iTopObject -objectClass 'Brand' -ouputFields '*' -uri $uri -authName $authName -authPwd $authPwd
}

Function Get-Model
{
<#
    .Synopsis
    Get a model typology or collection of brands

    .Description
    Get a model typology or collection of brands

    .Parameter authName
    Logon for the iTop web service

    .Parameter authPwd
    Password for the iTop web service

    .Parameter uri
    uri for the iTop web service

    .Example
    Get-Model -authName 'user' -authPwd 'password' -uri 'https://webservice.edu'

    .Example
    Get-Model -authName 'user' -authPwd 'password' -uri 'https://webservice.edu' | Where {$_.Name -eq 'R720'}
#>
    Param(
        [Parameter(Mandatory=$True)][string]$authName,
        [Parameter(Mandatory=$True)][string]$authPwd,
        [Parameter(Mandatory=$True)][string]$uri
    )

    Get-iTopObject -objectClass 'Model' -ouputFields '*' -uri $uri -authName $authName -authPwd $authPwd
}


Function Get-OSVersion
{
<#
    .Synopsis
    Get an OS Version typology or collection of OS Versions

    .Description
    Get an OS Version typology or collection of OS Versions

    .Parameter authName
    Logon for the iTop web service

    .Parameter authPwd
    Password for the iTop web service

    .Parameter uri
    uri for the iTop web service

    .Example
    Get-OSVersion -authName 'user' -authPwd 'password' -uri 'https://webservice.edu'

    .Example
    Get-OSVersion -authName 'user' -authPwd 'password' -uri 'https://webservice.edu' | Where {$_.Name -like 'Service Pack*'}
#>
    Param(
        [Parameter(Mandatory=$True)][string]$authName,
        [Parameter(Mandatory=$True)][string]$authPwd,
        [Parameter(Mandatory=$True)][string]$uri
    )

    Get-iTopObject -objectClass 'OSVersion' -ouputFields '*' -uri $uri -authName $authName -authPwd $authPwd
}

Function Get-OSFamily
{
<#
    .Synopsis
    Get an OS Family typology or collection of OS Families

    .Description
    Get an OS Family typology or collection of OS Families

    .Parameter authName
    Logon for the iTop web service

    .Parameter authPwd
    Password for the iTop web service

    .Parameter uri
    uri for the iTop web service

    .Example
    Get-OSFamily -authName 'user' -authPwd 'password' -uri 'https://webservice.edu'

    .Example
    Get-OSFamily -authName 'user' -authPwd 'password' -uri 'https://webservice.edu' | Where {$_.Name -like 'Microsoft*'}
#>
    Param(
        [Parameter(Mandatory=$True)][string]$authName,
        [Parameter(Mandatory=$True)][string]$authPwd,
        [Parameter(Mandatory=$True)][string]$uri
    )

    Get-iTopObject -objectClass 'OSFamily' -ouputFields '*' -uri $uri -authName $authName -authPwd $authPwd
}


Function Get-Server
{
<#
 .Synopsis
  Get a server collection of servers

 .Description
  Get a server collection of servers
  
 .Parameter name
  Optional, otherwise returns collection

 .Parameter authName
  Logon for the iTop web service

 .Parameter authPwd
  Password for the iTop web service

 .Parameter uri
  uri for the iTop web service

 .Example
   Get-Model -authName 'user' -authPwd 'password' -uri 'https://webservice.edu'

 .Example
   Get-Model -authName 'user' -authPwd 'password' -uri 'https://webservice.edu' | Where {$_.serialnumber -eq 'H1SMQ3'}
#>
    Param(
        [Parameter(Mandatory=$True)][string]$authName,
        [Parameter(Mandatory=$True)][string]$authPwd,
        [Parameter(Mandatory=$True)][string]$uri
    )

    Get-iTopObject -objectClass 'Server' -ouputFields '*' -uri $uri -authName $authName -authPwd $authPwd
}

Function Get-VirtualFarm
{
    <#
     .Synopsis
      Find a virtual farm / cluster

     .Description
      Find a virtual farm / cluster

     .Parameter name
      The name of the farm / cluster

     .Parameter authName
      Logon for the iTop web service

     .Parameter authPwd
      Password for the iTop web service

     .Parameter uri
      uri for the iTop web service

     .Example
       Get-VirtualFarm -authName 'user' -authPwd 'password' -uri 'https://webservice.edu' | Where {$_.Name -eq 'Cluster 1'}

    #>
    Param(
        [Parameter(Mandatory=$True)][string]$authName,
        [Parameter(Mandatory=$True)][string]$authPwd,
        [Parameter(Mandatory=$True)][string]$uri,
        [Parameter(Mandatory=$False)][string]$oqlFilter,
        [Parameter(Mandatory=$False)][string]$outputFields='*'
    )

    Get-iTopObject -objectClass 'Farm' -ouputFields $outputFields -oqlFilter $oqlFilter -uri $uri -authName $authName -authPwd $authPwd
}

Function Get-Hypervisor
{
<#
 .Synopsis
  Get a hypervisor or collection of hypervisors

 .Description
  Get a hypervisor or collection of hypervisors

 .Parameter authName
  Logon for the iTop web service

 .Parameter authPwd
  Password for the iTop web service

 .Parameter uri
  uri for the iTop web service

 .Example
    Get-Hypervisor -authName 'user' -authPwd 'password' -uri 'https://webservice.edu' -oqlFilter "WHERE name = `'esxServer01`'" -outputFields 'name,id'
    
 .Example
   Get-Hypervisor -authName 'user' -authPwd 'password' -uri 'https://webservice.edu' | Where {$_.Name -eq 'Server 1'}

#>
    Param(
        [Parameter(Mandatory=$True)][string]$authName,
        [Parameter(Mandatory=$True)][string]$authPwd,
        [Parameter(Mandatory=$True)][string]$uri,
        [Parameter(Mandatory=$False)][string]$oqlFilter,
        [Parameter(Mandatory=$False)][string]$outputFields='*'
    )

    Get-iTopObject -objectClass 'Hypervisor' -ouputFields $outputFields  -uri $uri -authName $authName -authPwd $authPwd -oqlFilter $oqlFilter
}

Function Get-Organization
{
    <#
     .Synopsis
      Get all organizations

     .Description
      Get all organizations

     .Parameter authName
      Logon for the iTop web service

     .Parameter authPwd
      Password for the iTop web service

     .Parameter uri
      uri for the iTop web service

     .Example
       Get-Organization -authName 'user' -authPwd 'password' -uri 'https://webservice.edu'
     
     .Example
       Get-Organization -authName 'user' -authPwd 'password' -uri 'https://webservice.edu' | Where {$_.Name -eq 'UC Berkeley'}

    #>
    Param(
        [Parameter(Mandatory=$True)][string]$authName,
        [Parameter(Mandatory=$True)][string]$authPwd,
        [Parameter(Mandatory=$True)][string]$uri
    )

    Get-iTopObject -objectClass 'Organization' -ouputFields '*' -uri $uri -authName $authName -authPwd $authPwd
}

Function Get-Contact
{
    <#
     .Synopsis
      Get a contact or collection of contacts

     .Description
      Get a contact or collection of contacts including people and teams

     .Parameter authName
      Logon for the iTop web service

     .Parameter authPwd
      Password for the iTop web service

     .Parameter uri
      uri for the iTop web service

     .Example
       Get-Person -authName 'user' -authPwd 'password' -uri 'https://webservice.edu'
     
     .Example
       Get-Person -employeeNumber 123456 -authName 'user' -authPwd 'password' -uri 'https://webservice.edu'

    #>
    Param(
        [Parameter(Mandatory=$True)][string]$authName,
        [Parameter(Mandatory=$True)][string]$authPwd,
        [Parameter(Mandatory=$True)][string]$uri,
        [Parameter(Mandatory=$False)][string]$oqlFilter,
        [Parameter(Mandatory=$False)][string]$outputFields='*'
    )

    Get-iTopObject -objectClass 'Contact' -uri $uri -authName $authName -authPwd $authPwd -oqlFilter $oqlFilter -ouputFields $outputFields
}

Function Get-Person
{
    <#
     .Synopsis
      Get a person or collection of people

     .Description
      Get a person or collection of people

     .Parameter employeeNumber
      Employee number, currently same as UID

     .Parameter authName
      Logon for the iTop web service

     .Parameter authPwd
      Password for the iTop web service

     .Parameter uri
      uri for the iTop web service

     .Example
       Get-Person -authName 'user' -authPwd 'password' -uri 'https://webservice.edu'
     
     .Example
       Get-Person -employeeNumber 123456 -authName 'user' -authPwd 'password' -uri 'https://webservice.edu'

    #>
    Param(
        [Parameter(Mandatory=$True)][string]$authName,
        [Parameter(Mandatory=$True)][string]$authPwd,
        [Parameter(Mandatory=$True)][string]$uri,
        [Parameter(Mandatory=$False)][string]$oqlFilter,
        [Parameter(Mandatory=$False)][string]$outputFields='*'
    )

    Get-iTopObject -objectClass 'Person' -uri $uri -authName $authName -authPwd $authPwd -oqlFilter $oqlFilter -ouputFields $outputFields
}

Function Get-CustomerContract
{
    <#
     .Synopsis
      Get a customer contract or collection of contracts

     .Description
      Get a customer contract or collection of contracts

     .Parameter authName
      Logon for the iTop web service

     .Parameter authPwd
      Password for the iTop web service

     .Parameter uri
      uri for the iTop web service

     .Example
       Get-CustomerContract -authName 'user' -authPwd 'password' -uri 'https://webservice.edu'
     
     .Example
       Get-CustomerContract -authName 'user' -authPwd 'password' -uri 'https://webservice.edu' | Where {$_.cost_unit -eq '1-66083-26328-43'}

    #>
    Param(
        [Parameter(Mandatory=$False)][string]$oqlFilter,        
        [Parameter(Mandatory=$True)][string]$authName,
        [Parameter(Mandatory=$True)][string]$authPwd,
        [Parameter(Mandatory=$True)][string]$uri
    )

    Get-iTopObject -objectClass 'CustomerContract' -ouputFields '*' -uri $uri -authName $authName -authPwd $authPwd -oqlFilter $oqlFilter
}


Function Get-LinkServiceToSLA
{
    Param(
        [Parameter(Mandatory=$True)][object]$service,
        [Parameter(Mandatory=$True)][object]$sla
    )

    $thisArray = @()
    $ciHash = @{}
    $ciHash.add("service_id",("SELECT Service WHERE id =`"$($service.key)`""))
    $ciHash.add("sla_id",("SELECT SLA WHERE id =`"$($sla.key)`""))

    $thisArray += $ciHash
    ,$thisArray
}

Function New-CustomerContract
{
    Param(
        [Parameter(Mandatory=$True)][string]$name,
        [Parameter(Mandatory=$True)][string]$org_name,
        [Parameter(Mandatory=$False)][string]$description,
        [Parameter(Mandatory=$False)][string]$start_date,
        [Parameter(Mandatory=$False)][string]$end_date,
        [Parameter(Mandatory=$False)][string]$cost,
        [Parameter(Mandatory=$False)][string]$cost_currency='dollars',
        [Parameter(Mandatory=$False)][string]$contracttype_name,        # use SELECT ContractType to get relation
        [Parameter(Mandatory=$False)]$contacts,            # collection of people CIs
        [Parameter(Mandatory=$False)][string]$billing_frequency,
        [Parameter(Mandatory=$False)][string]$cost_unit,
        [Parameter(Mandatory=$True)][string]$provider_name,               #use SELECT Organization of of provider
        [Parameter(Mandatory=$True)][string]$status='production',
        [Parameter(Mandatory=$False)]$functionalcis,            # collection of CI objects
        [Parameter(Mandatory=$False)]$linkServiceToSLA,            # array of service to sla hash, get-linkservicetosla
        [Parameter(Mandatory=$True)][string]$authName,
        [Parameter(Mandatory=$True)][string]$authPwd,
        [Parameter(Mandatory=$True)][string]$uri
    )

    # build our linked objects, iTop only want certain lookup fields per object, so we'll feed those in
    $contacts_list = @()
    foreach($ci in $contacts)
    {
        $ciHash = @{}
        $ciHash.add("contact_id",("SELECT Contact WHERE id =`"$($ci.key)`""))
        $contacts_list += $ciHash
    }

    $functionalcis_list = @()
    foreach($ci in $functionalcis)
    {
        $ciHash = @{}
        $ciHash.add("functionalci_id",("SELECT FunctionalCI WHERE id =`"$($ci.key)`""))
        $functionalcis_list += $ciHash
    }


    $fields = New-Object PSObject -Property @{
        name = $name
        org_id = "SELECT Organization WHERE name = `"$org_name`""
        description = $description
        start_date = $start_date
        end_date = $end_date
        cost = $cost
        cost_currency = $cost_currency
        contracttype_id = "SELECT ContractType WHERE name = `"$contracttype_name`""
        contacts_list = $contacts_list
        billing_frequency = $billing_frequency
        cost_unit = $cost_unit
        provider_id = "SELECT Organization WHERE name = `"$provider_name`""
        status = $status
        services_list = $linkServiceToSLA
        functionalcis_list = $functionalcis_list   
    }
    
    $operation = New-Object PSObject -Property @{ 
        operation = 'core/create'
        class = 'CustomerContract'
        comment = 'Synchronization from load scripts'
        output_fields = '*'
        fields = $fields
    }
    GenerateAndSendRequest -authName $authName -authPwd $authPwd -uri $uri -requestHash $operation
}

Function New-Organization
{
    Param(
        [Parameter(Mandatory=$True)][string]$name,
        [Parameter(Mandatory=$True)][string]$code,
        [Parameter(Mandatory=$True)][string]$parent_name,
        [Parameter(Mandatory=$True)][string]$authName,
        [Parameter(Mandatory=$True)][string]$authPwd,
        [Parameter(Mandatory=$True)][string]$uri
    )
    $fields = New-Object PSObject -Property @{
        parent_id = "SELECT Organization WHERE name = `"$parent_name`""
        name = $name
        code = $code
    }
    
    $operation = New-Object PSObject -Property @{ 
        operation = 'core/create'
        class = 'Organization'
        comment = 'Synchronization from load scripts'
        output_fields = '*'
        fields = $fields
    }
    GenerateAndSendRequest -authName $authName -authPwd $authPwd -uri $uri -requestHash $operation
}

Function New-Brand
{
    Param(
        [Parameter(Mandatory=$True)][string]$name,
        [Parameter(Mandatory=$True)][string]$authName,
        [Parameter(Mandatory=$True)][string]$authPwd,
        [Parameter(Mandatory=$True)][string]$uri
    )
    $fields = New-Object PSObject -Property @{
        name = $name
    }
    
    $operation = New-Object PSObject -Property @{ 
        operation = 'core/create'
        class = 'Brand'
        comment = 'Synchronization from load scripts'
        output_fields = '*'
        fields = $fields
    }
    GenerateAndSendRequest -authName $authName -authPwd $authPwd -uri $uri -requestHash $operation
}

Function New-Model
{
    Param(
        [Parameter(Mandatory=$True)][string]$name,
        [Parameter(Mandatory=$True)][string]$type,
        [Parameter(Mandatory=$True)][string]$brand_id,
        [Parameter(Mandatory=$True)][string]$authName,
        [Parameter(Mandatory=$True)][string]$authPwd,
        [Parameter(Mandatory=$True)][string]$uri
    )
    $fields = New-Object PSObject -Property @{
        name = $name
        type = $type
        brand_id = $brand_id
    }
    
    $operation = New-Object PSObject -Property @{ 
        operation = 'core/create'
        class = 'Model'
        comment = 'Synchronization from load scripts'
        output_fields = '*'
        fields = $fields
    }
    GenerateAndSendRequest -authName $authName -authPwd $authPwd -uri $uri -requestHash $operation
}

Function New-OSFamily
{
    Param(
        [Parameter(Mandatory=$True)][string]$name,
        [Parameter(Mandatory=$True)][string]$authName,
        [Parameter(Mandatory=$True)][string]$authPwd,
        [Parameter(Mandatory=$True)][string]$uri
    )
    $fields = New-Object PSObject -Property @{
        name = $name
    }
    
    $operation = New-Object PSObject -Property @{ 
        operation = 'core/create'
        class = 'OSFamily'
        comment = 'Created via API call'
        output_fields = '*'
        fields = $fields
    }
    GenerateAndSendRequest -authName $authName -authPwd $authPwd -uri $uri -requestHash $operation
}

Function New-OSVersion
{
    Param(
        [Parameter(Mandatory=$True)][string]$name,
        [Parameter(Mandatory=$True)][string]$osfamily_name,
        [Parameter(Mandatory=$True)][string]$authName,
        [Parameter(Mandatory=$True)][string]$authPwd,
        [Parameter(Mandatory=$True)][string]$uri
    )
    $fields = New-Object PSObject -Property @{
        name = $name
        osfamily_id = "SELECT OSFamily WHERE name = `"$osfamily_name`""
    }
    
    $operation = New-Object PSObject -Property @{ 
        operation = 'core/create'
        class = 'OSVersion'
        comment = 'Created via API call'
        output_fields = '*'
        fields = $fields
    }
    GenerateAndSendRequest -authName $authName -authPwd $authPwd -uri $uri -requestHash $operation
}

Function New-VirtualFarm
{
    Param(
        [Parameter(Mandatory=$True)][string]$uuid,
        [Parameter(Mandatory=$True)][string]$name,
        [Parameter(Mandatory=$True)][string]$org_name,
        [Parameter(Mandatory=$True)][string]$authName,
        [Parameter(Mandatory=$True)][string]$authPwd,
        [Parameter(Mandatory=$True)][string]$uri
    )
    $fields = New-Object PSObject -Property @{
        org_id = "SELECT Organization WHERE name = `"$org_name`""
        name = $name
        uuid = $uuid
    }
    
    $operation = New-Object PSObject -Property @{ 
        operation = 'core/create'
        class = 'Farm'
        comment = 'Synchronization from vCenter'
        output_fields = '*'
        fields = $fields
    }

    GenerateAndSendRequest -authName $authName -authPwd $authPwd -uri $uri -requestHash $operation
}

Function New-Hypervisor
{
    Param(
        [Parameter(Mandatory=$True)][string]$name,
        [Parameter(Mandatory=$True)][string]$org_name,
        [Parameter(Mandatory=$False)][string]$farm_name,
        [Parameter(Mandatory=$True)][string]$authName,
        [Parameter(Mandatory=$True)][string]$authPwd,
        [Parameter(Mandatory=$True)][string]$uri
    )

    $fields = New-Object PSObject -Property @{
        org_id = "SELECT Organization WHERE name = `"$org_name`""
        name = $name
    }

    # add optional parameters
    if($farm_name -ne $null -and $farm_name -ne 'host' -and $farm_name -ne '')
    {
        $fields | Add-Member -NotePropertyName farm_id -NotePropertyValue "SELECT Farm WHERE name = `"$farm_name`""
    }
    
    $operation = New-Object PSObject -Property @{ 
        operation = 'core/create'
        class = 'Hypervisor'
        comment = 'Synchronization from vCenter'
        output_fields = '*'
        fields = $fields
    }

    GenerateAndSendRequest -authName $authName -authPwd $authPwd -uri $uri -requestHash $operation
}

Function New-Person
{
    Param(
        [Parameter(Mandatory=$True)][string]$firstName,
        [Parameter(Mandatory=$True)][string]$lastName,
        [Parameter(Mandatory=$True)][string]$email,
        [Parameter(Mandatory=$False)][string]$phone = $null,
        [Parameter(Mandatory=$True)][string]$orgName,
        [Parameter(Mandatory=$True)][string]$authName,
        [Parameter(Mandatory=$True)][string]$authPwd,
        [Parameter(Mandatory=$True)][string]$uri
    )

    $fields = New-Object PSObject -Property @{
        org_id = "SELECT Organization WHERE name = `"$orgName`""
        name = $lastName
        first_name = $firstName
        email = $email
    }

    # add optional parameters
    if(![String]::IsNullOrEmpty($phone))
    {
        $fields | Add-Member -NotePropertyName 'phone' -NotePropertyValue $phone
    }

    $operation = New-Object PSObject -Property @{ 
        operation = 'core/create'
        class = 'Person'
        comment = 'Created using API'
        output_fields = '*'
        fields = $fields
    }
    
    GenerateAndSendRequest -authName $authName -authPwd $authPwd -uri $uri -requestHash $operation
}


Function New-VirtualMachine
{
<#
 .Synopsis
  Creates a VirtualMachine object in iTop

 .Description
  Creates a VirtualMachine object in iTop

 .Parameter uui
  The unique ID of the VM

 .Parameter name
  The name of the VM

 .Parameter numCPU
  Number of virtual CPUs
 
 .Parameter ramGB
  Amount of memory in GB

 .Parameter hostName
  The hypervisor name, must already exist in iTop

 .Parameter orgName
  The organization name, must already exist in iTop

 .Parameter authName
  Logon for the iTop web service

 .Parameter authPwd
  Password for the iTop web service

 .Parameter uri
  uri for the iTop web service

 .Example
   Create-VirtualMachine -uuid '0001-0002-0003-00004' -name 'MyVM' -orgName 'Department A' -authName 'user' -authPwd 'password' -uri 'https://webservice.edu'

#>

    Param(
        [Parameter(Mandatory=$True)][string]$uuid,
        [Parameter(Mandatory=$True)][string]$name,
        [Parameter(Mandatory=$False)][string]$numCPU,
        [Parameter(Mandatory=$False)][string]$ramGB,
        [Parameter(Mandatory=$True)][string]$hostName,
        [Parameter(Mandatory=$True)][string]$orgName,
        [Parameter(Mandatory=$True)][string]$authName,
        [Parameter(Mandatory=$True)][string]$authPwd,
        [Parameter(Mandatory=$True)][string]$uri
    )

    $fields = New-Object PSObject -Property @{
        org_id = "SELECT Organization WHERE name = `"$orgName`""
        virtualhost_id = "SELECT VirtualHost WHERE name = `"$hostName`""
        name = $name
        uuid = $uuid
    }

    # add optional parameters
    if($numCPU -ne $null)
    {
        $fields | Add-Member -NotePropertyName 'cpu' -NotePropertyValue $numCPU
    }
    if($ramGB -ne $null)
    {
        $fields | Add-Member -NotePropertyName 'ram' -NotePropertyValue $ramGB
    }

    $operation = New-Object PSObject -Property @{ 
        operation = 'core/create'
        class = 'VirtualMachine'
        comment = 'Synchronization from vCenter'
        output_fields = '*'
        fields = $fields
    }
    
    GenerateAndSendRequest -authName $authName -authPwd $authPwd -uri $uri -requestHash $operation
}

Function Get-VirtualMachine
{
    Param(
        [Parameter(Mandatory=$True)][string]$authName,
        [Parameter(Mandatory=$True)][string]$authPwd,
        [Parameter(Mandatory=$True)][string]$uri,
        [Parameter(Mandatory=$False)][string]$oqlFilter,
        [Parameter(Mandatory=$False)][string]$outputFields='*'
    )

    Get-iTopObject -objectClass 'VirtualMachine' -ouputFields $outputFields -uri $uri -authName $authName -authPwd $authPwd -oqlFilter $oqlFilter
}


Function Set-FunctionalCI
{
    Param(
        [Parameter(Mandatory=$True)][string]$authName,
        [Parameter(Mandatory=$True)][string]$authPwd,
        [Parameter(Mandatory=$True)][string]$uri,
        [Parameter(Mandatory=$True)]$ci,
        [Parameter(Mandatory=$False)]$contacts
    )


    # build our linked objects, iTop only want certain lookup fields per object, so we'll feed those in
    $contacts_list = @()
    foreach($contact in $contacts)
    {
        $contactHash = @{}
        $contactHash.add("contact_id",("SELECT Contact WHERE id = `"$($contact.key)`""))
        $contacts_list += $contactHash
    }

    $fields = New-Object PSObject -Property @{
        contacts_list = $contacts_list
    }
    
    $operation = New-Object PSObject -Property @{ 
        operation = 'core/update'
        class = 'FunctionalCI'
        key = $ci.key
        comment = 'update from API'
        output_fields = '*'
        fields = $fields
    }
    GenerateAndSendRequest -authName $authName -authPwd $authPwd -uri $uri -requestHash $operation
}

Function Get-Service
{
    Param(
        [Parameter(Mandatory=$True)][string]$authName,
        [Parameter(Mandatory=$True)][string]$authPwd,
        [Parameter(Mandatory=$True)][string]$uri,
        [Parameter(Mandatory=$False)][string]$oqlFilter,
        [Parameter(Mandatory=$False)][string]$outputFields='*'
    )

    Get-iTopObject -objectClass 'Service' -ouputFields $outputFields -uri $uri -authName $authName -authPwd $authPwd -oqlFilter $oqlFilter

}

Function Get-SLA
{
    Param(
        [Parameter(Mandatory=$True)][string]$authName,
        [Parameter(Mandatory=$True)][string]$authPwd,
        [Parameter(Mandatory=$True)][string]$uri,
        [Parameter(Mandatory=$False)][string]$oqlFilter,
        [Parameter(Mandatory=$False)][string]$outputFields='*'
    )

    Get-iTopObject -objectClass 'SLA' -ouputFields $outputFields -uri $uri -authName $authName -authPwd $authPwd -oqlFilter $oqlFilter
}

Function New-VirtualMachineReplica
{
    Param(
        [Parameter(Mandatory=$True)][string]$uuid,
        [Parameter(Mandatory=$True)][string]$name,
        [Parameter(Mandatory=$False)][string]$cpu,
        [Parameter(Mandatory=$False)][string]$ram,
        [Parameter(Mandatory=$True)][string]$virtualhost_id,
        [Parameter(Mandatory=$True)][string]$orgKey,
        [Parameter(Mandatory=$True)][object]$datasourceObject,
        [Parameter(Mandatory=$True)][string]$ITOP_DB_Server,
        [Parameter(Mandatory=$True)][string]$ITOP_DB_authName,
        [Parameter(Mandatory=$True)][string]$ITOP_DB_authPwd,
        [Parameter(Mandatory=$True)][string]$ITOP_DB_Name
    )

    $synchroTableName = $datasourceObject.database_table_name

    # Update VM synchro data
    $updateStatement = "INSERT INTO $synchroTableName (primary_key,uuid,virtualhost_id,name,org_id,business_criticity,status,cpu,ram) 
                        VALUES (`'$uuid`',`'$uuid`',`'$virtualhost_id`',`'$name`',$orgKey,`'$businessCriticity`',`'$status`',`'$cpu`',`'$ram`') 
                        ON DUPLICATE KEY UPDATE 
                            name=VALUES(name), 
                            org_id=VALUES(org_id),
                            virtualhost_id=VALUES(virtualhost_id), 
                            cpu=VALUES(cpu),
                            ram=VALUES(ram)"

    $res = Invoke-NonQuery -serverName $ITOP_DB_Server -userName $ITOP_DB_authName -password $ITOP_DB_authPwd -dbName $ITOP_DB_Name -query $updateStatement
    $res
}

<### HELPERS ###>

Function Get-iTopObject
{
<#
 .Synopsis
  Retrieves an iTop object.

 .Description
  Private helper function.  All Get-* commandlets go here.  This is a generic handler for getting iTop objects using the REST service and an OQL query. 

 .Parameter oqlFilter
  Optional WHERE clause of an OQL query to filter results.

 .Parameter objectClass
  The type iTop object to return.

 .Parameter ouputFields
  Fields to return.  By default returns all fields.

 .Parameter authName
  Logon for the iTop web service

 .Parameter authPwd
  Password for the iTop web service

 .Parameter uri
  uri for the iTop web service
#>

    Param(
        [Parameter(Mandatory=$False)][string]$oqlFilter,
        [Parameter(Mandatory=$True)][string]$objectClass,
        [Parameter(Mandatory=$False)][string]$ouputFields='*',
        [Parameter(Mandatory=$True)][string]$authName,
        [Parameter(Mandatory=$True)][string]$authPwd,
        [Parameter(Mandatory=$True)][string]$uri
    )

    $query = "SELECT $objectClass"

    # Calculate our OQL query
    if(![string]::IsNullOrEmpty($oqlFilter))
    {
        $query = $query + " $oqlFilter"
    }

    # Create a hash/psobject of our request
    $operation = New-Object PSObject -Property @{
        operation = 'core/get'
        class = $objectClass
        key = $query
        output_fields = $ouputFields
    }

    # Format the request and send it to iTop
    GenerateAndSendRequest -authName $authName -authPwd $authPwd -uri $uri -requestHash $operation
}

Function New-iTopObject
{
}

<#
 .Synopsis
  Invoke the synchro exec over the web.

 .Description
  Invoke the synchro exec over the web.  Expects the URI as something like https://myserver.com/itop/synchro/synchro_exec.php

 .Parameter dataSource
  The datasource object to run the synchro

 .Parameter authName
  Logon for the iTop web service

 .Parameter authPwd
  Password for the iTop web service

 .Parameter uri
  uri for the iTop web service
#>

Function Invoke-SynchroExec
{

    Param(
        [Parameter(Mandatory=$True)]$dataSource,
        [Parameter(Mandatory=$TRUE)]$authName,
        [Parameter(Mandatory=$TRUE)]$authPwd,
        [Parameter(Mandatory=$TRUE)]$uri
    )

    # build our paramter hash
    $requestBody =  @{
            data_sources=$dataSource.key
            auth_user="$authName"
            auth_pwd="$authPwd"
    }

    # send the web request and get the response
    $res = Invoke-WebRequest -Method Post -Uri $uri -Body $requestBody -TimeoutSec 960
    if($res.StatusCode -eq 200)
    {
        $res
    }
    else
    {
        Throw "Synchro_exec returned an error when running $res"
    }
}

Function Invoke-SyncScript
{
    Param(
        [Parameter(Mandatory=$TRUE)]$datasourceObject,
        [Parameter(Mandatory=$TRUE)]$ITOP_SSH_Username,
        [Parameter(Mandatory=$TRUE)]$ITOP_SSH_Key,
        [Parameter(Mandatory=$TRUE)]$ITOP_SSH_Passphrase,
        [Parameter(Mandatory=$TRUE)]$ITOP_authName,
        [Parameter(Mandatory=$TRUE)]$ITOP_authPwd,
        [Parameter(Mandatory=$TRUE)]$ITOP_sync_exe_path,
        [Parameter(Mandatory=$TRUE)]$ITOP_server_name
    )
    
    $session = New-SshSession -ComputerName $ITOP_server_name -Username $ITOP_SSH_Username -KeyFile $ITOP_SSH_Key -Passphrase $ITOP_SSH_Passphrase
    $command = 'php -q ' + $ITOP_sync_exe_path + ' --auth_user=' + $ITOP_authName + ' --auth_pwd=' + $ITOP_authPwd + ' --data_sources=' + $datasourceObject.key
    $sshRes = Invoke-SshCommand -ComputerName $ITOP_server_name -Command $command
    $sshRes
}

Function InsertObjectIntoSynchroDataSource
{
<# work in progress to create a generic data source updater #>
    Param(
        [Parameter(Mandatory=$True)][object]$itopObject,
        [Parameter(Mandatory=$True)][object]$datasourceObject,
        [Parameter(Mandatory=$True)][string]$ITOP_DB_Server,
        [Parameter(Mandatory=$True)][string]$ITOP_DB_authName,
        [Parameter(Mandatory=$True)][string]$ITOP_DB_authPwd,
        [Parameter(Mandatory=$True)][string]$ITOP_DB_Name,
        [Parameter(Mandatory=$True)][string]$ITOP_SYNC_TableName,
        [Parameter(Mandatory=$True)][string]$primaryKeyField
    )


    #TODO, read the columns from the source datasource, then use those to build our statment, duh!
    $source = Get-SynchroDataSource -authName $


    $updatePart = "INSERT INTO $ITOP_SYNC_TableName ("
    $valuesPart = "VALUES ("
    $onDupePart = "ON DUPLICATE KEY UPDATE "

    $itopObject | Get-Member | Where {$_.MemberType -eq "NoteProperty"} | Select -Property Name | % {
        $thisProp = $_.Name
        if(![String]::IsNullOrEmpty($itopObject.$thisProp))
        {
            $updatePart += $_.Name + ','
            $valuesPart += "`'$($itopObject.$thisProp)`',"
            $onDupePart += "$thisProp=VALUES($thisProp),"
        }
    }

    $updatePart += 'primary_key'
    $valuesPart += $itopObject.$primaryKeyField
    $onDupePart = $onDupePart.Substring(0,$onDupePart.Length -1)
    $updateStatement = $updatePart + $valuesPart + $onDupePart

    # add the primary_key part

    <#
    
                # Update VM synchro data
                $updateStatement = "INSERT INTO $synchroTableName (primary_key,volume_id,virtualdevice_id,size_used) 
                                    VALUES (`'$primary_key`',`'$volume_id`',`'$virtualdevice_id`',$size_used) 
                                    ON DUPLICATE KEY UPDATE 
                                        volume_id=VALUES(volume_id),
                                        virtualdevice_id=VALUES(virtualdevice_id),
                                        size_used=VALUES(size_used)"

                $res = Invoke-NonQuery -serverName $ITOP_DB_Server -userName $ITOP_DB_authName -password $ITOP_DB_authPwd -dbName $ITOP_DB_Name -query $updateStatement
    #>
}

Function GenerateAndSendRequest
{
<#
 .Synopsis
  Format an iTop request and hadle the result.

 .Description
  Private helper function.  Any interaction with iTop Get-/Set- should go through this.

 .Parameter requestHash
  A hashtable of the request to be converted to JSON

 .Parameter authName
  Logon for the iTop web service

 .Parameter authPwd
  Password for the iTop web service

 .Parameter uri
  uri for the iTop web service
#>
    Param(
        [Parameter(Mandatory=$True)][string]$authName,
        [Parameter(Mandatory=$True)][string]$authPwd,
        [Parameter(Mandatory=$True)][string]$uri,
        [Parameter(Mandatory=$True)]$requestHash
    )

    $requestBody =  @{
            version="1.1"
            auth_user="$authName"
            auth_pwd="$authPwd"
            json_data=ConvertTo-Json($requestHash) -Depth 10
    }

    $result = Invoke-RestMethod -Method Post -Uri $uri -Body $requestBody -TimeoutSec 480

    # Is there an unexpected exception?
    if($result.code -ne 0)
    {
        throw "Result code = $($result.code), $($result.message)"
    }

    # If we have empty objects that simply means the search worked but came up empty for some reason
    if($result.objects.count -eq 0)
    {
        $output = $null
    }
    else
    {
        # Convert the returned objects into a hashtable to deal with goofy JSON stuff
        foreach ($myPsObject in $result.objects) 
        { 
            $output = @{}; 
            $myPsObject | Get-Member -MemberType *Property | % { $output.($_.name) = $myPsObject.($_.name); }
        }
    }

    # Some other error occured
    if($output -ne $null  -and $output.Values[0].code -ne 0)
    {
        throw "Result code = $($output.Values[0].code), $($output.Values[0].message)"
    }

    # Looks like we have at least one object to return, let's create a nicer PSObject with the id
    # With latest version of iTop we now have access to a proper key field, let's add that as well for 
    # compatibility
    foreach ($key in $output.Keys)
    {
        $thisObject = $output.Item($key).fields
        if($output.Item($key).key -ne $null)
        {
            # using the newer API, let's add the key
            $thisObject | Add-Member Noteproperty -Name "key" -Value $output.Item($key).key
        }
        # put the object on the pipline
        $thisObject
    }
}


Export-ModuleMember -Function Get-Brand
Export-ModuleMember -Function Get-Contact
Export-ModuleMember -Function Get-CustomerContract
Export-ModuleMember -Function Get-Enclosure
Export-ModuleMember -Function Get-Hypervisor
Export-ModuleMember -Function Get-LinkServiceToSLA
Export-ModuleMember -Function Get-LogicalVolume
Export-ModuleMember -Function Get-Model
Export-ModuleMember -Function Get-Organization
Export-ModuleMember -Function Get-OSFamily
Export-ModuleMember -Function Get-OSVersion
Export-ModuleMember -Function Get-Person
Export-ModuleMember -Function Get-Server
Export-ModuleMember -Function Get-Service
Export-ModuleMember -Function Get-SLA
Export-ModuleMember -Function Get-StorageSystem
Export-ModuleMember -Function Get-SynchroDataSource
Export-ModuleMember -Function Get-VirtualFarm
Export-ModuleMember -Function Get-VirtualMachine


Export-ModuleMember -Function Invoke-SyncScript
Export-ModuleMember -Function Invoke-SynchroExec

Export-ModuleMember -Function New-Brand
Export-ModuleMember -Function New-CustomerContract
Export-ModuleMember -Function New-Hypervisor
Export-ModuleMember -Function New-OSFamily
Export-ModuleMember -Function New-OSVersion
Export-ModuleMember -Function New-VirtualFarm
Export-ModuleMember -Function New-VirtualMachine
Export-ModuleMember -Function New-VirtualMachineReplica
Export-ModuleMember -Function New-Model
Export-ModuleMember -Function New-Organization
Export-ModuleMember -Function New-Person

Export-ModuleMember -Function Set-FunctionalCI

Export-ModuleMember -Function InsertObjectIntoSynchroDataSource