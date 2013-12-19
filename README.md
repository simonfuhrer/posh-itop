posh-itop
=========

A PowerShell wrapper for Combodo's iTop CMDB

Requirements:

- POSH-MySQL
- iTop version 2.0.2 (REST API 1.1)

Notes:

I you want to use this please look at the following two cmdlets first

Get-iTopObject
GenerateAndSendRequest

Those are the main wrappers for the web services calls.  The JSON returned from the iTop rest service is wrapped up in a nice PowerShell object for consumption and for feeding into other cmdlets.

This is designed to be an easy to use PowerShell wrapper for the iTop REST API.  It's currently most useful for the CMDB (Configuration Management) and Service Management modules in iTop because that's all I'm using.  It can also invoke the data synchronization routines required when automating your CMDB.  


