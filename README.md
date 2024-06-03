# Manage-MSLicense
This PowerShell script helps to manage license assign and removal in a Microsoft 365 tenant. Please read the synopsis of the script for more details.

This script assigns and removes Microsoft 365 licenses to users and groups in your tenant. More precisely, it prompts you interactively a list of the options to allow easily to choose the license and the relative service plans to add / remove, using the friendly name of the product (e.g., you find "Microsoft 365 E5" instead of "SPE_E5" or the GUID of the license).
The script needs the .csv file "LicenseMappingTable.csv" which lists all the Microsft licenses with the GUID and the available Service Plans, that must be saved in the same folder of the script. The file has buin built based on the table reported here https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference.

Written by: Stefano Viti - stefano.viti1995@gmail.com
Follow me at https://www.linkedin.com/in/stefano-viti/
