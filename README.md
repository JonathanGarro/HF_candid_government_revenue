# Fetch Grantee Government Revenue Data from Candid API

This script accepts an Excel file with a list of EIN (Employer Identification Number) and sends each one to the Candid API to fetch financial information. The API has dozens of financial fields, but this script is focused on the total revenue, the revenue from contributions, and the revenue from government sources. 

A Candid Premiere v3 token is required. To save API calls, you should only pass a list of unique EINs. If your final analytical product will include duplicate organizations, use a lookup within Excel to populate. 