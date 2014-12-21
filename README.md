### ifGenerator
PowerShell utility to create horrible, horrible =IF() statement chains for Sharepoint List Derived Columns, or Excel.

Splits output into groups of 4+remainder to allow as deep as an IF statement chain as needed.

Usage:  
Read the **IMPORTANT** heading below -- you must modify the ```$ifType``` variable in the code for your needs.

Example output:

```
=IF((A1=11),"item11Name",  
IF((A1=10),"item10Name",  
IF((A1=9),"item9Name",  
IF((A1=8),"item8Name",  
("")))))  
&  
IF((A1=7),"item7Name",  
IF((A1=6),"item6Name",  
IF((A1=5),"item5Name",  
IF((A1=4),"item4Name",  
("")))))  
&  
IF((A1=3),"item3Name",  
IF((A1=2),"item2Name",  
IF((A1=1),"item1Name",  
(""))))  
```

####  IMPORTANT  
** *This information appears in the ifGenerator.ps1 code as well* **  

In the code, leave only one of the ```$ifType``` variables (lines) uncommented. Excel and Sharepoint derived
columns use different formats. In addition, the specific cell or column name needs to be changed inside the ```$ifType``` variable string.

#### For Excel  
Replace `A1` in the ```$ifType``` string with the Excel cell in question:

*&#42;&#42;&#42; Comment out this line if creating a Sharepoint =IF()*

```
$ifType = "IF((A1=" # Uncomment this line for Excel. 
```


#### For Sharepoint  
Replace `item#` in the string below with the Sharepoint derived column name in question:

*&#42;&#42;&#42; Comment out this line if creating an Excel =IF()*

```
$ifType = "IF(([item#]=" # Uncomment this line for Sharepoint. 
```  
