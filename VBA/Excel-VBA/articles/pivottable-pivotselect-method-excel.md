---
title: PivotTable.PivotSelect Method (Excel)
keywords: vbaxl10.chm235137
f1_keywords:
- vbaxl10.chm235137
ms.prod: excel
api_name:
- Excel.PivotTable.PivotSelect
ms.assetid: e9beda74-c022-3ba7-b3af-d607024846f2
ms.date: 06/08/2017
---


# PivotTable.PivotSelect Method (Excel)

Selects part of a PivotTable report.


## Syntax

 _expression_ . **PivotSelect**( **_Name_** , **_Mode_** , **_UseStandardName_** )

 _expression_ A variable that represents a **PivotTable** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The part of the PivotTable report to select.|
| _Mode_|Optional| **[XlPTSelectionMode](xlptselectionmode-enumeration-excel.md)**|Specifies the structured selection mode.|
| _UseStandardName_|Optional| **Variant**| **True** for recorded macros that will play back in other locales.|

## Remarks

You can use the specified mode only to select the corresponding item in the PivotTable report. For example, you cannot select data and labels by using  **xlButton** mode; likewise, you cannot select buttons by using **xlDataOnly** mode.


## Example

This example selects all date labels in the first PivotTable report on worksheet one.


```vb
Worksheets(1).PivotTables(1).PivotSelect "date[All]", xlLabelOnly
```

In the following example, 
* A list of levels are specified
* Level unique names are quoted, followed by brackets.
* At the before the closing bracket, qualifiers may be added
  * Data to include data for the level
  * Totals to include totals for the level
* Only the last qualifier specified for a level is observed
* Levels on the same axis union
* Levels on opposing axis intersect
* Each level is separated by a space
* Within brackets, included member keys can be listed
* , separates members in the list.
* : acts as a range operator
* Values are specified using a Values level.

```vb
Worksheets(1).PivotTables(1)..PivotSelect _
  "'[Date].[Calendar Year].[Calendar Year]'[" & _
    "'[Date].[Calendar Year].&[2011]':'[Date].[Calendar Year].&[2012]'" & _
  ";Data] " & _
  "'[Date].[Calendar Year].[Calendar Year]'[" & _
    "'[Date].[Calendar Year].&[2014]' " & _
  "] " & _
  "'[Date].[Month Name].[Month Name]'[" & _
    "'[Date].[Month Name].&[January]'," & _
    "'[Date].[Month Name].&[March]'" & _
  "] " & _
  "'[Geography].[Country Region Name].[Country Region Name]'[" & _
    "'[Geography].[Country Region Name].&[Australia]'," & _
    "'[Geography].[Country Region Name].&[France]'" & _
  ";Data;Total] " & _
  "Values[" & _
    "'[Measures].[Internet Total Sales]'," & _
    "'[Measures].[Internet Total Units]'" & _
  "] ", _
  xlDataOnly

```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

