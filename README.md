# Q-extractor
VBA data extractor for capability studies

Capability studies like Cp/Cpk require randomly selected data subset of production data of defined subgroup size and samples count or daily base count.
It is very time-consuming to fulfill such requirements manually. Therefore, script automation is desirable.

Q-extractor is Excel macro (VBA) automation solution. There are three lists in a workbook.
1. Settings - see the picture below
2. Input - fill it with whole data set
3. Output - data for capability SW

![obrazek](https://user-images.githubusercontent.com/3974820/190404266-f32756f7-da1e-42ba-ba3e-66bfca83b248.png)
 
 ThisWorkbook macro clear In/Output sheets
 Sheet1 macro automate the process of data selection
