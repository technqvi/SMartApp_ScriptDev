# About
There are serveral modules, all of them have been developed to support [SMartApp](https://github.com/technqvi/SMartApp) System.


### [SMart_SiteGrade_Report](https://github.com/technqvi/SMartApp_ScriptDev/tree/main/SMart_SiteGrade_Report)
* Retrieve the following data such as storage, server,	software, network,	incident case, and request case of each company that Site managers are in charge of.
* Take data to calculate the score and weight as the level of given ranges to find rank
* After that, the team lead will take the final score calculated from prev step to assign a site manager in the team to take care of the customer's project proportionally based on a quarterly and yearly basis.

## [InventoryImportApp](https://github.com/technqvi/SMartApp_ScriptDev/tree/main/InventoryImportApp)
* Click Export inventory from Inventory Management(http://essm.yipintsoi.com/inventories/).
* Use recently exported file as  template to add new inventory on excel.
* Run this script to add tons of inventories to database once as [doc-usermanual](https://github.com/technqvi/InventoryImportApp/tree/master/doc-usermanual).
* All reference data like brand ,model, product type will use name to find foreign-key automatically.

## [SmartExcelReport](https://github.com/technqvi/SMartApp_ScriptDev/tree/main/SmartExcelReport)
Retrieve incident data from postgresql database to transform data in order to create several new columns on data frame and save as excel file report.
* Excel Table Report
* Excel Pivot Report
### [EmployessFiles](https://github.com/technqvi/SMartApp_ScriptDev/tree/main/EmployessFiles)
Import employees into app_employee table in SMartApp

### [SM-AdminCompany](https://github.com/technqvi/SMartApp_ScriptDev/tree/main/SM-AdminCompany)
Import company for adminstrator manaagement to SMartApp as well as related data.
