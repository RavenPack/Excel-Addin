# Excel-Addin

Excel Addin for retrieving analytics via the RavenPack API

For more details on the RavenPack REST API, please see the [API documention.](https://app.ravenpack.com/help/) on the RavenPack website.

In order to use the RavenPack Excel Addin, you must have a valid API_KEY. To request a key or for more information, please contact client services by emailing, [support@ravenpack.com](mailto:support@ravenpack.com).

## Running the Addin

The RavenPack Excel Addin is distributed as a macro-enabled spreadsheet and was optimized for Office 2010 forward. To install, simply download the latest version from the dist folder and run on your local machine. You will need to Enable Macros. Additionally, you may be asked to Enable Editing and Enable Content. You can verify that everything has been setup properly by checking that the RavenPack tab is now visible in your ribbon at the top.

## Usage

There are two ways to use the Excel Plugin.

### RavenPack Ribbon

Once the Addin has been correctly run, you will see the RavenPack tab on your ribbon. If you click on this, you will be presented with a number of different options that will manipulate the data in your spreadsheet.

![RavenPack Excel Ribbon](https://raw.githubusercontent.com/RavenPack/Excel-Addin/master/resources/excel_ribbon.png)

* Server Status: Check that you are able to correctly connect to the RavenPack REST API.
* Set API_KEY: Set your API_KEY for subsequent requests.
* List Datasets: List the datasets that you have configured on the RavenPack platform.
* Data Request: Request data from a predefined dataset. The data will be placed in the active sheet.
* Map Entities: Map indentifying information from your spreadsheet to RavenPack entities.
* Reference Data: Request reference information for any of the entities supported by RavenPack Analytics.
* Event Taxonomy: Request the full event taxonomy supported by RavenPack Analytics.
* Function Library: View documentation about available functions.

### Functions

A number of functions are available that may be called from any cell in the spreadsheet.

* RPMapEntity: Return an RP_ENTITY_ID, given a set of identifying information.
* RPEntityName: Return an entity name, given an RP_ENTITY_ID.
* RPGetDailySentiment: Return the daily sentiment on a given day for a given RP_ENTITY_ID.
* RPGetDailyBuzz: Return the daily media buzz on a given day for a given RP_ENTITY_ID.
* RPGetDailyVolume: Return the daily media volume on a given day for a given RP_ENTITY_ID.
* RPGetDailyValue: Return a value from a given dataset, on a given day for a given RP_ENTITY_ID.
* RPGetRecordCount: Return the number of records in a given dataset for a given period of time.

Each of the functions requires an API_KEY in order to run.
