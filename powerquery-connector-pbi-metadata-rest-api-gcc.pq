/*
    Author: John Kerski

This connector wraps a few of the "Get" endpoints in the Power BI Metadata API, so that OAuth can be used to 
authenticate to the service.  
   
The endpoints for this custom connector on NOT part of the v1.0 Microsoft Power REST APIS. Microsoft lacks 
API support for dataflows.  As of July 2023, Microsoft does not provide APIs for:

    1) Error Details in Dataflow Refreshes
    2) Whether a Dataflow is Scheduled for a Refresh
    3) The details of the refresh schedule

*This lack of support severely hinders the ability to monitor for issues.*

This connector serves as a way to have a library of Power Query functions to build datasets 
without the need for storing client secrets or passwords in the dataset.  The custom data connector 
will need to be installed in the a Power BI Gateway in order to refresh datasets leveraging this custom connector.

*/
[Version = "1.0.1"]
section PBIMETADATARESTAPIGCC;
[
    DataSource.Kind = "PBIMETADATARESTAPIGCC",
    Publish = "PBIMETADATARESTAPIGCC.Publish"
]
/*****************
    Navigation Table
*****************/
shared PBIMETADATARESTAPIGCC.FunctionCatalog = () =>
let
    objects = #table(
            {"Name","Key","Data","ItemKind","ItemName","IsLeaf"},{
            /* Dataflows */            
            {"GetDataflowGallery","GetDataflowGallery",PBIMETADATARESTAPIGCC.GetDataflowGallery,"Function","GetDataflowGallery",true},                                                      
            {"GetDataflowTranscationDetails","GetDataflowTranscationDetails",PBIMETADATARESTAPIGCC.GetDataflowTranscationDetails,"Function","GetDataflowTranscationDetails",true},                                                      
            {"GetDataflowExtendedSettings","GetDataflowExtendedSettings",PBIMETADATARESTAPIGCC.GetDataflowExtendedSettings,"Function","GetDataflowExtendedSettings",true}    
        }),
    navTable = Table.ForceToNavigationTable(objects, {"Key"},"Name","Data","ItemKind","ItemName","IsLeaf")
in
    navTable;

/*****************
    Dataflows
*****************/
/*** GetDataflowGallery**/
shared PBIMETADATARESTAPIGCC.GetDataflowGallery = Value.ReplaceType(
    GetDataflowGallery,
    GetDataflowGalleryType
);

GetDataflowGalleryType = type function (
) as table meta [
    Documentation.Name = "PBIMETADATARESTAPIGCC.GetDataflowGallery",
    Documentation.LongDescription = "Returns a list of all dataflows and properties from the metadata API.",
    Documentation.Examples = {
        [
            Code = "=PBIMETADATARESTAPIGCC.GetDataflowGallery()",
            Result = "Non-traditional API endpoint"
        ]
    }
];

GetDataflowGallery = () =>
    let
        result = GetMetadataJSON("metadata/gallery/dataflows")
    in
        result;


/*** GetDataflowTranscationDetails**/
shared PBIMETADATARESTAPIGCC.GetDataflowTranscationDetails = Value.ReplaceType(
    GetDataflowTranscationDetails,
    GetDataflowTranscationDetailsType
);

GetDataflowTranscationDetailsType = type function (
    optional transactionID as (
        type text meta [
            Documentation.FieldCaption = "Refresh ID",
            Documentation.FieldDescription = "Refresh ID",
            Documentation.SampleValues = {
                "2023-06-03T13:10:53.4023962Z@36c4acbe-e906-41c9-b8a2-dd452ceee2c2"
            }
        ]
    )
) as table meta [
    Documentation.Name = "PBIMETADATARESTAPIGCC.GetDataflowTranscationDetails",
    Documentation.LongDescription = "Returns a list of transactions details for a specific dataflow refresh based on the Refresh ID.",
    Documentation.Examples = {
        [
            Code = "=PBIMETADATARESTAPIGCC.GetDataflowTransactionsInGroup(""2023-06-03T13:10:53.4023962Z@36c4acbe-e906-41c9-b8a2-dd452ceee2c2"")",
            Result = "Non-traditional API endpoint"
        ]
    }
];

GetDataflowTranscationDetails= (optional refreshId as text) =>
    let
        result = GetMetadataJSON("/metadata/dataflows/transactions/" & refreshId & "/refreshhistory")
    in
        result;

/*** GetDataflowExtendedSettings **/
shared PBIMETADATARESTAPIGCC.GetDataflowExtendedSettings = Value.ReplaceType(
    GetDataflowExtendedSettings,
    GetDataflowExtendedSettingsType
);

GetDataflowExtendedSettingsType = type function (
    optional transactionID as (
        type text meta [
            Documentation.FieldCaption = "Internal Dataflow ID",
            Documentation.FieldDescription = "Internal Dataflow ID",
            Documentation.SampleValues = {
                "123456"
            }
        ]
    )
) as table meta [
    Documentation.Name = "PBIMETADATARESTAPIGCC.GetDataflowExtendedSettings",
    Documentation.LongDescription = "Returns a list of details for a specific dataflow.",
    Documentation.Examples = {
        [
            Code = "=PBIMETADATARESTAPIGCC.GetDataflowExtendedSettings(""123456"")",
            Result = "Non-traditional API endpoint"
        ]
    }
];

GetDataflowExtendedSettings = (optional dataflowId as text) =>
    let
        result = GetMetadataJSON("/metadata/dataflows/" & dataflowId & "/extendedSettings")
    in
        result;        


/*****************
    HELPER FUNCTIONS
*****************/
GetMetadataJSON = (optional relPath as text) =>
    let
        params = [Headers = [#"Content-Type" = "application/json"], RelativePath = relPath],
        //Get Pages
        //Hard-coded Base URL because Power BI Service won't recognize parameter despite numerous attempts
        webSource = Web.Contents(
            "https://wabi-us-gov-virginia-redirect.analysis.usgovcloudapi.net",
            params
        ),
        //Convert to Json
        result = Text.FromBinary(webSource)
    in
        result;


// Data Source Kind description
PBIMETADATARESTAPIGCC = [
    TestConnection = (dataSourcePath) =>  { "PBIMETADATARESTAPIGCC.FunctionCatalog"},
    Authentication = [
        Aad = [
            AuthorizationUri = "https://login.windows.net/common/oauth2/authorize",
            Resource = "https://analysis.usgovcloudapi.net/powerbi/api"
        ]
    ]
];

// Data Source UI publishing description
PBIMETADATARESTAPIGCC.Publish = [
    Beta = true,
    Category = "Other",
    ButtonText = {
        Extension.LoadString(
            "ButtonTitle"
        ),
        Extension.LoadString(
            "ButtonHelp"
        )
    },
    LearnMoreUrl = "https://github.com/kerski/powerquery-connector-pbi-metadata-rest-api-gcc/blob/main/README.md",
    SourceImage = PBIMETADATARESTAPIGCC.Icons,
    SourceTypeImage = PBIMETADATARESTAPIGCC.Icons
];

PBIMETADATARESTAPIGCC.Icons = [
    Icon16 = {
        Extension.Contents(
            "powerquery-connector-pbi-metadata-rest-api-gcc16.png"
        ),
        Extension.Contents(
            "powerquery-connector-pbi-metadata-rest-api-gcc20.png"
        ),
        Extension.Contents(
            "powerquery-connector-pbi-metadata-rest-api-gcc24.png"
        ),
        Extension.Contents(
            "powerquery-connector-pbi-metadata-rest-api-gcc32.png"
        )
    },
    Icon32 = {
        Extension.Contents(
            "powerquery-connector-pbi-metadata-rest-api-gcc32.png"
        ),
        Extension.Contents(
            "powerquery-connector-pbi-metadata-rest-api-gcc40.png"
        ),
        Extension.Contents(
            "powerquery-connector-pbi-metadata-rest-api-gcc48.png"
        ),
        Extension.Contents(
            "powerquery-connector-pbi-metadata-rest-api-gcc64.png"
        )
    }
];

//To Navigation Table
Table.ToNavigationTable = (
    table as table,
    keyColumns as list,
    nameColumn as text,
    dataColumn as text,
    itemKindColumn as text,
    itemNameColumn as text,
    isLeafColumn as text
) as table =>
    let
        tableType = Value.Type(
            table
        ),
        newTableType = Type.AddTableKey(
            tableType,
            keyColumns,
            true
        ) meta [
            NavigationTable.NameColumn = nameColumn,
            NavigationTable.DataColumn = dataColumn,
            NavigationTable.ItemKindColumn = itemKindColumn,
            Preview.DelayColumn = itemNameColumn,
            NavigationTable.IsLeafColumn = isLeafColumn
        ],
        navigationTable = Value.ReplaceType(
            table,
            newTableType
        )
    in
        navigationTable;

// Force Navigation Table
Table.ForceToNavigationTable = (
    table as table,
    keyColumns as list,
    nameColumn as text,
    dataColumn as text,
    itemKindColumn as text,
    itemNameColumn as text,
    isLeafColumn as text
) as table =>
    let
        tableType = Value.Type(
            table
        ),
        newTableType = Type.AddTableKey(
            tableType,
            keyColumns,
            true
        ) meta [
            NavigationTable.NameColumn = nameColumn,
            NavigationTable.DataColumn = dataColumn,
            NavigationTable.ItemKindColumn = itemKindColumn,
            NavigationTable.IsLeafColumn = isLeafColumn
        ],
        navigationTable = Value.ReplaceType(
            table,
            newTableType
        )
    in
        navigationTable;