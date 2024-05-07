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
[Version = "1.0.4"]
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
            /* Apps */
            {"GetAppDetails","GetAppDetails",PBIMETADATARESTAPIGCC.GetAppDetails,"Function","GetAppDetails",true},
            /* Dataflows */            
            {"GetDataflowGallery","GetDataflowGallery",PBIMETADATARESTAPIGCC.GetDataflowGallery,"Function","GetDataflowGallery",true},                                                      
            {"GetDataflowTranscationDetails","GetDataflowTranscationDetails",PBIMETADATARESTAPIGCC.GetDataflowTranscationDetails,"Function","GetDataflowTranscationDetails",true},                                                      
            {"GetDataflowExtendedSettings","GetDataflowExtendedSettings",PBIMETADATARESTAPIGCC.GetDataflowExtendedSettings,"Function","GetDataflowExtendedSettings",true},    
            /* Pipelines */
            {"GetALMPipelines","GetALMPipelines",PBIMETADATARESTAPIGCC.GetALMPipelines,"Function","GetALMPipelines",true},    
            {"GetALMPipelineOperation","GetALMPipelineOperation",@PBIMETADATARESTAPIGCC.GetALMPipelineOperation,"Function","GetALMPipelineOperation",true},
            {"GetALMPipelineStageDifferences","GetALMPipelineStageDifferences",@PBIMETADATARESTAPIGCC.GetALMPipelineStageDifferences,"Function","GetALMPipelineStageDifferences",true},
            {"DiscoverArtifacts","DiscoverArtifacts",PBIMETADATARESTAPIGCC.DiscoverArtifacts,"Function","DiscoverArtifacts",true},    
            /* Datasets */
            {"GetSharedDatasets","GetSharedDatasets",PBIMETADATARESTAPIGCC.GetSharedDatasets,"Function","GetSharedDatasets",true},   
            {"GetOneDriveRefreshEnforcementPolicy","GetOneDriveRefreshEnforcementPolicy",PBIMETADATARESTAPIGCC.GetOneDriveRefreshEnforcementPolicy,"Function","GetOneDriveRefreshEnforcementPolicy",true}
        }),
    navTable = Table.ForceToNavigationTable(objects, {"Key"},"Name","Data","ItemKind","ItemName","IsLeaf")
in
    navTable;

/*****************
    Apps
*****************/
/*** GetAppDetails **/
shared PBIMETADATARESTAPIGCC.GetAppDetails = Value.ReplaceType(
    GetAppDetails,
    GetAppDetailsType
);

GetAppDetailsType = type function (
    optional transactionID as (
        type text meta [
            Documentation.FieldCaption = "Group ID",
            Documentation.FieldDescription = "Group ID",
            Documentation.SampleValues = {
                "xxx-yyyy"
            }
        ]
    )
) as table meta [
    Documentation.Name = "PBIMETADATARESTAPIGCC.GetAppDetails",
    Documentation.LongDescription = "Returns details on an App based on the Workspace id that hosts that App.",
    Documentation.Examples = {
        [
            Code = "=PBIMETADATARESTAPIGCC.GetAppDetails(""xxx-yyyy"")",
            Result = "Non-traditional API endpoint"
        ]
    }
];

GetAppDetails = (optional workspaceId as text) =>
    let
        result = GetMetadataJSON("/metadata/appmodel/apps/" & workspaceId & "?requestDataType=7")
    in
        result;

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
    optional workspaceID as (
        type text meta [
            Documentation.FieldCaption = "Internal Workspace ID",
            Documentation.FieldDescription = "Internal Workspace ID",
            Documentation.SampleValues = {
                "123456"
            }
        ]
    )
) as table meta [
    Documentation.Name = "PBIMETADATARESTAPIGCC.GetDataflowExtendedSettings",
    Documentation.LongDescription = "Returns a list of details for dataflows in a specific workspace.",
    Documentation.Examples = {
        [
            Code = "=PBIMETADATARESTAPIGCC.GetDataflowExtendedSettings(""123456"")",
            Result = "Non-traditional API endpoint"
        ]
    }
];

GetDataflowExtendedSettings = (optional workspaceID as text) =>
    let
        result = GetMetadataJSON("/metadata/dataflows/" & workspaceID & "/extendedSettings")
    in
        result;        

/*****************
    Pipelines
*****************/
// GetALMPipelines
shared PBIMETADATARESTAPIGCC.GetALMPipelines = Value.ReplaceType(
    GetALMPipelines,
    GetALMPipelinesType
);

GetALMPipelinesType = type function (
) as table meta [
    Documentation.Name = "PBIMETADATARESTAPIGCC.GetALMPipelines",
    Documentation.LongDescription = "Returns metadata for deployment pipelines of current signed-in user",
    Documentation.Examples = {
        [
            Code = "=PBIMETADATARESTAPIGCC.GetALMPipelines()",
            Result = "Non-traditional API endpoint"
        ]
    }
];

GetALMPipelines = () =>
    let
        result = GetMetadataJSON("/metadata/almpipelines")
    in
        result; 


// GetALMPipelineOperation
shared PBIMETADATARESTAPIGCC.GetALMPipelineOperation = Value.ReplaceType(
    GetALMPipelineOperation,
    GetALMPipelineOperationType
);

GetALMPipelineOperationType = type function (
    optional pipelineID as (
        type text meta [
            Documentation.FieldCaption = "Internal Pipeline ID",
            Documentation.FieldDescription = "Internal Pipeline ID",
            Documentation.SampleValues = {
                "123"
            }
        ]
    ),
    optional operationID as (
        type text meta [
            Documentation.FieldCaption = "GUID of the operation in the deployment pipeline",
            Documentation.FieldDescription = "GUID of the operation in the deployment pipeline",
            Documentation.SampleValues = {
                "xxxxxxxx-8123-497d-9fc0-cefae4b8d11e"
            }
        ]
    )    
) as table meta [
    Documentation.Name = "PBIMETADATARESTAPIGCC.GetALMPipelineOperation",
    Documentation.LongDescription = "Returns a list of details for dataflows in a specific workspace.",
    Documentation.Examples = {
        [
            Code = "=PBIMETADATARESTAPIGCC.GetALMPipelineOperation(""123"",""xxxxxxxx-8123-497d-9fc0-cefae4b8d11e"")",
            Result = "Non-traditional API endpoint"
        ]
    }
];

GetALMPipelineOperation = (optional pipelineID as text, optional operationID) =>
    let
        result = GetMetadataJSON("/metadata/almpipelines/" & pipelineID & "/operations/" & operationID)
    in
        result; 

// GetALMPipelineStageDifferences
shared PBIMETADATARESTAPIGCC.GetALMPipelineStageDifferences = Value.ReplaceType(
    GetALMPipelineStageDifferences,
    GetALMPipelineStageDifferencesType
);

GetALMPipelineStageDifferencesType = type function (
    optional pipelineID as (
        type text meta [
            Documentation.FieldCaption = "Internal Pipeline ID",
            Documentation.FieldDescription = "Internal Pipeline ID",
            Documentation.SampleValues = {
                "123"
            }
        ]
    ),
    optional stageID as (
        type text meta [
            Documentation.FieldCaption = "Integer representing target workspace in pipeline (ex. 1)",
            Documentation.FieldDescription = "Integer representing target workspace in pipeline (ex. 1)",
            Documentation.SampleValues = {
                "1"
            }
        ]
    )    
) as table meta [
    Documentation.Name = "PBIMETADATARESTAPIGCC.GetALMPipelineStageDifferences",
    Documentation.LongDescription = "Returns a list comparing the source and target stage artifacts",
    Documentation.Examples = {
        [
            Code = "=PBIMETADATARESTAPIGCC.GetALMPipelineStageDifferences(""123"",""2"")",
            Result = "Non-traditional API endpoint"
        ]
    }
];

GetALMPipelineStageDifferences = (optional pipelineID as text, optional stageID as text) =>
    let
        result = GetMetadataJSON("/metadata/almpipelines/" & pipelineID & "/stages/" & stageID & "/diff")
    in
        result;                   

// DiscoverArtifacts 

shared PBIMETADATARESTAPIGCC.DiscoverArtifacts = Value.ReplaceType(
    DiscoverArtifacts,
    DiscoverArtifactsType
);

DiscoverArtifactsType = type function (    
    optional pipelineID as (
        type text meta [
            Documentation.FieldCaption = "Internal Pipeline ID",
            Documentation.FieldDescription = "Internal Pipeline ID",
            Documentation.SampleValues = {
                "123"
            }
        ]
    ),
    optional sourceObjectID as (
        type text meta [
            Documentation.FieldCaption = "GUID of the artifact in the source workspace",
            Documentation.FieldDescription = "GUID of the artifact in the source workspace",
            Documentation.SampleValues = {
                "xxxxxxxx-8123-497d-9fc0-cefae4b8d11e"
            }
        ]
    ),
    optional targetObjectID as (
        type text meta [
            Documentation.FieldCaption = "GUID of the artifact in the target workspace",
            Documentation.FieldDescription = "GUID of the artifact in the target workspace",
            Documentation.SampleValues = {
                "xxxxxxxx-1234-497d-9fc0-cefae4b8d11e"
            }
        ]
    ),
    optional typeName as (
        type text meta [
            Documentation.FieldCaption = "Type of artifact in Power BI",
            Documentation.FieldDescription = "Type of artifact in Power BI",
            Documentation.SampleValues = {
                "dataset"
            }
        ]
    )            
) as table meta [
    Documentation.Name = "PBIMETADATARESTAPIGCC.DiscoverArtifacts",
    Documentation.LongDescription = "Returns metadata for pipeline associated with the artifact.",
    Documentation.Examples = {
        [
            Code = "=PBIMETADATARESTAPIGCC.DiscoverArtifacts()",
            Result = "Non-traditional API endpoint"
        ]
    }
];

DiscoverArtifacts = (optional pipelineID as text, optional sourceObjectID as text, optional targetObjectID as text, optional typeName as text) =>
    let
        params = [RelativePath = "/metadata/almpipelines/" & pipelineID & "/DiscoverArtifacts",
                    Headers = [ #"Content-Type" = "application/json"],
                    Content = Text.ToBinary("{""artifacts"":[{""sourceObjectId"":""" & sourceObjectID & """,""targetObjectId"":""" & targetObjectID & """,""type"":"""& typeName &"""}]}")
                ],   

        result = PostMetadataJSON(params)     
    in
        result; 

/*****************
    Datasets
*****************/
// GetSharedDatasets 
shared PBIMETADATARESTAPIGCC.GetSharedDatasets = Value.ReplaceType(
    GetSharedDatasets,
    GetSharedDatasetsType
);

GetSharedDatasetsType = type function (
) as table meta [
    Documentation.Name = "PBIMETADATARESTAPIGCC.GetSharedDatasets",
    Documentation.LongDescription = "Returns metadata for datasets of current signed-in user",
    Documentation.Examples = {
        [
            Code = "=PBIMETADATARESTAPIGCC.GetSharedDatasets()",
            Result = "Non-traditional API endpoint"
        ]
    }
];

GetSharedDatasets = () =>
    let
        result = GetMetadataJSON("/metadata/gallery/SharedDatasets")
    in
        result; 

// GetOneDriveRefreshEnforcementPolicy
shared PBIMETADATARESTAPIGCC.GetOneDriveRefreshEnforcementPolicy = Value.ReplaceType(
    GetOneDriveRefreshEnforcementPolicy,
    GetOneDriveRefreshEnforcementPolicyType
);

GetOneDriveRefreshEnforcementPolicyType = type function (
    optional datasetID as (
        type text meta [
            Documentation.FieldCaption = "Internal Dataset ID",
            Documentation.FieldDescription = "Internal Dataset ID",
            Documentation.SampleValues = {
                "123456"
            }
        ]
    )
) as table meta [
    Documentation.Name = "PBIMETADATARESTAPIGCC.GetOneDriveRefreshEnforcementPolicy",
    Documentation.LongDescription = "Returns OneDrive dataset settings for specific dataset",
    Documentation.Examples = {
        [
            Code = "=PBIMETADATARESTAPIGCC.GetOneDriveRefreshEnforcementPolicy(""123456"")",
            Result = "Non-traditional API endpoint"
        ]
    }
];

GetOneDriveRefreshEnforcementPolicy = (optional datasetID as text) =>
    let
        result = GetMetadataJSON("metadata/oneDriveRefresh/" & datasetID & "/refreshEnforcementPolicy")
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

PostMetadataJSON = (optional params as record) =>
    let
        //Post
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