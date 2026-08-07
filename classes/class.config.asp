<%
'#######################################################################
'
'  InfoPower Importer PRO v3.0
'  File : class.config.asp
'  Description : Global Configuration
'
'#######################################################################

Option Explicit

'=======================================================================
' APPLICATION
'=======================================================================

Const APP_NAME               = "InfoPower Importer PRO"
Const APP_VERSION            = "3.0.0"
Const APP_BUILD              = "20260807"

'=======================================================================
' SOURCE WEBSITE
'=======================================================================

Const SOURCE_NAME            = "Genpower"

Const SOURCE_SITE            = "https://www.genpower.com.tr"

Const PRODUCT_LIST_URL       = _
"https://www.genpower.com.tr/tr/urunler/dizel-jeneratorleri"

Const USER_AGENT             = _
"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"

'=======================================================================
' IMPORT SETTINGS
'=======================================================================

Const IMPORT_PRODUCTS        = True
Const IMPORT_IMAGES          = True
Const IMPORT_PDF             = True
Const IMPORT_SPECS           = True
Const IMPORT_DESCRIPTION     = True
Const IMPORT_SEO             = True

Const UPDATE_EXISTING        = True
Const INSERT_NEW_PRODUCTS    = True
Const DELETE_MISSING         = False

'=======================================================================
' HTTP
'=======================================================================

Const HTTP_CONNECT_TIMEOUT   = 30000
Const HTTP_SEND_TIMEOUT      = 30000
Const HTTP_RECEIVE_TIMEOUT   = 60000
Const HTTP_RESPONSE_TIMEOUT  = 60000

'=======================================================================
' CACHE
'=======================================================================

Const CACHE_FOLDER           = "cache"
Const CACHE_INDEX_FILE       = "index.html"
Const CACHE_PRODUCT_FOLDER   = "products"

'=======================================================================
' IMAGE
'=======================================================================

Const IMAGE_FOLDER           = "/images/products/"
Const IMAGE_MAX_COUNT        = 30

'=======================================================================
' PDF
'=======================================================================

Const PDF_FOLDER             = "/downloads/"

'=======================================================================
' LOG
'=======================================================================

Const LOG_BROWSER            = True
Const LOG_DATABASE           = True

'=======================================================================
' DATABASE
'=======================================================================

Const DB_SCHEMA              = "hatmedya_Hknp-3bWc"

Const TBL_PRODUCTS           = "[" & DB_SCHEMA & "].[importer_products]"
Const TBL_IMAGES             = "[" & DB_SCHEMA & "].[importer_product_images]"
Const TBL_SPECS              = "[" & DB_SCHEMA & "].[importer_product_specs]"
Const TBL_FILES              = "[" & DB_SCHEMA & "].[importer_product_files]"
Const TBL_QUEUE              = "[" & DB_SCHEMA & "].[importer_queue]"
Const TBL_HISTORY            = "[" & DB_SCHEMA & "].[importer_history]"
Const TBL_LOGS               = "[" & DB_SCHEMA & "].[importer_logs]"
Const TBL_SETTINGS           = "[" & DB_SCHEMA & "].[importer_settings]"
Const TBL_BRANDS             = "[" & DB_SCHEMA & "].[importer_brands]"
Const TBL_CATEGORIES         = "[" & DB_SCHEMA & "].[importer_categories]"

'=======================================================================
' PRODUCT STATUS
'=======================================================================

Const STATUS_PASSIVE         = 0
Const STATUS_ACTIVE          = 1

'=======================================================================
' QUEUE STATUS
'=======================================================================

Const QUEUE_WAITING          = 0
Const QUEUE_RUNNING          = 1
Const QUEUE_COMPLETED        = 2
Const QUEUE_ERROR            = 3

'=======================================================================
' GLOBAL VARIABLES
'=======================================================================

Dim AppRoot
Dim CacheRoot
Dim CacheProductRoot
Dim ImageRoot
Dim PdfRoot

Dim ImportStartTime

Dim TotalProducts
Dim ImportedProducts
Dim UpdatedProducts
Dim ErrorProducts

Dim TotalImages
Dim TotalPDF

'=======================================================================
' INITIALIZE
'=======================================================================

Sub Config_Initialize()

    AppRoot = Server.MapPath("..")

    CacheRoot = AppRoot & "\" & CACHE_FOLDER

    CacheProductRoot = CacheRoot & "\" & CACHE_PRODUCT_FOLDER

    ImageRoot = Server.MapPath(IMAGE_FOLDER)

    PdfRoot = Server.MapPath(PDF_FOLDER)

    ImportStartTime = Now()

    TotalProducts = 0
    ImportedProducts = 0
    UpdatedProducts = 0
    ErrorProducts = 0

    TotalImages = 0
    TotalPDF = 0

End Sub

'#######################################################################
' END OF FILE
'#######################################################################
%>