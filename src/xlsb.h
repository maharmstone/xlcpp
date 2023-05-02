#pragma once

enum class xlsb_type {
    BrtRowHdr = 0,
    BrtCellBlank = 1,
    BrtCellRk = 2,
    BrtCellError = 3,
    BrtCellBool = 4,
    BrtCellReal = 5,
    BrtCellSt = 6,
    BrtCellIsst = 7,
    BrtFmlaString = 8,
    BrtFmlaNum = 9,
    BrtFmlaBool = 10,
    BrtFmlaError = 11,
    BrtSSTItem = 19,
    BrtPCDIMissing = 20,
    BrtPCDINumber = 21,
    BrtPCDIBoolean = 22,
    BrtPCDIError = 23,
    BrtPCDIString = 24,
    BrtPCDIDatetime = 25,
    BrtPCDIIndex = 26,
    BrtPCDIAMissing = 27,
    BrtPCDIANumber = 28,
    BrtPCDIABoolean = 29,
    BrtPCDIAError = 30,
    BrtPCDIAString = 31,
    BrtPCDIADatetime = 32,
    BrtPCRRecord = 33,
    BrtPCRRecordDt = 34,
    BrtFRTBegin = 35,
    BrtFRTEnd = 36,
    BrtACBegin = 37,
    BrtACEnd = 38,
    BrtName = 39,
    BrtIndexRowBlock = 40,
    BrtIndexBlock = 42,
    BrtFont = 43,
    BrtFmt = 44,
    BrtFill = 45,
    BrtBorder = 46,
    BrtXF = 47,
    BrtStyle = 48,
    BrtCellMeta = 49,
    BrtValueMeta = 50,
    BrtMdb = 51,
    BrtBeginFmd = 52,
    BrtEndFmd = 53,
    BrtBeginMdx = 54,
    BrtEndMdx = 55,
    BrtBeginMdxTuple = 56,
    BrtEndMdxTuple = 57,
    BrtMdxMbrIstr = 58,
    BrtStr = 59,
    BrtColInfo = 60,
    BrtCellRString = 62,
    BrtDVal = 64,
    BrtSxvcellNum = 65,
    BrtSxvcellStr = 66,
    BrtSxvcellBool = 67,
    BrtSxvcellErr = 68,
    BrtSxvcellDate = 69,
    BrtSxvcellNil = 70,
    BrtFileVersion = 128,
    BrtBeginSheet = 129,
    BrtEndSheet = 130,
    BrtBeginBook = 131,
    BrtEndBook = 132,
    BrtBeginWsViews = 133,
    BrtEndWsViews = 134,
    BrtBeginBookViews = 135,
    BrtEndBookViews = 136,
    BrtBeginWsView = 137,
    BrtEndWsView = 138,
    BrtBeginCsViews = 139,
    BrtEndCsViews = 140,
    BrtBeginCsView = 141,
    BrtEndCsView = 142,
    BrtBeginBundleShs = 143,
    BrtEndBundleShs = 144,
    BrtBeginSheetData = 145,
    BrtEndSheetData = 146,
    BrtWsProp = 147,
    BrtWsDim = 148,
    BrtPane = 151,
    BrtSel = 152,
    BrtWbProp = 153,
    BrtWbFactoid = 154,
    BrtFileRecover = 155,
    BrtBundleSh = 156,
    BrtCalcProp = 157,
    BrtBookView = 158,
    BrtBeginSst = 159,
    BrtEndSst = 160,
    BrtBeginAFilter = 161,
    BrtEndAFilter = 162,
    BrtBeginFilterColumn = 163,
    BrtEndFilterColumn = 164,
    BrtBeginFilters = 165,
    BrtEndFilters = 166,
    BrtFilter = 167,
    BrtColorFilter = 168,
    BrtIconFilter = 169,
    BrtTop10Filter = 170,
    BrtDynamicFilter = 171,
    BrtBeginCustomFilters = 172,
    BrtEndCustomFilters = 173,
    BrtCustomFilter = 174,
    BrtAFilterDateGroupItem = 175,
    BrtMergeCell = 176,
    BrtBeginMergeCells = 177,
    BrtEndMergeCells = 178,
    BrtBeginPivotCacheDef = 179,
    BrtEndPivotCacheDef = 180,
    BrtBeginPCDFields = 181,
    BrtEndPCDFields = 182,
    BrtBeginPCDField = 183,
    BrtEndPCDField = 184,
    BrtBeginPCDSource = 185,
    BrtEndPCDSource = 186,
    BrtBeginPCDSRange = 187,
    BrtEndPCDSRange = 188,
    BrtBeginPCDFAtbl = 189,
    BrtEndPCDFAtbl = 190,
    BrtBeginPCDIRun = 191,
    BrtEndPCDIRun = 192,
    BrtBeginPivotCacheRecords = 193,
    BrtEndPivotCacheRecords = 194,
    BrtBeginPCDHierarchies = 195,
    BrtEndPCDHierarchies = 196,
    BrtBeginPCDHierarchy = 197,
    BrtEndPCDHierarchy = 198,
    BrtBeginPCDHFieldsUsage = 199,
    BrtEndPCDHFieldsUsage = 200,
    BrtBeginExtConnection = 201,
    BrtEndExtConnection = 202,
    BrtBeginECDbProps = 203,
    BrtEndECDbProps = 204,
    BrtBeginECOlapProps = 205,
    BrtEndECOlapProps = 206,
    BrtBeginPCDSConsol = 207,
    BrtEndPCDSConsol = 208,
    BrtBeginPCDSCPages = 209,
    BrtEndPCDSCPages = 210,
    BrtBeginPCDSCPage = 211,
    BrtEndPCDSCPage = 212,
    BrtBeginPCDSCPItem = 213,
    BrtEndPCDSCPItem = 214,
    BrtBeginPCDSCSets = 215,
    BrtEndPCDSCSets = 216,
    BrtBeginPCDSCSet = 217,
    BrtEndPCDSCSet = 218,
    BrtBeginPCDFGroup = 219,
    BrtEndPCDFGroup = 220,
    BrtBeginPCDFGItems = 221,
    BrtEndPCDFGItems = 222,
    BrtBeginPCDFGRange = 223,
    BrtEndPCDFGRange = 224,
    BrtBeginPCDFGDiscrete = 225,
    BrtEndPCDFGDiscrete = 226,
    BrtBeginPCDSDTupleCache = 227,
    BrtEndPCDSDTupleCache = 228,
    BrtBeginPCDSDTCEntries = 229,
    BrtEndPCDSDTCEntries = 230,
    BrtBeginPCDSDTCEMembers = 231,
    BrtEndPCDSDTCEMembers = 232,
    BrtBeginPCDSDTCEMember = 233,
    BrtEndPCDSDTCEMember = 234,
    BrtBeginPCDSDTCQueries = 235,
    BrtEndPCDSDTCQueries = 236,
    BrtBeginPCDSDTCQuery = 237,
    BrtEndPCDSDTCQuery = 238,
    BrtBeginPCDSDTCSets = 239,
    BrtEndPCDSDTCSets = 240,
    BrtBeginPCDSDTCSet = 241,
    BrtEndPCDSDTCSet = 242,
    BrtBeginPCDCalcItems = 243,
    BrtEndPCDCalcItems = 244,
    BrtBeginPCDCalcItem = 245,
    BrtEndPCDCalcItem = 246,
    BrtBeginPRule = 247,
    BrtEndPRule = 248,
    BrtBeginPRFilters = 249,
    BrtEndPRFilters = 250,
    BrtBeginPRFilter = 251,
    BrtEndPRFilter = 252,
    BrtBeginPNames = 253,
    BrtEndPNames = 254,
    BrtBeginPName = 255,
    BrtEndPName = 256,
    BrtBeginPNPairs = 257,
    BrtEndPNPairs = 258,
    BrtBeginPNPair = 259,
    BrtEndPNPair = 260,
    BrtBeginECWebProps = 261,
    BrtEndECWebProps = 262,
    BrtBeginEcWpTables = 263,
    BrtEndECWPTables = 264,
    BrtBeginECParams = 265,
    BrtEndECParams = 266,
    BrtBeginECParam = 267,
    BrtEndECParam = 268,
    BrtBeginPCDKPIs = 269,
    BrtEndPCDKPIs = 270,
    BrtBeginPCDKPI = 271,
    BrtEndPCDKPI = 272,
    BrtBeginDims = 273,
    BrtEndDims = 274,
    BrtBeginDim = 275,
    BrtEndDim = 276,
    BrtIndexPartEnd = 277,
    BrtBeginStyleSheet = 278,
    BrtEndStyleSheet = 279,
    BrtBeginSXView = 280,
    BrtEndSXVI = 281,
    BrtBeginSXVI = 282,
    BrtBeginSXVIs = 283,
    BrtEndSXVIs = 284,
    BrtBeginSXVD = 285,
    BrtEndSXVD = 286,
    BrtBeginSXVDs = 287,
    BrtEndSXVDs = 288,
    BrtBeginSXPI = 289,
    BrtEndSXPI = 290,
    BrtBeginSXPIs = 291,
    BrtEndSXPIs = 292,
    BrtBeginSXDI = 293,
    BrtEndSXDI = 294,
    BrtBeginSXDIs = 295,
    BrtEndSXDIs = 296,
    BrtBeginSXLI = 297,
    BrtEndSXLI = 298,
    BrtBeginSXLIRws = 299,
    BrtEndSXLIRws = 300,
    BrtBeginSXLICols = 301,
    BrtEndSXLICols = 302,
    BrtBeginSXFormat = 303,
    BrtEndSXFormat = 304,
    BrtBeginSXFormats = 305,
    BrtEndSxFormats = 306,
    BrtBeginSxSelect = 307,
    BrtEndSxSelect = 308,
    BrtBeginISXVDRws = 309,
    BrtEndISXVDRws = 310,
    BrtBeginISXVDCols = 311,
    BrtEndISXVDCols = 312,
    BrtEndSXLocation = 313,
    BrtBeginSXLocation = 314,
    BrtEndSXView = 315,
    BrtBeginSXTHs = 316,
    BrtEndSXTHs = 317,
    BrtBeginSXTH = 318,
    BrtEndSXTH = 319,
    BrtBeginISXTHRws = 320,
    BrtEndISXTHRws = 321,
    BrtBeginISXTHCols = 322,
    BrtEndISXTHCols = 323,
    BrtBeginSXTDMPS = 324,
    BrtEndSXTDMPs = 325,
    BrtBeginSXTDMP = 326,
    BrtEndSXTDMP = 327,
    BrtBeginSXTHItems = 328,
    BrtEndSXTHItems = 329,
    BrtBeginSXTHItem = 330,
    BrtEndSXTHItem = 331,
    BrtBeginMetadata = 332,
    BrtEndMetadata = 333,
    BrtBeginEsmdtinfo = 334,
    BrtMdtinfo = 335,
    BrtEndEsmdtinfo = 336,
    BrtBeginEsmdb = 337,
    BrtEndEsmdb = 338,
    BrtBeginEsfmd = 339,
    BrtEndEsfmd = 340,
    BrtBeginSingleCells = 341,
    BrtEndSingleCells = 342,
    BrtBeginList = 343,
    BrtEndList = 344,
    BrtBeginListCols = 345,
    BrtEndListCols = 346,
    BrtBeginListCol = 347,
    BrtEndListCol = 348,
    BrtBeginListXmlCPr = 349,
    BrtEndListXmlCPr = 350,
    BrtListCCFmla = 351,
    BrtListTrFmla = 352,
    BrtBeginExternals = 353,
    BrtEndExternals = 354,
    BrtSupBookSrc = 355,
    BrtSupSelf = 357,
    BrtSupSame = 358,
    BrtSupTabs = 359,
    BrtBeginSupBook = 360,
    BrtPlaceholderName = 361,
    BrtExternSheet = 362,
    BrtExternTableStart = 363,
    BrtExternTableEnd = 364,
    BrtExternRowHdr = 366,
    BrtExternCellBlank = 367,
    BrtExternCellReal = 368,
    BrtExternCellBool = 369,
    BrtExternCellError = 370,
    BrtExternCellString = 371,
    BrtBeginEsmdx = 372,
    BrtEndEsmdx = 373,
    BrtBeginMdxSet = 374,
    BrtEndMdxSet = 375,
    BrtBeginMdxMbrProp = 376,
    BrtEndMdxMbrProp = 377,
    BrtBeginMdxKPI = 378,
    BrtEndMdxKPI = 379,
    BrtBeginEsstr = 380,
    BrtEndEsstr = 381,
    BrtBeginPRFItem = 382,
    BrtEndPRFItem = 383,
    BrtBeginPivotCacheIDs = 384,
    BrtEndPivotCacheIDs = 385,
    BrtBeginPivotCacheID = 386,
    BrtEndPivotCacheID = 387,
    BrtBeginISXVIs = 388,
    BrtEndISXVIs = 389,
    BrtBeginColInfos = 390,
    BrtEndColInfos = 391,
    BrtBeginRwBrk = 392,
    BrtEndRwBrk = 393,
    BrtBeginColBrk = 394,
    BrtEndColBrk = 395,
    BrtBrk = 396,
    BrtUserBookView = 397,
    BrtInfo = 398,
    BrtCUsr = 399,
    BrtUsr = 400,
    BrtBeginUsers = 401,
    BrtEOF = 403,
    BrtUCR = 404,
    BrtRRInsDel = 405,
    BrtRREndInsDel = 406,
    BrtRRMove = 407,
    BrtRREndMove = 408,
    BrtRRChgCell = 409,
    BrtRREndChgCell = 410,
    BrtRRHeader = 411,
    BrtRRUserView = 412,
    BrtRRRenSheet = 413,
    BrtRRInsertSh = 414,
    BrtRRDefName = 415,
    BrtRRNote = 416,
    BrtRRConflict = 417,
    BrtRRTQSIF = 418,
    BrtRRFormat = 419,
    BrtRREndFormat = 420,
    BrtRRAutoFmt = 421,
    BrtBeginUserShViews = 422,
    BrtBeginUserShView = 423,
    BrtEndUserShView = 424,
    BrtEndUserShViews = 425,
    BrtArrFmla = 426,
    BrtShrFmla = 427,
    BrtTable = 428,
    BrtBeginExtConnections = 429,
    BrtEndExtConnections = 430,
    BrtBeginPCDCalcMems = 431,
    BrtEndPCDCalcMems = 432,
    BrtBeginPCDCalcMem = 433,
    BrtEndPCDCalcMem = 434,
    BrtBeginPCDHGLevels = 435,
    BrtEndPCDHGLevels = 436,
    BrtBeginPCDHGLevel = 437,
    BrtEndPCDHGLevel = 438,
    BrtBeginPCDHGLGroups = 439,
    BrtEndPCDHGLGroups = 440,
    BrtBeginPCDHGLGroup = 441,
    BrtEndPCDHGLGroup = 442,
    BrtBeginPCDHGLGMembers = 443,
    BrtEndPCDHGLGMembers = 444,
    BrtBeginPCDHGLGMember = 445,
    BrtEndPCDHGLGMember = 446,
    BrtBeginQSI = 447,
    BrtEndQSI = 448,
    BrtBeginQSIR = 449,
    BrtEndQSIR = 450,
    BrtBeginDeletedNames = 451,
    BrtEndDeletedNames = 452,
    BrtBeginDeletedName = 453,
    BrtEndDeletedName = 454,
    BrtBeginQSIFs = 455,
    BrtEndQSIFs = 456,
    BrtBeginQSIF = 457,
    BrtEndQSIF = 458,
    BrtBeginAutoSortScope = 459,
    BrtEndAutoSortScope = 460,
    BrtBeginConditionalFormatting = 461,
    BrtEndConditionalFormatting = 462,
    BrtBeginCFRule = 463,
    BrtEndCFRule = 464,
    BrtBeginIconSet = 465,
    BrtEndIconSet = 466,
    BrtBeginDatabar = 467,
    BrtEndDatabar = 468,
    BrtBeginColorScale = 469,
    BrtEndColorScale = 470,
    BrtCFVO = 471,
    BrtExternValueMeta = 472,
    BrtBeginColorPalette = 473,
    BrtEndColorPalette = 474,
    BrtIndexedColor = 475,
    BrtMargins = 476,
    BrtPrintOptions = 477,
    BrtPageSetup = 478,
    BrtBeginHeaderFooter = 479,
    BrtEndHeaderFooter = 480,
    BrtBeginSXCrtFormat = 481,
    BrtEndSXCrtFormat = 482,
    BrtBeginSXCrtFormats = 483,
    BrtEndSXCrtFormats = 484,
    BrtWsFmtInfo = 485,
    BrtBeginMgs = 486,
    BrtEndMGs = 487,
    BrtBeginMGMaps = 488,
    BrtEndMGMaps = 489,
    BrtBeginMG = 490,
    BrtEndMG = 491,
    BrtBeginMap = 492,
    BrtEndMap = 493,
    BrtHLink = 494,
    BrtBeginDCon = 495,
    BrtEndDCon = 496,
    BrtBeginDRefs = 497,
    BrtEndDRefs = 498,
    BrtDRef = 499,
    BrtBeginScenMan = 500,
    BrtEndScenMan = 501,
    BrtBeginSct = 502,
    BrtEndSct = 503,
    BrtSlc = 504,
    BrtBeginDXFs = 505,
    BrtEndDXFs = 506,
    BrtDXF = 507,
    BrtBeginTableStyles = 508,
    BrtEndTableStyles = 509,
    BrtBeginTableStyle = 510,
    BrtEndTableStyle = 511,
    BrtTableStyleElement = 512,
    BrtTableStyleClient = 513,
    BrtBeginVolDeps = 514,
    BrtEndVolDeps = 515,
    BrtBeginVolType = 516,
    BrtEndVolType = 517,
    BrtBeginVolMain = 518,
    BrtEndVolMain = 519,
    BrtBeginVolTopic = 520,
    BrtEndVolTopic = 521,
    BrtVolSubtopic = 522,
    BrtVolRef = 523,
    BrtVolNum = 524,
    BrtVolErr = 525,
    BrtVolStr = 526,
    BrtVolBool = 527,
    BrtBeginSortState = 530,
    BrtEndSortState = 531,
    BrtBeginSortCond = 532,
    BrtEndSortCond = 533,
    BrtBookProtection = 534,
    BrtSheetProtection = 535,
    BrtRangeProtection = 536,
    BrtPhoneticInfo = 537,
    BrtBeginECTxtWiz = 538,
    BrtEndECTxtWiz = 539,
    BrtBeginECTWFldInfoLst = 540,
    BrtEndECTWFldInfoLst = 541,
    BrtBeginECTwFldInfo = 542,
    BrtFileSharing = 548,
    BrtOleSize = 549,
    BrtDrawing = 550,
    BrtLegacyDrawing = 551,
    BrtLegacyDrawingHF = 552,
    BrtWebOpt = 553,
    BrtBeginWebPubItems = 554,
    BrtEndWebPubItems = 555,
    BrtBeginWebPubItem = 556,
    BrtEndWebPubItem = 557,
    BrtBeginSXCondFmt = 558,
    BrtEndSXCondFmt = 559,
    BrtBeginSXCondFmts = 560,
    BrtEndSXCondFmts = 561,
    BrtBkHim = 562,
    BrtColor = 564,
    BrtBeginIndexedColors = 565,
    BrtEndIndexedColors = 566,
    BrtBeginMRUColors = 569,
    BrtEndMRUColors = 570,
    BrtMRUColor = 572,
    BrtBeginDVals = 573,
    BrtEndDVals = 574,
    BrtSupNameStart = 577,
    BrtSupNameValueStart = 578,
    BrtSupNameValueEnd = 579,
    BrtSupNameNum = 580,
    BrtSupNameErr = 581,
    BrtSupNameSt = 582,
    BrtSupNameNil = 583,
    BrtSupNameBool = 584,
    BrtSupNameFmla = 585,
    BrtSupNameBits = 586,
    BrtSupNameEnd = 587,
    BrtEndSupBook = 588,
    BrtCellSmartTagProperty = 589,
    BrtBeginCellSmartTag = 590,
    BrtEndCellSmartTag = 591,
    BrtBeginCellSmartTags = 592,
    BrtEndCellSmartTags = 593,
    BrtBeginSmartTags = 594,
    BrtEndSmartTags = 595,
    BrtSmartTagType = 596,
    BrtBeginSmartTagTypes = 597,
    BrtEndSmartTagTypes = 598,
    BrtBeginSXFilters = 599,
    BrtEndSXFilters = 600,
    BrtBeginSXFILTER = 601,
    BrtEndSXFilter = 602,
    BrtBeginFills = 603,
    BrtEndFills = 604,
    BrtBeginCellWatches = 605,
    BrtEndCellWatches = 606,
    BrtCellWatch = 607,
    BrtBeginCRErrs = 608,
    BrtEndCRErrs = 609,
    BrtCrashRecErr = 610,
    BrtBeginFonts = 611,
    BrtEndFonts = 612,
    BrtBeginBorders = 613,
    BrtEndBorders = 614,
    BrtBeginFmts = 615,
    BrtEndFmts = 616,
    BrtBeginCellXFs = 617,
    BrtEndCellXFs = 618,
    BrtBeginStyles = 619,
    BrtEndStyles = 620,
    BrtBigName = 625,
    BrtBeginCellStyleXFs = 626,
    BrtEndCellStyleXFs = 627,
    BrtBeginComments = 628,
    BrtEndComments = 629,
    BrtBeginCommentAuthors = 630,
    BrtEndCommentAuthors = 631,
    BrtCommentAuthor = 632,
    BrtBeginCommentList = 633,
    BrtEndCommentList = 634,
    BrtBeginComment = 635,
    BrtEndComment = 636,
    BrtCommentText = 637,
    BrtBeginOleObjects = 638,
    BrtOleObject = 639,
    BrtEndOleObjects = 640,
    BrtBeginSxrules = 641,
    BrtEndSxRules = 642,
    BrtBeginActiveXControls = 643,
    BrtActiveX = 644,
    BrtEndActiveXControls = 645,
    BrtBeginPCDSDTCEMembersSortBy = 646,
    BrtBeginCellIgnoreECs = 648,
    BrtCellIgnoreEC = 649,
    BrtEndCellIgnoreECs = 650,
    BrtCsProp = 651,
    BrtCsPageSetup = 652,
    BrtBeginUserCsViews = 653,
    BrtEndUserCsViews = 654,
    BrtBeginUserCsView = 655,
    BrtEndUserCsView = 656,
    BrtBeginPcdSFCIEntries = 657,
    BrtEndPCDSFCIEntries = 658,
    BrtPCDSFCIEntry = 659,
    BrtBeginListParts = 660,
    BrtListPart = 661,
    BrtEndListParts = 662,
    BrtSheetCalcProp = 663,
    BrtBeginFnGroup = 664,
    BrtFnGroup = 665,
    BrtEndFnGroup = 666,
    BrtSupAddin = 667,
    BrtSXTDMPOrder = 668,
    BrtCsProtection = 669,
    BrtBeginWsSortMap = 671,
    BrtEndWsSortMap = 672,
    BrtBeginRRSort = 673,
    BrtEndRRSort = 674,
    BrtRRSortItem = 675,
    BrtFileSharingIso = 676,
    BrtBookProtectionIso = 677,
    BrtSheetProtectionIso = 678,
    BrtCsProtectionIso = 679,
    BrtRangeProtectionIso = 680,
    BrtDValList = 681,
    BrtRwDescent = 1024,
    BrtKnownFonts = 1025,
    BrtBeginSXTupleSet = 1026,
    BrtEndSXTupleSet = 1027,
    BrtBeginSXTupleSetHeader = 1028,
    BrtEndSXTupleSetHeader = 1029,
    BrtSXTupleSetHeaderItem = 1030,
    BrtBeginSXTupleSetData = 1031,
    BrtEndSXTupleSetData = 1032,
    BrtBeginSXTupleSetRow = 1033,
    BrtEndSXTupleSetRow = 1034,
    BrtSXTupleSetRowItem = 1035,
    BrtNameExt = 1036,
    BrtPCDH14 = 1037,
    BrtBeginPCDCalcMem14 = 1038,
    BrtEndPCDCalcMem14 = 1039,
    BrtSXTH14 = 1040,
    BrtBeginSparklineGroup = 1041,
    BrtEndSparklineGroup = 1042,
    BrtSparkline = 1043,
    BrtSXDI14 = 1044,
    BrtWsFmtInfoEx14 = 1045,
    BrtBeginConditionalFormatting14 = 1046,
    BrtEndConditionalFormatting14 = 1047,
    BrtBeginCFRule14 = 1048,
    BrtEndCFRule14 = 1049,
    BrtCFVO14 = 1050,
    BrtBeginDatabar14 = 1051,
    BrtBeginIconSet14 = 1052,
    BrtDVal14 = 1053,
    BrtBeginDVals14 = 1054,
    BrtColor14 = 1055,
    BrtBeginSparklines = 1056,
    BrtEndSparklines = 1057,
    BrtBeginSparklineGroups = 1058,
    BrtEndSparklineGroups = 1059,
    BrtSXVD14 = 1061,
    BrtBeginSxView14 = 1062,
    BrtEndSxView14 = 1063,
    BrtBeginSXView16 = 1064,
    BrtEndSXView16 = 1065,
    BrtBeginPCD14 = 1066,
    BrtEndPCD14 = 1067,
    BrtBeginExtConn14 = 1068,
    BrtEndExtConn14 = 1069,
    BrtBeginSlicerCacheIDs = 1070,
    BrtEndSlicerCacheIDs = 1071,
    BrtBeginSlicerCacheID = 1072,
    BrtEndSlicerCacheID = 1073,
    BrtBeginSlicerCache = 1075,
    BrtEndSlicerCache = 1076,
    BrtBeginSlicerCacheDef = 1077,
    BrtEndSlicerCacheDef = 1078,
    BrtBeginSlicersEx = 1079,
    BrtEndSlicersEx = 1080,
    BrtBeginSlicerEx = 1081,
    BrtEndSlicerEx = 1082,
    BrtBeginSlicer = 1083,
    BrtEndSlicer = 1084,
    BrtSlicerCachePivotTables = 1085,
    BrtBeginSlicerCacheOlapImpl = 1086,
    BrtEndSlicerCacheOlapImpl = 1087,
    BrtBeginSlicerCacheLevelsData = 1088,
    BrtEndSlicerCacheLevelsData = 1089,
    BrtBeginSlicerCacheLevelData = 1090,
    BrtEndSlicerCacheLevelData = 1091,
    BrtBeginSlicerCacheSiRanges = 1092,
    BrtEndSlicerCacheSiRanges = 1093,
    BrtBeginSlicerCacheSiRange = 1094,
    BrtEndSlicerCacheSiRange = 1095,
    BrtSlicerCacheOlapItem = 1096,
    BrtBeginSlicerCacheSelections = 1097,
    BrtSlicerCacheSelection = 1098,
    BrtEndSlicerCacheSelections = 1099,
    BrtBeginSlicerCacheNative = 1100,
    BrtEndSlicerCacheNative = 1101,
    BrtSlicerCacheNativeItem = 1102,
    BrtRangeProtection14 = 1103,
    BrtRangeProtectionIso14 = 1104,
    BrtCellIgnoreEC14 = 1105,
    BrtList14 = 1111,
    BrtCFIcon = 1112,
    BrtBeginSlicerCachesPivotCacheIDs = 1113,
    BrtEndSlicerCachesPivotCacheIDs = 1114,
    BrtBeginSlicers = 1115,
    BrtEndSlicers = 1116,
    BrtWbProp14 = 1117,
    BrtBeginSXEdit = 1118,
    BrtEndSXEdit = 1119,
    BrtBeginSXEdits = 1120,
    BrtEndSXEdits = 1121,
    BrtBeginSXChange = 1122,
    BrtEndSXChange = 1123,
    BrtBeginSXChanges = 1124,
    BrtEndSXChanges = 1125,
    BrtSXTupleItems = 1126,
    BrtBeginSlicerStyle = 1128,
    BrtEndSlicerStyle = 1129,
    BrtSlicerStyleElement = 1130,
    BrtBeginStyleSheetExt14 = 1131,
    BrtEndStyleSheetExt14 = 1132,
    BrtBeginSlicerCachesPivotCacheID = 1133,
    BrtEndSlicerCachesPivotCacheID = 1134,
    BrtBeginConditionalFormattings = 1135,
    BrtEndConditionalFormattings = 1136,
    BrtBeginPCDCalcMemExt = 1137,
    BrtEndPCDCalcMemExt = 1138,
    BrtBeginPCDCalcMemsExt = 1139,
    BrtEndPCDCalcMemsExt = 1140,
    BrtPCDField14 = 1141,
    BrtBeginSlicerStyles = 1142,
    BrtEndSlicerStyles = 1143,
    BrtBeginSlicerStyleElements = 1144,
    BrtEndSlicerStyleElements = 1145,
    BrtCFRuleExt = 1146,
    BrtBeginSXCondFmt14 = 1147,
    BrtEndSXCondFmt14 = 1148,
    BrtBeginSXCondFmts14 = 1149,
    BrtEndSXCondFmts14 = 1150,
    BrtBeginSortCond14 = 1152,
    BrtEndSortCond14 = 1153,
    BrtEndDVals14 = 1154,
    BrtEndIconSet14 = 1155,
    BrtEndDatabar14 = 1156,
    BrtBeginColorScale14 = 1157,
    BrtEndColorScale14 = 1158,
    BrtBeginSxrules14 = 1159,
    BrtEndSxrules14 = 1160,
    BrtBeginPRule14 = 1161,
    BrtEndPRule14 = 1162,
    BrtBeginPRFilters14 = 1163,
    BrtEndPRFilters14 = 1164,
    BrtBeginPRFilter14 = 1165,
    BrtEndPRFilter14 = 1166,
    BrtBeginPRFItem14 = 1167,
    BrtEndPRFItem14 = 1168,
    BrtBeginCellIgnoreECs14 = 1169,
    BrtEndCellIgnoreECs14 = 1170,
    BrtDxf14 = 1171,
    BrtBeginDxF14s = 1172,
    BrtEndDxf14s = 1173,
    BrtFilter14 = 1177,
    BrtBeginCustomFilters14 = 1178,
    BrtCustomFilter14 = 1180,
    BrtIconFilter14 = 1181,
    BrtPivotCacheConnectionName = 1182,
    BrtBeginDecoupledPivotCacheIDs = 2048,
    BrtEndDecoupledPivotCacheIDs = 2049,
    BrtDecoupledPivotCacheID = 2050,
    BrtBeginPivotTableRefs = 2051,
    BrtEndPivotTableRefs = 2052,
    BrtPivotTableRef = 2053,
    BrtSlicerCacheBookPivotTables = 2054,
    BrtBeginSxvcells = 2055,
    BrtEndSxvcells = 2056,
    BrtBeginSxRow = 2057,
    BrtEndSxRow = 2058,
    BrtPcdCalcMem15 = 2060,
    BrtQsi15 = 2067,
    BrtBeginWebExtensions = 2068,
    BrtEndWebExtensions = 2069,
    BrtWebExtension = 2070,
    BrtAbsPath15 = 2071,
    BrtBeginPivotTableUISettings = 2072,
    BrtEndPivotTableUISettings = 2073,
    BrtTableSlicerCacheIDs = 2075,
    BrtTableSlicerCacheID = 2076,
    BrtBeginTableSlicerCache = 2077,
    BrtEndTableSlicerCache = 2078,
    BrtSxFilter15 = 2079,
    BrtBeginTimelineCachePivotCacheIDs = 2080,
    BrtEndTimelineCachePivotCacheIDs = 2081,
    BrtTimelineCachePivotCacheID = 2082,
    BrtBeginTimelineCacheIDs = 2083,
    BrtEndTimelineCacheIDs = 2084,
    BrtBeginTimelineCacheID = 2085,
    BrtEndTimelineCacheID = 2086,
    BrtBeginTimelinesEx = 2087,
    BrtEndTimelinesEx = 2088,
    BrtBeginTimelineEx = 2089,
    BrtEndTimelineEx = 2090,
    BrtWorkBookPr15 = 2091,
    BrtPCDH15 = 2092,
    BrtBeginTimelineStyle = 2093,
    BrtEndTimelineStyle = 2094,
    BrtTimelineStyleElement = 2095,
    BrtBeginTimelineStylesheetExt15 = 2096,
    BrtEndTimelineStylesheetExt15 = 2097,
    BrtBeginTimelineStyles = 2098,
    BrtEndTimelineStyles = 2099,
    BrtBeginTimelineStyleElements = 2100,
    BrtEndTimelineStyleElements = 2101,
    BrtDxf15 = 2102,
    BrtBeginDxfs15 = 2103,
    BrtEndDXFs15 = 2104,
    BrtSlicerCacheHideItemsWithNoData = 2105,
    BrtBeginItemUniqueNames = 2106,
    BrtEndItemUniqueNames = 2107,
    BrtItemUniqueName = 2108,
    BrtBeginExtConn15 = 2109,
    BrtEndExtConn15 = 2110,
    BrtBeginOledbPr15 = 2111,
    BrtEndOledbPr15 = 2112,
    BrtBeginDataFeedPr15 = 2113,
    BrtEndDataFeedPr15 = 2114,
    BrtTextPr15 = 2115,
    BrtRangePr15 = 2116,
    BrtDbCommand15 = 2117,
    BrtBeginDbTables15 = 2118,
    BrtEndDbTables15 = 2119,
    BrtDbTable15 = 2120,
    BrtBeginDataModel = 2121,
    BrtEndDataModel = 2122,
    BrtBeginModelTables = 2123,
    BrtEndModelTables = 2124,
    BrtModelTable = 2125,
    BrtBeginModelRelationships = 2126,
    BrtEndModelRelationships = 2127,
    BrtModelRelationship = 2128,
    BrtBeginECTxtWiz15 = 2129,
    BrtEndECTxtWiz15 = 2130,
    BrtBeginECTWFldInfoLst15 = 2131,
    BrtEndECTWFldInfoLst15 = 2132,
    BrtBeginECTWFldInfo15 = 2133,
    BrtFieldListActiveItem = 2134,
    BrtPivotCacheIdVersion = 2135,
    BrtSXDI15 = 2136,
    brtBeginModelTimeGroupings = 2137,
    brtEndModelTimeGroupings = 2138,
    brtBeginModelTimeGrouping = 2139,
    brtEndModelTimeGrouping = 2140,
    brtModelTimeGroupingCalcCol = 2141,
    brtRevisionPtr = 3073,
    BrtBeginDynamicArrayPr = 4096,
    BrtEndDynamicArrayPr = 4097,
    BrtBeginRichValueBlock = 5002,
    BrtEndRichValueBlock = 5003,
    BrtBeginRichFilters = 5081,
    BrtEndRichFilters = 5082,
    BrtRichFilter = 5083,
    BrtBeginRichFilterColumn = 5084,
    BrtEndRichFilterColumn = 5085,
    BrtBeginCustomRichFilters = 5086,
    BrtEndCustomRichFilters = 5087,
    BRTCustomRichFilter = 5088,
    BrtTop10RichFilter = 5089,
    BrtDynamicRichFilter = 5090,
    BrtBeginRichSortCondition = 5092,
    BrtEndRichSortCondition = 5093,
    BrtRichFilterDateGroupItem = 5094,
    BrtBeginCalcFeatures = 5095,
    BrtEndCalcFeatures = 5096,
    BrtCalcFeature = 5097,
    BrtExternalLinksPr = 5099
};

struct brt_bundle_sh {
    uint32_t hsState;
    uint32_t iTabID;
};

#pragma pack(push,1)

struct brt_row_hdr {
    uint32_t rw;
    uint32_t ixfe;
    uint16_t miyRw;
    // FIXME - make sure bitfields are in correct order
    uint16_t fExtraAsc : 1;
    uint16_t fExtraDsc : 1;
    uint16_t reserved1 : 6;
    uint16_t iOutLevel : 3;
    uint16_t fCollapsed : 1;
    uint16_t fDyZero : 1;
    uint16_t fUnsynced : 1;
    uint16_t fGhostDirty : 1;
    uint16_t fReserved : 1;
    uint8_t fPhShow : 1;
    uint8_t reserved2 : 7;
    uint32_t ccolspan;
};

struct xlsb_cell {
    uint32_t column;
    uint32_t iStyleRef : 24;
    uint32_t fPhShow : 1;
    uint32_t reserved : 7;
};

struct rk_number {
    uint32_t fx100 : 1;
    uint32_t fInt : 1;
    uint32_t num : 30;
};

#pragma pack(pop)

struct brt_cell_rk {
    xlsb_cell cell;
    rk_number value;
};

struct brt_cell_bool {
    xlsb_cell cell;
    uint8_t fBool;
};

struct brt_cell_real {
    xlsb_cell cell;
    double xnum;
};

struct brt_cell_st {
    xlsb_cell cell;
    uint32_t len;
    char16_t str[0];
};

#pragma pack(push,1)

struct rich_str {
    uint8_t fRichStr : 1;
    uint8_t fExtStr : 1;
    uint8_t unused1 : 6;
    uint32_t len;
    char16_t str[0];
};

#pragma pack(pop)

struct brt_sst_item {
    rich_str richStr;
};

struct brt_cell_rstring {
    xlsb_cell cell;
    rich_str value;
};

struct brt_cell_isst {
    xlsb_cell cell;
    uint32_t isst;
};

#pragma pack(push,1)

struct brt_xf {
    uint16_t ixfeParent;
    uint16_t iFmt;
    uint16_t iFont;
    uint16_t iFill;
    uint16_t ixBOrder;
    uint8_t trot;
    uint8_t indent;
    uint16_t f123Prefix : 1;
    uint16_t fSxButton : 1;
    uint16_t fHidden : 1;
    uint16_t fLocked : 1;
    uint16_t iReadingOrder : 2;
    uint16_t fMergeCell : 1;
    uint16_t fShrinkToFit : 1;
    uint16_t fJustLast : 1;
    uint16_t fWrap : 1;
    uint16_t alcv : 3;
    uint16_t alc : 3;
    uint16_t unused : 10;
    uint16_t xfGrbitAtr : 6;
};

#pragma pack(pop)

struct brt_wb_prop {
    uint32_t f1904 : 1;
    uint32_t reserved1 : 1;
    uint32_t fHideBorderUnselLists : 1;
    uint32_t fFilterPrivacy : 1;
    uint32_t fBuggedUserABoutSolution : 1;
    uint32_t fShowInkAnnotation : 1;
    uint32_t fBackup : 1;
    uint32_t fNoSaveSup : 1;
    uint32_t grbitUpdateLinks : 2;
    uint32_t fHidePivotTableFList : 1;
    uint32_t fPublishedBookItems : 1;
    uint32_t fCheckCompat : 1;
    uint32_t mdDspObj : 2;
    uint32_t fShowPivotChartFilter : 1;
    uint32_t fAutoCompressPictures : 1;
    uint32_t reserved2 : 1;
    uint32_t fRefreshAll : 1;
    uint32_t unused : 13;
    uint32_t dwThemeVersion;
    uint32_t strName_len;
};

#pragma pack(push,1)

struct brt_fmt {
    uint16_t ifmt;
    uint32_t stFmtCode_len;
    char16_t stFmtCode[0];
};

#pragma pack(pop)
