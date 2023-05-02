#include <fstream>
#include <iostream>
#include "cfbf.h"
#include "sha1.h"

using namespace std;

enum class record_type : uint16_t {
    Formula = 6,
    Eof = 10,
    CalcCount = 12,
    CalcMode = 13,
    CalcPrecision = 14,
    CalcRefMode = 15,
    CalcDelta = 16,
    CalcIter = 17,
    Protect = 18,
    Password = 19,
    Header = 20,
    Footer = 21,
    ExternSheet = 23,
    Lbl = 24,
    WinProtect = 25,
    VerticalPageBreaks = 26,
    HorizontalPageBreaks = 27,
    Note = 28,
    Selection = 29,
    Date1904 = 34,
    ExternName = 35,
    LeftMargin = 38,
    RightMargin = 39,
    TopMargin = 40,
    BottomMargin = 41,
    PrintRowCol = 42,
    PrintGrid = 43,
    FilePass = 47,
    Font = 49,
    PrintSize = 51,
    Continue = 60,
    Window1 = 61,
    Backup = 64,
    Pane = 65,
    CodePage = 66,
    Pls = 77,
    DCon = 80,
    DConRef = 81,
    DConName = 82,
    DefColWidth = 85,
    XCT = 89,
    CRN = 90,
    FileSharing = 91,
    WriteAccess = 92,
    Obj = 93,
    Uncalced = 94,
    CalcSaveRecalc = 95,
    Template = 96,
    Intl = 97,
    ObjProtect = 99,
    ColInfo = 125,
    Guts = 128,
    WsBool = 129,
    GridSet = 130,
    HCenter = 131,
    VCenter = 132,
    BoundSheet8 = 133,
    WriteProtect = 134,
    Country = 140,
    HideObj = 141,
    Sort = 144,
    Palette = 146,
    Sync = 151,
    LPr = 152,
    DxGCol = 153,
    FnGroupName = 154,
    FilterMode = 155,
    BuiltInFnGroupCount = 156,
    AutoFilterInfo = 157,
    AutoFilter = 158,
    Scl = 160,
    Setup = 161,
    ScenMan = 174,
    SCENARIO = 175,
    SxView = 176,
    Sxvd = 177,
    SXVI = 178,
    SxIvd = 180,
    SXLI = 181,
    SXPI = 182,
    DocRoute = 184,
    RecipName = 185,
    MulRk = 189,
    MulBlank = 190,
    Mms = 193,
    SXDI = 197,
    SXDB = 198,
    SXFDB = 199,
    SXDBB = 200,
    SXNum = 201,
    SxBool = 202,
    SxErr = 203,
    SXInt = 204,
    SXString = 205,
    SXDtr = 206,
    SxNil = 207,
    SXTbl = 208,
    SXTBRGIITM = 209,
    SxTbpg = 210,
    ObProj = 211,
    SXStreamID = 213,
    DBCell = 215,
    SXRng = 216,
    SxIsxoper = 217,
    BookBool = 218,
    DbOrParamQry = 220,
    ScenarioProtect = 221,
    OleObjectSize = 222,
    XF = 224,
    InterfaceHdr = 225,
    InterfaceEnd = 226,
    SXVS = 227,
    MergeCells = 229,
    BkHim = 233,
    MsoDrawingGroup = 235,
    MsoDrawing = 236,
    MsoDrawingSelection = 237,
    PhoneticInfo = 239,
    SxRule = 240,
    SXEx = 241,
    SxFilt = 242,
    SxDXF = 244,
    SxItm = 245,
    SxName = 246,
    SxSelect = 247,
    SXPair = 248,
    SxFmla = 249,
    SxFormat = 251,
    SST = 252,
    LabelSst = 253,
    ExtSST = 255,
    SXVDEx = 256,
    SXFormula = 259,
    SXDBEx = 290,
    RRDInsDel = 311,
    RRDHead = 312,
    RRDChgCell = 315,
    RRTabId = 317,
    RRDRenSheet = 318,
    RRSort = 319,
    RRDMove = 320,
    RRFormat = 330,
    RRAutoFmt = 331,
    RRInsertSh = 333,
    RRDMoveBegin = 334,
    RRDMoveEnd = 335,
    RRDInsDelBegin = 336,
    RRDInsDelEnd = 337,
    RRDConflict = 338,
    RRDDefName = 339,
    RRDRstEtxp = 340,
    LRng = 351,
    UsesELFs = 352,
    DSF = 353,
    CUsr = 401,
    CbUsr = 402,
    UsrInfo = 403,
    UsrExcl = 404,
    FileLock = 405,
    RRDInfo = 406,
    BCUsrs = 407,
    UsrChk = 408,
    UserBView = 425,
    UserSViewBegin = 426,
    UserSViewBegin_Chart = 426,
    UserSViewEnd = 427,
    RRDUserView = 428,
    Qsi = 429,
    SupBook = 430,
    Prot4Rev = 431,
    CondFmt = 432,
    CF = 433,
    DVal = 434,
    DConBin = 437,
    TxO = 438,
    RefreshAll = 439,
    HLink = 440,
    Lel = 441,
    CodeName = 442,
    SXFDBType = 443,
    Prot4RevPass = 444,
    ObNoMacros = 445,
    Dv = 446,
    Excel9File = 448,
    RecalcId = 449,
    EntExU2 = 450,
    Dimensions = 512,
    Blank = 513,
    Number = 515,
    Label = 516,
    BoolErr = 517,
    String = 519,
    Row = 520,
    Index = 523,
    Array = 545,
    DefaultRowHeight = 549,
    Table = 566,
    Window2 = 574,
    RK = 638,
    Style = 659,
    BigName = 1048,
    Format = 1054,
    ContinueBigName = 1084,
    ShrFmla = 1212,
    HLinkTooltip = 2048,
    WebPub = 2049,
    QsiSXTag = 2050,
    DBQueryExt = 2051,
    ExtString = 2052,
    TxtQry = 2053,
    Qsir = 2054,
    Qsif = 2055,
    RRDTQSIF = 2056,
    BOF = 2057,
    OleDbConn = 2058,
    WOpt = 2059,
    SXViewEx = 2060,
    SXTH = 2061,
    SXPIEx = 2062,
    SXVDTEx = 2063,
    SXViewEx9 = 2064,
    ContinueFrt = 2066,
    RealTimeData = 2067,
    ChartFrtInfo = 2128,
    FrtWrapper = 2129,
    StartBlock = 2130,
    EndBlock = 2131,
    StartObject = 2132,
    EndObject = 2133,
    CatLab = 2134,
    YMult = 2135,
    SXViewLink = 2136,
    PivotChartBits = 2137,
    FrtFontList = 2138,
    SheetExt = 2146,
    BookExt = 2147,
    SXAddl = 2148,
    CrErr = 2149,
    HFPicture = 2150,
    FeatHdr = 2151,
    Feat = 2152,
    DataLabExt = 2154,
    DataLabExtContents = 2155,
    CellWatch = 2156,
    FeatHdr11 = 2161,
    Feature11 = 2162,
    DropDownObjIds = 2164,
    ContinueFrt11 = 2165,
    DConn = 2166,
    List12 = 2167,
    Feature12 = 2168,
    CondFmt12 = 2169,
    CF12 = 2170,
    CFEx = 2171,
    XFCRC = 2172,
    XFExt = 2173,
    AutoFilter12 = 2174,
    ContinueFrt12 = 2175,
    MDTInfo = 2180,
    MDXStr = 2181,
    MDXTuple = 2182,
    MDXSet = 2183,
    MDXProp = 2184,
    MDXKPI = 2185,
    MDB = 2186,
    PLV = 2187,
    Compat12 = 2188,
    DXF = 2189,
    TableStyles = 2190,
    TableStyle = 2191,
    TableStyleElement = 2192,
    StyleExt = 2194,
    NamePublish = 2195,
    NameCmt = 2196,
    SortData = 2197,
    Theme = 2198,
    GUIDTypeLib = 2199,
    FnGrp12 = 2200,
    NameFnGrp12 = 2201,
    MTRSettings = 2202,
    CompressPictures = 2203,
    HeaderFooter = 2204,
    CrtLayout12 = 2205,
    CrtMlFrt = 2206,
    CrtMlFrtContinue = 2207,
    ForceFullCalculation = 2211,
    ShapePropsStream = 2212,
    TextPropsStream = 2213,
    RichTextStream = 2214,
    CrtLayout12A = 2215,
    Units = 4097,
    Chart = 4098,
    Series = 4099,
    DataFormat = 4102,
    LineFormat = 4103,
    MarkerFormat = 4105,
    AreaFormat = 4106,
    PieFormat = 4107,
    AttachedLabel = 4108,
    SeriesText = 4109,
    ChartFormat = 4116,
    Legend = 4117,
    SeriesList = 4118,
    Bar = 4119,
    Line = 4120,
    Pie = 4121,
    Area = 4122,
    Scatter = 4123,
    CrtLine = 4124,
    Axis = 4125,
    Tick = 4126,
    ValueRange = 4127,
    CatSerRange = 4128,
    AxisLine = 4129,
    CrtLink = 4130,
    DefaultText = 4132,
    Text = 4133,
    FontX = 4134,
    ObjectLink = 4135,
    Frame = 4146,
    Begin = 4147,
    End = 4148,
    PlotArea = 4149,
    Chart3d = 4154,
    PicF = 4156,
    DropBar = 4157,
    Radar = 4158,
    Surf = 4159,
    RadarArea = 4160,
    AxisParent = 4161,
    LegendException = 4163,
    ShtProps = 4164,
    SerToCrt = 4165,
    AxesUsed = 4166,
    SBaseRef = 4168,
    SerParent = 4170,
    SerAuxTrend = 4171,
    IFmtRecord = 4174,
    Pos = 4175,
    AlRuns = 4176,
    BRAI = 4177,
    SerAuxErrBar = 4187,
    ClrtClient = 4188,
    SerFmt = 4189,
    Chart3DBarShape = 4191,
    Fbi = 4192,
    BopPop = 4193,
    AxcExt = 4194,
    Dat = 4195,
    PlotGrowth = 4196,
    SIIndex = 4197,
    GelFrame = 4198,
    BopPopCustom = 4199,
    Fbi2 = 4200
};

template<>
struct std::formatter<enum record_type> {
    constexpr auto parse(format_parse_context& ctx) {
        auto it = ctx.begin();

        if (it != ctx.end() && *it != '}')
            throw format_error("invalid format");

        return it;
    }

    template<typename format_context>
    auto format(enum record_type t, format_context& ctx) const {
        switch (t) {
            case record_type::Formula:
                return format_to(ctx.out(), "Formula");
            case record_type::Eof:
                return format_to(ctx.out(), "Eof");
            case record_type::CalcCount:
                return format_to(ctx.out(), "CalcCount");
            case record_type::CalcMode:
                return format_to(ctx.out(), "CalcMode");
            case record_type::CalcPrecision:
                return format_to(ctx.out(), "CalcPrecision");
            case record_type::CalcRefMode:
                return format_to(ctx.out(), "CalcRefMode");
            case record_type::CalcDelta:
                return format_to(ctx.out(), "CalcDelta");
            case record_type::CalcIter:
                return format_to(ctx.out(), "CalcIter");
            case record_type::Protect:
                return format_to(ctx.out(), "Protect");
            case record_type::Password:
                return format_to(ctx.out(), "Password");
            case record_type::Header:
                return format_to(ctx.out(), "Header");
            case record_type::Footer:
                return format_to(ctx.out(), "Footer");
            case record_type::ExternSheet:
                return format_to(ctx.out(), "ExternSheet");
            case record_type::Lbl:
                return format_to(ctx.out(), "Lbl");
            case record_type::WinProtect:
                return format_to(ctx.out(), "WinProtect");
            case record_type::VerticalPageBreaks:
                return format_to(ctx.out(), "VerticalPageBreaks");
            case record_type::HorizontalPageBreaks:
                return format_to(ctx.out(), "HorizontalPageBreaks");
            case record_type::Note:
                return format_to(ctx.out(), "Note");
            case record_type::Selection:
                return format_to(ctx.out(), "Selection");
            case record_type::Date1904:
                return format_to(ctx.out(), "Date1904");
            case record_type::ExternName:
                return format_to(ctx.out(), "ExternName");
            case record_type::LeftMargin:
                return format_to(ctx.out(), "LeftMargin");
            case record_type::RightMargin:
                return format_to(ctx.out(), "RightMargin");
            case record_type::TopMargin:
                return format_to(ctx.out(), "TopMargin");
            case record_type::BottomMargin:
                return format_to(ctx.out(), "BottomMargin");
            case record_type::PrintRowCol:
                return format_to(ctx.out(), "PrintRowCol");
            case record_type::PrintGrid:
                return format_to(ctx.out(), "PrintGrid");
            case record_type::FilePass:
                return format_to(ctx.out(), "FilePass");
            case record_type::Font:
                return format_to(ctx.out(), "Font");
            case record_type::PrintSize:
                return format_to(ctx.out(), "PrintSize");
            case record_type::Continue:
                return format_to(ctx.out(), "Continue");
            case record_type::Window1:
                return format_to(ctx.out(), "Window1");
            case record_type::Backup:
                return format_to(ctx.out(), "Backup");
            case record_type::Pane:
                return format_to(ctx.out(), "Pane");
            case record_type::CodePage:
                return format_to(ctx.out(), "CodePage");
            case record_type::Pls:
                return format_to(ctx.out(), "Pls");
            case record_type::DCon:
                return format_to(ctx.out(), "DCon");
            case record_type::DConRef:
                return format_to(ctx.out(), "DConRef");
            case record_type::DConName:
                return format_to(ctx.out(), "DConName");
            case record_type::DefColWidth:
                return format_to(ctx.out(), "DefColWidth");
            case record_type::XCT:
                return format_to(ctx.out(), "XCT");
            case record_type::CRN:
                return format_to(ctx.out(), "CRN");
            case record_type::FileSharing:
                return format_to(ctx.out(), "FileSharing");
            case record_type::WriteAccess:
                return format_to(ctx.out(), "WriteAccess");
            case record_type::Obj:
                return format_to(ctx.out(), "Obj");
            case record_type::Uncalced:
                return format_to(ctx.out(), "Uncalced");
            case record_type::CalcSaveRecalc:
                return format_to(ctx.out(), "CalcSaveRecalc");
            case record_type::Template:
                return format_to(ctx.out(), "Template");
            case record_type::Intl:
                return format_to(ctx.out(), "Intl");
            case record_type::ObjProtect:
                return format_to(ctx.out(), "ObjProtect");
            case record_type::ColInfo:
                return format_to(ctx.out(), "ColInfo");
            case record_type::Guts:
                return format_to(ctx.out(), "Guts");
            case record_type::WsBool:
                return format_to(ctx.out(), "WsBool");
            case record_type::GridSet:
                return format_to(ctx.out(), "GridSet");
            case record_type::HCenter:
                return format_to(ctx.out(), "HCenter");
            case record_type::VCenter:
                return format_to(ctx.out(), "VCenter");
            case record_type::BoundSheet8:
                return format_to(ctx.out(), "BoundSheet8");
            case record_type::WriteProtect:
                return format_to(ctx.out(), "WriteProtect");
            case record_type::Country:
                return format_to(ctx.out(), "Country");
            case record_type::HideObj:
                return format_to(ctx.out(), "HideObj");
            case record_type::Sort:
                return format_to(ctx.out(), "Sort");
            case record_type::Palette:
                return format_to(ctx.out(), "Palette");
            case record_type::Sync:
                return format_to(ctx.out(), "Sync");
            case record_type::LPr:
                return format_to(ctx.out(), "LPr");
            case record_type::DxGCol:
                return format_to(ctx.out(), "DxGCol");
            case record_type::FnGroupName:
                return format_to(ctx.out(), "FnGroupName");
            case record_type::FilterMode:
                return format_to(ctx.out(), "FilterMode");
            case record_type::BuiltInFnGroupCount:
                return format_to(ctx.out(), "BuiltInFnGroupCount");
            case record_type::AutoFilterInfo:
                return format_to(ctx.out(), "AutoFilterInfo");
            case record_type::AutoFilter:
                return format_to(ctx.out(), "AutoFilter");
            case record_type::Scl:
                return format_to(ctx.out(), "Scl");
            case record_type::Setup:
                return format_to(ctx.out(), "Setup");
            case record_type::ScenMan:
                return format_to(ctx.out(), "ScenMan");
            case record_type::SCENARIO:
                return format_to(ctx.out(), "SCENARIO");
            case record_type::SxView:
                return format_to(ctx.out(), "SxView");
            case record_type::Sxvd:
                return format_to(ctx.out(), "Sxvd");
            case record_type::SXVI:
                return format_to(ctx.out(), "SXVI");
            case record_type::SxIvd:
                return format_to(ctx.out(), "SxIvd");
            case record_type::SXLI:
                return format_to(ctx.out(), "SXLI");
            case record_type::SXPI:
                return format_to(ctx.out(), "SXPI");
            case record_type::DocRoute:
                return format_to(ctx.out(), "DocRoute");
            case record_type::RecipName:
                return format_to(ctx.out(), "RecipName");
            case record_type::MulRk:
                return format_to(ctx.out(), "MulRk");
            case record_type::MulBlank:
                return format_to(ctx.out(), "MulBlank");
            case record_type::Mms:
                return format_to(ctx.out(), "Mms");
            case record_type::SXDI:
                return format_to(ctx.out(), "SXDI");
            case record_type::SXDB:
                return format_to(ctx.out(), "SXDB");
            case record_type::SXFDB:
                return format_to(ctx.out(), "SXFDB");
            case record_type::SXDBB:
                return format_to(ctx.out(), "SXDBB");
            case record_type::SXNum:
                return format_to(ctx.out(), "SXNum");
            case record_type::SxBool:
                return format_to(ctx.out(), "SxBool");
            case record_type::SxErr:
                return format_to(ctx.out(), "SxErr");
            case record_type::SXInt:
                return format_to(ctx.out(), "SXInt");
            case record_type::SXString:
                return format_to(ctx.out(), "SXString");
            case record_type::SXDtr:
                return format_to(ctx.out(), "SXDtr");
            case record_type::SxNil:
                return format_to(ctx.out(), "SxNil");
            case record_type::SXTbl:
                return format_to(ctx.out(), "SXTbl");
            case record_type::SXTBRGIITM:
                return format_to(ctx.out(), "SXTBRGIITM");
            case record_type::SxTbpg:
                return format_to(ctx.out(), "SxTbpg");
            case record_type::ObProj:
                return format_to(ctx.out(), "ObProj");
            case record_type::SXStreamID:
                return format_to(ctx.out(), "SXStreamID");
            case record_type::DBCell:
                return format_to(ctx.out(), "DBCell");
            case record_type::SXRng:
                return format_to(ctx.out(), "SXRng");
            case record_type::SxIsxoper:
                return format_to(ctx.out(), "SxIsxoper");
            case record_type::BookBool:
                return format_to(ctx.out(), "BookBool");
            case record_type::DbOrParamQry:
                return format_to(ctx.out(), "DbOrParamQry");
            case record_type::ScenarioProtect:
                return format_to(ctx.out(), "ScenarioProtect");
            case record_type::OleObjectSize:
                return format_to(ctx.out(), "OleObjectSize");
            case record_type::XF:
                return format_to(ctx.out(), "XF");
            case record_type::InterfaceHdr:
                return format_to(ctx.out(), "InterfaceHdr");
            case record_type::InterfaceEnd:
                return format_to(ctx.out(), "InterfaceEnd");
            case record_type::SXVS:
                return format_to(ctx.out(), "SXVS");
            case record_type::MergeCells:
                return format_to(ctx.out(), "MergeCells");
            case record_type::BkHim:
                return format_to(ctx.out(), "BkHim");
            case record_type::MsoDrawingGroup:
                return format_to(ctx.out(), "MsoDrawingGroup");
            case record_type::MsoDrawing:
                return format_to(ctx.out(), "MsoDrawing");
            case record_type::MsoDrawingSelection:
                return format_to(ctx.out(), "MsoDrawingSelection");
            case record_type::PhoneticInfo:
                return format_to(ctx.out(), "PhoneticInfo");
            case record_type::SxRule:
                return format_to(ctx.out(), "SxRule");
            case record_type::SXEx:
                return format_to(ctx.out(), "SXEx");
            case record_type::SxFilt:
                return format_to(ctx.out(), "SxFilt");
            case record_type::SxDXF:
                return format_to(ctx.out(), "SxDXF");
            case record_type::SxItm:
                return format_to(ctx.out(), "SxItm");
            case record_type::SxName:
                return format_to(ctx.out(), "SxName");
            case record_type::SxSelect:
                return format_to(ctx.out(), "SxSelect");
            case record_type::SXPair:
                return format_to(ctx.out(), "SXPair");
            case record_type::SxFmla:
                return format_to(ctx.out(), "SxFmla");
            case record_type::SxFormat:
                return format_to(ctx.out(), "SxFormat");
            case record_type::SST:
                return format_to(ctx.out(), "SST");
            case record_type::LabelSst:
                return format_to(ctx.out(), "LabelSst");
            case record_type::ExtSST:
                return format_to(ctx.out(), "ExtSST");
            case record_type::SXVDEx:
                return format_to(ctx.out(), "SXVDEx");
            case record_type::SXFormula:
                return format_to(ctx.out(), "SXFormula");
            case record_type::SXDBEx:
                return format_to(ctx.out(), "SXDBEx");
            case record_type::RRDInsDel:
                return format_to(ctx.out(), "RRDInsDel");
            case record_type::RRDHead:
                return format_to(ctx.out(), "RRDHead");
            case record_type::RRDChgCell:
                return format_to(ctx.out(), "RRDChgCell");
            case record_type::RRTabId:
                return format_to(ctx.out(), "RRTabId");
            case record_type::RRDRenSheet:
                return format_to(ctx.out(), "RRDRenSheet");
            case record_type::RRSort:
                return format_to(ctx.out(), "RRSort");
            case record_type::RRDMove:
                return format_to(ctx.out(), "RRDMove");
            case record_type::RRFormat:
                return format_to(ctx.out(), "RRFormat");
            case record_type::RRAutoFmt:
                return format_to(ctx.out(), "RRAutoFmt");
            case record_type::RRInsertSh:
                return format_to(ctx.out(), "RRInsertSh");
            case record_type::RRDMoveBegin:
                return format_to(ctx.out(), "RRDMoveBegin");
            case record_type::RRDMoveEnd:
                return format_to(ctx.out(), "RRDMoveEnd");
            case record_type::RRDInsDelBegin:
                return format_to(ctx.out(), "RRDInsDelBegin");
            case record_type::RRDInsDelEnd:
                return format_to(ctx.out(), "RRDInsDelEnd");
            case record_type::RRDConflict:
                return format_to(ctx.out(), "RRDConflict");
            case record_type::RRDDefName:
                return format_to(ctx.out(), "RRDDefName");
            case record_type::RRDRstEtxp:
                return format_to(ctx.out(), "RRDRstEtxp");
            case record_type::LRng:
                return format_to(ctx.out(), "LRng");
            case record_type::UsesELFs:
                return format_to(ctx.out(), "UsesELFs");
            case record_type::DSF:
                return format_to(ctx.out(), "DSF");
            case record_type::CUsr:
                return format_to(ctx.out(), "CUsr");
            case record_type::CbUsr:
                return format_to(ctx.out(), "CbUsr");
            case record_type::UsrInfo:
                return format_to(ctx.out(), "UsrInfo");
            case record_type::UsrExcl:
                return format_to(ctx.out(), "UsrExcl");
            case record_type::FileLock:
                return format_to(ctx.out(), "FileLock");
            case record_type::RRDInfo:
                return format_to(ctx.out(), "RRDInfo");
            case record_type::BCUsrs:
                return format_to(ctx.out(), "BCUsrs");
            case record_type::UsrChk:
                return format_to(ctx.out(), "UsrChk");
            case record_type::UserBView:
                return format_to(ctx.out(), "UserBView");
            case record_type::UserSViewBegin:
                return format_to(ctx.out(), "UserSViewBegin");
            case record_type::UserSViewEnd:
                return format_to(ctx.out(), "UserSViewEnd");
            case record_type::RRDUserView:
                return format_to(ctx.out(), "RRDUserView");
            case record_type::Qsi:
                return format_to(ctx.out(), "Qsi");
            case record_type::SupBook:
                return format_to(ctx.out(), "SupBook");
            case record_type::Prot4Rev:
                return format_to(ctx.out(), "Prot4Rev");
            case record_type::CondFmt:
                return format_to(ctx.out(), "CondFmt");
            case record_type::CF:
                return format_to(ctx.out(), "CF");
            case record_type::DVal:
                return format_to(ctx.out(), "DVal");
            case record_type::DConBin:
                return format_to(ctx.out(), "DConBin");
            case record_type::TxO:
                return format_to(ctx.out(), "TxO");
            case record_type::RefreshAll:
                return format_to(ctx.out(), "RefreshAll");
            case record_type::HLink:
                return format_to(ctx.out(), "HLink");
            case record_type::Lel:
                return format_to(ctx.out(), "Lel");
            case record_type::CodeName:
                return format_to(ctx.out(), "CodeName");
            case record_type::SXFDBType:
                return format_to(ctx.out(), "SXFDBType");
            case record_type::Prot4RevPass:
                return format_to(ctx.out(), "Prot4RevPass");
            case record_type::ObNoMacros:
                return format_to(ctx.out(), "ObNoMacros");
            case record_type::Dv:
                return format_to(ctx.out(), "Dv");
            case record_type::Excel9File:
                return format_to(ctx.out(), "Excel9File");
            case record_type::RecalcId:
                return format_to(ctx.out(), "RecalcId");
            case record_type::EntExU2:
                return format_to(ctx.out(), "EntExU2");
            case record_type::Dimensions:
                return format_to(ctx.out(), "Dimensions");
            case record_type::Blank:
                return format_to(ctx.out(), "Blank");
            case record_type::Number:
                return format_to(ctx.out(), "Number");
            case record_type::Label:
                return format_to(ctx.out(), "Label");
            case record_type::BoolErr:
                return format_to(ctx.out(), "BoolErr");
            case record_type::String:
                return format_to(ctx.out(), "String");
            case record_type::Row:
                return format_to(ctx.out(), "Row");
            case record_type::Index:
                return format_to(ctx.out(), "Index");
            case record_type::Array:
                return format_to(ctx.out(), "Array");
            case record_type::DefaultRowHeight:
                return format_to(ctx.out(), "DefaultRowHeight");
            case record_type::Table:
                return format_to(ctx.out(), "Table");
            case record_type::Window2:
                return format_to(ctx.out(), "Window2");
            case record_type::RK:
                return format_to(ctx.out(), "RK");
            case record_type::Style:
                return format_to(ctx.out(), "Style");
            case record_type::BigName:
                return format_to(ctx.out(), "BigName");
            case record_type::Format:
                return format_to(ctx.out(), "Format");
            case record_type::ContinueBigName:
                return format_to(ctx.out(), "ContinueBigName");
            case record_type::ShrFmla:
                return format_to(ctx.out(), "ShrFmla");
            case record_type::HLinkTooltip:
                return format_to(ctx.out(), "HLinkTooltip");
            case record_type::WebPub:
                return format_to(ctx.out(), "WebPub");
            case record_type::QsiSXTag:
                return format_to(ctx.out(), "QsiSXTag");
            case record_type::DBQueryExt:
                return format_to(ctx.out(), "DBQueryExt");
            case record_type::ExtString:
                return format_to(ctx.out(), "ExtString");
            case record_type::TxtQry:
                return format_to(ctx.out(), "TxtQry");
            case record_type::Qsir:
                return format_to(ctx.out(), "Qsir");
            case record_type::Qsif:
                return format_to(ctx.out(), "Qsif");
            case record_type::RRDTQSIF:
                return format_to(ctx.out(), "RRDTQSIF");
            case record_type::BOF:
                return format_to(ctx.out(), "BOF");
            case record_type::OleDbConn:
                return format_to(ctx.out(), "OleDbConn");
            case record_type::WOpt:
                return format_to(ctx.out(), "WOpt");
            case record_type::SXViewEx:
                return format_to(ctx.out(), "SXViewEx");
            case record_type::SXTH:
                return format_to(ctx.out(), "SXTH");
            case record_type::SXPIEx:
                return format_to(ctx.out(), "SXPIEx");
            case record_type::SXVDTEx:
                return format_to(ctx.out(), "SXVDTEx");
            case record_type::SXViewEx9:
                return format_to(ctx.out(), "SXViewEx9");
            case record_type::ContinueFrt:
                return format_to(ctx.out(), "ContinueFrt");
            case record_type::RealTimeData:
                return format_to(ctx.out(), "RealTimeData");
            case record_type::ChartFrtInfo:
                return format_to(ctx.out(), "ChartFrtInfo");
            case record_type::FrtWrapper:
                return format_to(ctx.out(), "FrtWrapper");
            case record_type::StartBlock:
                return format_to(ctx.out(), "StartBlock");
            case record_type::EndBlock:
                return format_to(ctx.out(), "EndBlock");
            case record_type::StartObject:
                return format_to(ctx.out(), "StartObject");
            case record_type::EndObject:
                return format_to(ctx.out(), "EndObject");
            case record_type::CatLab:
                return format_to(ctx.out(), "CatLab");
            case record_type::YMult:
                return format_to(ctx.out(), "YMult");
            case record_type::SXViewLink:
                return format_to(ctx.out(), "SXViewLink");
            case record_type::PivotChartBits:
                return format_to(ctx.out(), "PivotChartBits");
            case record_type::FrtFontList:
                return format_to(ctx.out(), "FrtFontList");
            case record_type::SheetExt:
                return format_to(ctx.out(), "SheetExt");
            case record_type::BookExt:
                return format_to(ctx.out(), "BookExt");
            case record_type::SXAddl:
                return format_to(ctx.out(), "SXAddl");
            case record_type::CrErr:
                return format_to(ctx.out(), "CrErr");
            case record_type::HFPicture:
                return format_to(ctx.out(), "HFPicture");
            case record_type::FeatHdr:
                return format_to(ctx.out(), "FeatHdr");
            case record_type::Feat:
                return format_to(ctx.out(), "Feat");
            case record_type::DataLabExt:
                return format_to(ctx.out(), "DataLabExt");
            case record_type::DataLabExtContents:
                return format_to(ctx.out(), "DataLabExtContents");
            case record_type::CellWatch:
                return format_to(ctx.out(), "CellWatch");
            case record_type::FeatHdr11:
                return format_to(ctx.out(), "FeatHdr11");
            case record_type::Feature11:
                return format_to(ctx.out(), "Feature11");
            case record_type::DropDownObjIds:
                return format_to(ctx.out(), "DropDownObjIds");
            case record_type::ContinueFrt11:
                return format_to(ctx.out(), "ContinueFrt11");
            case record_type::DConn:
                return format_to(ctx.out(), "DConn");
            case record_type::List12:
                return format_to(ctx.out(), "List12");
            case record_type::Feature12:
                return format_to(ctx.out(), "Feature12");
            case record_type::CondFmt12:
                return format_to(ctx.out(), "CondFmt12");
            case record_type::CF12:
                return format_to(ctx.out(), "CF12");
            case record_type::CFEx:
                return format_to(ctx.out(), "CFEx");
            case record_type::XFCRC:
                return format_to(ctx.out(), "XFCRC");
            case record_type::XFExt:
                return format_to(ctx.out(), "XFExt");
            case record_type::AutoFilter12:
                return format_to(ctx.out(), "AutoFilter12");
            case record_type::ContinueFrt12:
                return format_to(ctx.out(), "ContinueFrt12");
            case record_type::MDTInfo:
                return format_to(ctx.out(), "MDTInfo");
            case record_type::MDXStr:
                return format_to(ctx.out(), "MDXStr");
            case record_type::MDXTuple:
                return format_to(ctx.out(), "MDXTuple");
            case record_type::MDXSet:
                return format_to(ctx.out(), "MDXSet");
            case record_type::MDXProp:
                return format_to(ctx.out(), "MDXProp");
            case record_type::MDXKPI:
                return format_to(ctx.out(), "MDXKPI");
            case record_type::MDB:
                return format_to(ctx.out(), "MDB");
            case record_type::PLV:
                return format_to(ctx.out(), "PLV");
            case record_type::Compat12:
                return format_to(ctx.out(), "Compat12");
            case record_type::DXF:
                return format_to(ctx.out(), "DXF");
            case record_type::TableStyles:
                return format_to(ctx.out(), "TableStyles");
            case record_type::TableStyle:
                return format_to(ctx.out(), "TableStyle");
            case record_type::TableStyleElement:
                return format_to(ctx.out(), "TableStyleElement");
            case record_type::StyleExt:
                return format_to(ctx.out(), "StyleExt");
            case record_type::NamePublish:
                return format_to(ctx.out(), "NamePublish");
            case record_type::NameCmt:
                return format_to(ctx.out(), "NameCmt");
            case record_type::SortData:
                return format_to(ctx.out(), "SortData");
            case record_type::Theme:
                return format_to(ctx.out(), "Theme");
            case record_type::GUIDTypeLib:
                return format_to(ctx.out(), "GUIDTypeLib");
            case record_type::FnGrp12:
                return format_to(ctx.out(), "FnGrp12");
            case record_type::NameFnGrp12:
                return format_to(ctx.out(), "NameFnGrp12");
            case record_type::MTRSettings:
                return format_to(ctx.out(), "MTRSettings");
            case record_type::CompressPictures:
                return format_to(ctx.out(), "CompressPictures");
            case record_type::HeaderFooter:
                return format_to(ctx.out(), "HeaderFooter");
            case record_type::CrtLayout12:
                return format_to(ctx.out(), "CrtLayout12");
            case record_type::CrtMlFrt:
                return format_to(ctx.out(), "CrtMlFrt");
            case record_type::CrtMlFrtContinue:
                return format_to(ctx.out(), "CrtMlFrtContinue");
            case record_type::ForceFullCalculation:
                return format_to(ctx.out(), "ForceFullCalculation");
            case record_type::ShapePropsStream:
                return format_to(ctx.out(), "ShapePropsStream");
            case record_type::TextPropsStream:
                return format_to(ctx.out(), "TextPropsStream");
            case record_type::RichTextStream:
                return format_to(ctx.out(), "RichTextStream");
            case record_type::CrtLayout12A:
                return format_to(ctx.out(), "CrtLayout12A");
            case record_type::Units:
                return format_to(ctx.out(), "Units");
            case record_type::Chart:
                return format_to(ctx.out(), "Chart");
            case record_type::Series:
                return format_to(ctx.out(), "Series");
            case record_type::DataFormat:
                return format_to(ctx.out(), "DataFormat");
            case record_type::LineFormat:
                return format_to(ctx.out(), "LineFormat");
            case record_type::MarkerFormat:
                return format_to(ctx.out(), "MarkerFormat");
            case record_type::AreaFormat:
                return format_to(ctx.out(), "AreaFormat");
            case record_type::PieFormat:
                return format_to(ctx.out(), "PieFormat");
            case record_type::AttachedLabel:
                return format_to(ctx.out(), "AttachedLabel");
            case record_type::SeriesText:
                return format_to(ctx.out(), "SeriesText");
            case record_type::ChartFormat:
                return format_to(ctx.out(), "ChartFormat");
            case record_type::Legend:
                return format_to(ctx.out(), "Legend");
            case record_type::SeriesList:
                return format_to(ctx.out(), "SeriesList");
            case record_type::Bar:
                return format_to(ctx.out(), "Bar");
            case record_type::Line:
                return format_to(ctx.out(), "Line");
            case record_type::Pie:
                return format_to(ctx.out(), "Pie");
            case record_type::Area:
                return format_to(ctx.out(), "Area");
            case record_type::Scatter:
                return format_to(ctx.out(), "Scatter");
            case record_type::CrtLine:
                return format_to(ctx.out(), "CrtLine");
            case record_type::Axis:
                return format_to(ctx.out(), "Axis");
            case record_type::Tick:
                return format_to(ctx.out(), "Tick");
            case record_type::ValueRange:
                return format_to(ctx.out(), "ValueRange");
            case record_type::CatSerRange:
                return format_to(ctx.out(), "CatSerRange");
            case record_type::AxisLine:
                return format_to(ctx.out(), "AxisLine");
            case record_type::CrtLink:
                return format_to(ctx.out(), "CrtLink");
            case record_type::DefaultText:
                return format_to(ctx.out(), "DefaultText");
            case record_type::Text:
                return format_to(ctx.out(), "Text");
            case record_type::FontX:
                return format_to(ctx.out(), "FontX");
            case record_type::ObjectLink:
                return format_to(ctx.out(), "ObjectLink");
            case record_type::Frame:
                return format_to(ctx.out(), "Frame");
            case record_type::Begin:
                return format_to(ctx.out(), "Begin");
            case record_type::End:
                return format_to(ctx.out(), "End");
            case record_type::PlotArea:
                return format_to(ctx.out(), "PlotArea");
            case record_type::Chart3d:
                return format_to(ctx.out(), "Chart3d");
            case record_type::PicF:
                return format_to(ctx.out(), "PicF");
            case record_type::DropBar:
                return format_to(ctx.out(), "DropBar");
            case record_type::Radar:
                return format_to(ctx.out(), "Radar");
            case record_type::Surf:
                return format_to(ctx.out(), "Surf");
            case record_type::RadarArea:
                return format_to(ctx.out(), "RadarArea");
            case record_type::AxisParent:
                return format_to(ctx.out(), "AxisParent");
            case record_type::LegendException:
                return format_to(ctx.out(), "LegendException");
            case record_type::ShtProps:
                return format_to(ctx.out(), "ShtProps");
            case record_type::SerToCrt:
                return format_to(ctx.out(), "SerToCrt");
            case record_type::AxesUsed:
                return format_to(ctx.out(), "AxesUsed");
            case record_type::SBaseRef:
                return format_to(ctx.out(), "SBaseRef");
            case record_type::SerParent:
                return format_to(ctx.out(), "SerParent");
            case record_type::SerAuxTrend:
                return format_to(ctx.out(), "SerAuxTrend");
            case record_type::IFmtRecord:
                return format_to(ctx.out(), "IFmtRecord");
            case record_type::Pos:
                return format_to(ctx.out(), "Pos");
            case record_type::AlRuns:
                return format_to(ctx.out(), "AlRuns");
            case record_type::BRAI:
                return format_to(ctx.out(), "BRAI");
            case record_type::SerAuxErrBar:
                return format_to(ctx.out(), "SerAuxErrBar");
            case record_type::ClrtClient:
                return format_to(ctx.out(), "ClrtClient");
            case record_type::SerFmt:
                return format_to(ctx.out(), "SerFmt");
            case record_type::Chart3DBarShape:
                return format_to(ctx.out(), "Chart3DBarShape");
            case record_type::Fbi:
                return format_to(ctx.out(), "Fbi");
            case record_type::BopPop:
                return format_to(ctx.out(), "BopPop");
            case record_type::AxcExt:
                return format_to(ctx.out(), "AxcExt");
            case record_type::Dat:
                return format_to(ctx.out(), "Dat");
            case record_type::PlotGrowth:
                return format_to(ctx.out(), "PlotGrowth");
            case record_type::SIIndex:
                return format_to(ctx.out(), "SIIndex");
            case record_type::GelFrame:
                return format_to(ctx.out(), "GelFrame");
            case record_type::BopPopCustom:
                return format_to(ctx.out(), "BopPopCustom");
            case record_type::Fbi2:
                return format_to(ctx.out(), "Fbi2");
            default:
                return format_to(ctx.out(), "{:x}", (unsigned int)t);
        }
    }
};

struct xl_record {
    enum record_type type;
    uint16_t length;
};

static void dump_workbook(span<const uint8_t> wb) {
    while (wb.size() >= sizeof(xl_record)) {
        auto r = *(xl_record*)wb.data();

        cout << format("type = {}, length = {:x}", r.type, r.length);

        wb = wb.subspan(sizeof(xl_record));

        if (r.length == 0) {
            cout << endl;
            continue;
        }

        cout << ", ";

        for (unsigned int i = 0; i < r.length; i++) {
            cout << format("{:02x} ", wb[i]);
        }
        cout << endl;

        wb = wb.subspan(r.length);
    }
}

static void cfbf_test(const filesystem::path& fn) {
#ifdef _WIN32
    unique_handle hup{CreateFileW((LPCWSTR)fn.u16string().c_str(), FILE_READ_DATA | DELETE, FILE_SHARE_READ, nullptr, OPEN_EXISTING,
                                  FILE_ATTRIBUTE_NORMAL, nullptr)};
    if (hup.get() == INVALID_HANDLE_VALUE)
        throw last_error("CreateFile", GetLastError());
#else
    unique_handle hup{open(fn.string().c_str(), O_RDONLY)};

    if (hup.get() == -1)
        throw formatted_error("open failed (errno = {})", errno);
#endif

    mmap m(hup.get());

    auto mem = m.map();

    cfbf c(mem);
    string enc_info, enc_package, workbook;

    for (unsigned int num = 0; const auto& e : c.entries) {
//         fmt::print("{}\n", e.name);
//         fmt::print("  type = {:x}\n", (unsigned int)e.de.type);
//         fmt::print("  colour = {:x}\n", (unsigned int)e.de.colour);
//         fmt::print("  sid_left_sibling = {:x}\n", e.de.sid_left_sibling);
//         fmt::print("  sid_right_sibling = {:x}\n", e.de.sid_right_sibling);
//         fmt::print("  sid_child = {:x}\n", e.de.sid_child);
//         fmt::print("  clsid = {:x}\n", *(uint64_t*)e.de.clsid);
//         fmt::print("  user_flags = {:x}\n", e.de.user_flags);
//         fmt::print("  create_time = {:x}\n", e.de.create_time);
//         fmt::print("  modify_time = {:x}\n", e.de.modify_time);
//         fmt::print("  sect_start = {:x}\n", e.de.sect_start);
//         fmt::print("  size = {:x}\n", e.de.size);
//         fmt::print("--\n");

        if (num == 0) { // root
            num++;
            continue;
        }

        if (e.name == "/EncryptionInfo") {
            enc_info.resize(e.get_size());

            uint64_t off = 0;
            auto buf = span((std::byte*)enc_info.data(), enc_info.size());

            while (true) {
                auto size = e.read(buf, off);

                if (size == 0)
                    break;

                off += size;
            }
        } else if (e.name == "/EncryptedPackage") {
            enc_package.resize(e.get_size());

            uint64_t off = 0;
            auto buf = span((std::byte*)enc_package.data(), enc_package.size());

            while (true) {
                auto size = e.read(buf, off);

                if (size == 0)
                    break;

                off += size;
            }
        } else if (e.name == "/Workbook") {
            workbook.resize(e.get_size());

            uint64_t off = 0;
            auto buf = span((std::byte*)workbook.data(), workbook.size());

            while (true) {
                auto size = e.read(buf, off);

                if (size == 0)
                    break;

                off += size;
            }
        }

        ofstream f("file" + to_string(num));

        uint64_t off = 0;
        std::byte buf[4096];

        while (true) {
            auto size = e.read(buf, off);

            if (size == 0)
                break;

            f.write((char*)buf, size); // FIXME - throw exception if fails

            off += size;
        }

        num++;
    }

    if (!workbook.empty()) {
        dump_workbook(span((uint8_t*)workbook.data(), workbook.size()));
        return;
    }

    if (enc_info.empty())
        throw runtime_error("EncryptionInfo not found.");

    c.parse_enc_info(span((uint8_t*)enc_info.data(), enc_info.size()), u"password");
    c.decrypt(span((uint8_t*)enc_package.data(), enc_package.size()));
}

#if 0
static void test_sha1() {
    auto hash = sha1(span<uint8_t>());

    for (auto c : hash) {
        fmt::print("{:02x} ", c);
    }
    fmt::print("\n");

    auto sv = string_view("abc");

    {
        SHA1_CTX ctx;

        for (auto c : sv) {
            ctx.update(span((uint8_t*)&c, 1));
        }

        ctx.finalize(hash);
    }

    for (auto c : hash) {
        fmt::print("{:02x} ", c);
    }
    fmt::print("\n");
}

static void test_key() {
    const u16string_view password = u"Password1234_";
    array<uint8_t, 16> salt{0xe8, 0x82, 0x66, 0x49, 0x0c, 0x5b, 0xd1, 0xee, 0xbd, 0x2b, 0x43, 0x94, 0xe3, 0xf8, 0x30, 0xef};

    auto key = generate_key(password, salt);

    for (auto c : key) {
        fmt::print("{:02x} ", c);
    }
    fmt::print("\n");
}
#endif

int main(int argc, char* argv[]) {
    if (argc < 2) {
        cerr << "No filename given.\n";
        return 1;
    }

//     test_sha1();
//     test_key();

    try {
        cfbf_test(argv[1]);
    } catch (const exception& e) {
        cerr << e.what() << endl;
        return 1;
    }

    return 0;
}
