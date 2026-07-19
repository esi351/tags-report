import pandas as pd
import re
import os

# تنظیم مسیرها
current_dir = os.getcwd()
input_file = os.path.join(current_dir, "/workspaces/me/Daily Report.xlsx")
output_file = os.path.join(current_dir, "tag.xlsx")

print(f"در حال پردازش فایل: {input_file} ...")

if not os.path.exists(input_file):
    print("خطا: فایل Daily Report.xlsx پیدا نشد.")
    exit()

try:
    df = pd.read_excel(input_file)
except Exception as e:
    print(f"خطا در خواندن اکسل: {e}")
    exit()

# شناسایی ستون‌های H و M
cols_to_check = []
if 'H' in df.columns: cols_to_check.append('H')
if 'M' in df.columns: cols_to_check.append('M')

if not cols_to_check:
    # اگر نام ستون نبود، از اندیس استفاده کن (ستون 8 و 13 یعنی اندیس 7 و 12)
    if len(df.columns) >= 13:
        cols_to_check = [df.columns[7], df.columns[12]]
    else:
        print("ستون‌های مورد نظر یافت نشدند.")
        exit()

# لیست سیاه کلمات بی‌ربط (فارسی و انگلیسی) که نباید تگ باشند
BLACKLIST_WORDS = [
    # کلمات فارسی رایج در گزارش‌ها
    "کالیبره", "کالیبراسیون", "تست","تست","از","استارت","الکتروموتور", "تعویض", "نصب", "بررسی", "چک", "تنظیم", "سرویس",
    "موتور", "پمپ", "کمپرسور", "چیلر", "بویلر", "توربین", "ژنراتور", "مخزن", "شیر", "ولو",
    "فلان", "بهمان", "شد", "گردید", "گرفت", "داد", "کرد", "باش", "نیاز", "دارد", "بود", "هست",
    "لامپ", "فیوز", "کابل", "گلند", "فیکسچر", "JB", "تابلو", "سنسور", "گیج", "ترانسمیتر",
    "اقدام", "انجام", "ادامه", "کار", "شیفت", "شب", "روز", "صبح", "عصر", "ظهر",
    "مشکل", "خراب", "سالم", "اوکی", "ok", "نرمال", "غیرنرمال", "خطا", "فالت", "آلارم",
    "برق", "آب", "هوا", "گاز", "روغن", "سوخت", "آتیش", "حریق", "ایمنی", "تجهیزات", "ابزار",
    "دقیق", "مکانیک", "الکتریک", "سیویل", "ابزار دقیق", "تاسیسات", "بهره برداری", "نت",
    "آماده", "سازی", "راه اندازی", "خاموش", "روشن", "استارت", "استوپ", "تریپ", "ریست",
    "باز", "بسته", "نیمه باز", "نیمه بسته", "درصد", "بار", "فشار", "دما", "سطح", "دبی",
    "فلو", "جریان", "ولتاژ", "آمپر", "فرکانس", "توان", "قدرت", "گشتاور", "لرزش", "ارتعاش",
    "صدا", "بو", "رنگ", "دود", "نشتی", "چکه", "ریزش", "ریختن", "پاشیدن", "پاشش", "تمیز",
    "کثیف", "گرفتن", "بستن", "باز کردن", "پیچ کردن", "باز کردن", "سفت", "شل", "لق", "محکم",
    "جدید", "قدیم", "نو", "کهنه", "اصل", "فیک", "ایرانی", "خارجی", "چینی", "آلمانی", "ژاپنی",
    "کره ای", "هندی", "ترک", "اروپایی", "آمریکایی", "روسی", "فرانسوی", "ایتالیایی", "اسپانیایی",
    "انگلیسی", "کانادایی", "استرالیایی", "برزیلی", "آرژانتینی", "مکزیکی", "شیلیایی", "پرویی",
    "کلمبیایی", "ونزوئلایی", "اکوادوری", "بولیویایی", "پاراگوئه ای", "اوروگوئه ای", "گویانی",
    "سورینامی", "گینه ای", "سیرالئونی", "لیبریایی", "ساحل عاجی", "غانایی", "توگویی", "بنینی",
    "نیجریه ای", "کامرونی", "گابنی", "کنگویی", "آنگولایی", "نامیبیایی", "بوتسوانایی", "آفریقای جنوبی",
    "لسوتویی", "سوازیلندی", "موزامبیکی", "زیمبابوه ای", "زامبیایی", "مالاویایی", "تانزانایی",
    "کنیایی", "اوگاندایی", "رواندایی", "بوروندیایی", "اتیوپیایی", "اریتره ای", "جیبوتیایی",
    "سومالیایی", "سودانی", "جنوب سودانی", "مصری", "لیبیایی", "تونسایی", " الجزایری", "مراکشی",
    "موریتانیایی", "مالیایی", "نیجری", "چادی", "سودانی", "جمهوری آفریقای مرکزی", "کامرونی",
    "گینه استوایی", "سائوتومه ای", "پرنسیپی", "کیپ وردی", "کوموری", "ماداگاسکاری", "موریسی",
    "سیشلی", "مایوتی", "رئونیونی", "سنت هلنا", "اسنشن", "تریستان دا کونا", "والیس و فوتونا",
    "ساموآ", "ساموآی آمریکا", "تونگا", "فیجی", "وانواتو", "سالومون", "پاپوآ گینه نو", "ناورو",
    "کیریباتی", "تووالو", "پالائو", "میکرونزی", "مارشال", "گوام", "مارینای شمالی", "هاوایی",
    "آلاسکا", "کالیفرنیا", "تگزاس", "فلوریدا", "نیویورک", "ایلینوی", "پنسیلوانیا", "اوهایو",
    "جورجیا", "کارولینای شمالی", "میشیگان", "نیوجرسی", "ویرجینیا", "واشینگتن", "آریزونا",
    "ماساچوست", "تنسی", "ایندیانا", "میسوری", "مریلند", "ویسکانسین", "کلرادو", "مینسوتا",
    "کارولینای جنوبی", "آلاباما", "لوئیزیانا", "کنتاکی", "اورگان", "اکلاهما", "کانکتیکات",
    "یوتا", "آیووا", "نوادا", "آرکانزاس", "میسیسیپی", "کانزاس", "نیومکزیکو", "نبراسکا",
    "ویرجینیای غربی", "آیداهو", "هاوایی", "نیوهمپشایر", "ماین", "رود آیلند", "مونتانا",
    "دلaware", "داکوتای جنوبی", "داکوتای شمالی", "ورمونت", "وایومینگ", "آلبرتا", "بریتیش کلمبیا",
    "منیتوبا", "نیوبرانزویک", "نیوفاندلند و لابرادور", "نوا اسکوشیا", "انتاریو", "جزیره پرنس ادوارد",
    "کبک", "ساسکاچوان", "یوکان", "نورت وست تریتوریز", "نوناولت", "اونتاریو", "کبک", "نوا اسکوشیا",
    "نیوبرانزویک", "منیتوبا", "بریتیش کلمبیا", "جزیره پرنس ادوارد", "ساسکاچوان", "آلبرتا",
    "نیوفاندلند و لابرادور", "یوکان", "نورت وست تریتوریز", "نوناولت",
    # کلمات انگلیسی بی‌ربط
    "ACTION", "ACKNOWLEDGE", "ACKNOWLAGE", "ACTIVE", "ADD", "ADDRESS", "ADMIN", "AERTH",
    "AGITATOR", "AIRCRAFT", "ALARM", "ALIGN", "ALIGNMENT", "ALL", "ALYZER", "AMP", "AMPER",
    "ANOMALI", "ANOMALY", "ANTI", "ARANGE", "AREA", "ARRANGE", "ARRENG", "ARRENGE", "ASSIGN",
    "ATOMIZE", "AUTO", "BACK", "BACKUP", "BADY", "BAG", "BAGGING", "BAGING", "BALANCE",
    "BALAST", "BALL", "BALLAST", "BAORD", "BAR", "BARG", "BARRIER", "BASEMENT", "BASIN",
    "BATTER", "BATTERI", "BATTERY", "BATTON", "BATTR", "BATTRY", "BEAM", "BEARING", "BEEM",
    "BELLOWS", "BERTHING", "BIG", "BIGBAG", "BLIND", "BLOCK", "BLOW", "BLOWER", "BLUE",
    "BOARD", "BODY", "BOILER", "BOILLER", "BOLT", "BONO", "BOOK", "BOOST", "BOTTOM",
    "BOTTON", "BOTTUN", "BOX", "BOY", "BREAK", "BREAKER", "BRINE", "BRUSH", "BSPT",
    "BUBLLE", "BURNER", "BURNNER", "BURNURE", "BURRNER", "BUS", "BUSBAR", "BUTTER",
    "BUTTERFLY", "BUTTOM", "BUTTON", "BUTTUN", "BUZZER", "BYPASS", "CABBLE", "CABEL",
    "CABIN", "CABLE", "CAD", "CADWELD", "CALCULATE", "CALIB", "CALIBRATION", "CALIBRE",
    "CALIBRETION", "CALL", "CALY", "CAN", "CAP", "CAPACITY", "CAPILAR", "CAPILARY",
    "CAPILLARY", "CATWELD", "CBCT", "CELL", "CELOS", "CERTIFATE", "CFG", "CG", "CHAIN",
    "CHAMPER", "CHANGE", "CHANGER", "CHARGE", "CHARGER", "CHART", "CHEAK", "CHECK",
    "CHEEK", "CHEK", "CHELLER", "CHENGER", "CHILD", "CHILER", "CHILLD", "CHILLED",
    "CHILLER", "CHRAT", "CINETION", "CIOSE", "CIRCULATE", "CIRCULATION", "CIVIL",
    "CLAMP", "CLASS", "CLASSIFIER", "CLEAN", "CLIENT", "CLOCK", "CLOSE", "CMMD",
    "CMMMS", "CMMS", "CMS", "CMSS", "CNMMS", "CNNECTION", "CODE", "COLD", "COLIING",
    "COLLING", "COM", "COMBUSTION", "COMING", "COMMAND", "COMMON", "COMPRESSOR",
    "COMPRESSURE", "CONDANCE", "CONDANSATE", "CONDASATE", "CONDENSER", "CONDUIT",
    "CONDUT", "CONE", "CONECTION", "CONFIG", "CONFIGURATION", "CONNCT", "CONNCTION",
    "CONNECT", "CONNECTING", "CONNECTION", "CONNECTON", "CONSERVATOR", "CONTAC",
    "CONTACT", "CONTACTOR", "CONTOLER", "CONTOLS", "CONTOROL", "CONTRLER", "CONTROL",
    "CONTROLER", "CONTROLLER", "CONTROLS", "CONTROREL", "COOLER", "COOLING", "CORE",
    "CORROSION", "COS", "COULD", "COUPLE", "COUPLER", "COUPLING", "COVER", "CPU",
    "CRACK", "CUSTOMER", "CYCLE", "DAMAGE", "DATA", "DATE", "DAY", "DCS", "DEAD",
    "DEBUG", "DEFECT", "DELAY", "DELETE", "DELTA", "DEMAND", "DENIED", "DEPTH",
    "DESCRIPTION", "DESIGN", "DETECT", "DETECTOR", "DEV", "DEVICE", "DIAGNOSIS",
    "DIAGNOSTIC", "DIFF", "DIFFERENTIAL", "DIGITAL", "DIRTY", "DISABLE", "DISC",
    "DISCHARGE", "DISCONNECT", "DISCOVERY", "DISP", "DISPLAY", "DIST", "DISTANCE",
    "DISTRIBUTION", "DO", "DOC", "DOCUMENT", "DONE", "DOWN", "DRAIN", "DRAWING",
    "DRY", "DUAL", "DUMP", "DURING", "DUTY", "EARTH", "EASY", "ECO", "EDGE", "EDIT",
    "EFFECT", "EFFICIENCY", "EIGHT", "ELECTRIC", "ELECTRICAL", "ELECTRONIC", "ELEMENT",
    "EMERGENCY", "EMPTY", "ENABLE", "END", "ENERGY", "ENG", "ENGINE", "ENGINEER",
    "ENGLISH", "ENTER", "ENTRY", "ENV", "ENVIRONMENT", "EQ", "EQP", "EQUIP", "EQUIPMENT",
    "ERROR", "EST", "ETH", "ETHERNET", "EVENT", "EX", "EXCEED", "EXCEL", "EXCEPTION",
    "EXCHANGE", "EXEC", "EXECUTE", "EXHAUST", "EXIT", "EXP", "EXPANSION", "EXPLOSION",
    "EXPORT", "EXT", "EXTERNAL", "FAIL", "FAILURE", "FALSE", "FAN", "FAST", "FAULT",
    "FEED", "FEEDBACK", "FIELD", "FILE", "FILL", "FILTER", "FINAL", "FIND", "FIRE",
    "FIRST", "FIX", "FLAG", "FLAME", "FLANGE", "FLASH", "FLOW", "FLUID", "FORCE",
    "FORM", "FORMAT", "FORWARD", "FOUND", "FOUR", "FREE", "FREQ", "FREQUENCY", "FROM",
    "FRONT", "FULL", "FUNCTION", "FUSE", "FUTURE", "GAIN", "GAS", "GATE", "GAUGE",
    "GENERAL", "GENERATOR", "GET", "GIVE", "GLOBAL", "GO", "GOOD", "GROUND", "GROUP",
    "GROWTH", "HAND", "HANDLE", "HARD", "HARDWARE", "HAVE", "HEAD", "HEALTH", "HEAT",
    "HEATER", "HEAVY", "HEIGHT", "HELP", "HIGH", "HIST", "HISTORY", "HOLD", "HOME",
    "HOOK", "HOT", "HOUR", "HOUSE", "HOW", "HUMAN", "HVAC", "HYD", "HYDRAULIC", "ID",
    "IDENT", "IDENTIFICATION", "IDENTIFY", "IDLE", "IF", "IGNORE", "IMAGE", "IMMEDIATE",
    "IMP", "IMPORT", "IMPORTANT", "IN", "INC", "INCLUDE", "INCLUDING", "INCOME", "INCOMING",
    "INCREASE", "IND", "INDEX", "INDICATE", "INDICATOR", "INDIRECT", "INDIVIDUAL", "INDUSTRY",
    "INFO", "INFORMATION", "INIT", "INITIAL", "INITIALIZE", "INPUT", "INS", "INSIDE",
    "INSPECT", "INSPECTION", "INSTALL", "INSTALLATION", "INSTANT", "INSTR", "INSTRUMENT",
    "INT", "INTEGER", "INTEGRAL", "INTEGRATE", "INTEGRITY", "INTER", "INTERFACE",
    "INTERNAL", "INTERNATIONAL", "INTERNET", "INTERVAL", "INTO", "INVALID", "INVENTORY",
    "INVERSE", "INVITE", "IO", "IP", "IR", "IS", "ISOLATION", "ISSUE", "ITEM", "JOB",
    "JOIN", "JOURNAL", "KEY", "KIT", "LABEL", "LABORATORY", "LACK", "LAG", "LAN", "LAST",
    "LATCH", "LATENCY", "LATER", "LAUNCH", "LAYER", "LEAK", "LEFT", "LEG", "LENGTH",
    "LEVEL", "LIB", "LIBRARY", "LIGHT", "LIGHTING", "LIKE", "LIMIT", "LINE", "LINK",
    "LIST", "LISTEN", "LOAD", "LOCAL", "LOCATE", "LOCATION", "LOCK", "LOG", "LOGGER",
    "LOGIC", "LONG", "LOOP", "LOSS", "LOST", "LOW", "MA", "MAIN", "MAINT", "MAINTENANCE",
    "MAKE", "MAN", "MANAGE", "MANAGER", "MANUAL", "MAP", "MARK", "MASTER", "MATCH",
    "MATERIAL", "MAX", "MAXIMUM", "MEAS", "MEASURE", "MEASUREMENT", "MECHANICAL", "MED",
    "MEDIUM", "MEM", "MEMORY", "MENU", "MERGE", "MESSAGE", "MET", "METER", "METHOD",
    "MIDDLE", "MIN", "MINIMUM", "MINOR", "MINUTE", "MODE", "MODEL", "MODIFY", "MODULE",
    "MONITOR", "MONTH", "MORE", "MOTOR", "MOVE", "MOVEMENT", "MULTI", "MUSIC", "NAME",
    "NATIVE", "NATURE", "NAV", "NAVIGATION", "NEAR", "NEG", "NEGATIVE", "NETWORK", "NEW",
    "NEXT", "NO", "NODE", "NOISE", "NONE", "NORMAL", "NOT", "NOTE", "NOTICE", "NOTIFY",
    "NOW", "NULL", "NUMBER", "NUMERIC", "OBJECT", "OBSERVE", "OCCUPANCY", "OCCUPIED",
    "OFF", "OFFSET", "OK", "OLD", "ON", "ONE", "ONLY", "OPEN", "OPER", "OPERATE",
    "OPERATION", "OPERATOR", "OPPOSITE", "OPT", "OPTION", "OR", "ORDER", "ORDINAL",
    "ORG", "ORGANIZATION", "ORIGINAL", "OTHER", "OUT", "OUTPUT", "OUTSIDE", "OVER",
    "OVERFLOW", "OVERLOAD", "OVERRIDE", "OWN", "PACKAGE", "PAGE", "PAIR", "PAN", "PANEL",
    "PARAM", "PARAMETER", "PART", "PARTIAL", "PASS", "PASSWORD", "PAST", "PATCH", "PATH",
    "PATTERN", "PAUSE", "PAY", "PAYMENT", "PC", "PEAK", "PER", "PERCENT", "PERFORMANCE",
    "PERIOD", "PERMANENT", "PERSON", "PHASE", "PHONE", "PHOTO", "PHYSICAL", "PI", "PID",
    "PIECE", "PIN", "PIPE", "PIPING", "PLACE", "PLAN", "PLANT", "PLATE", "PLAY", "PLOT",
    "PLUG", "PLUS", "POINT", "POLICY", "POS", "POSITION", "POSITIVE", "POWER", "PR",
    "PRE", "PRECISION", "PREF", "PREFERENCE", "PREFIX", "PREPARE", "PRESENT", "PRESS",
    "PRESSURE", "PREV", "PREVIOUS", "PRIMARY", "PRINT", "PRINTER", "PRIORITY", "PRIVATE",
    "PROC", "PROCESS", "PRODUCE", "PRODUCT", "PRODUCTION", "PROF", "PROFILE", "PROG",
    "PROGRAM", "PROJECT", "PROTECT", "PROTECTION", "PROTOCOL", "PUBLIC", "PULSE", "PUMP",
    "PURGE", "QUALITY", "QUANTITY", "QUERY", "QUEUE", "QUICK", "QUIT", "QUOTE", "RAD",
    "RADIO", "RANGE", "RATE", "RATIO", "RAW", "READ", "READER", "READY", "REAL", "REAR",
    "REASON", "REBOOT", "REC", "RECEIVE", "RECENT", "RECORD", "RECOVER", "RECOVERY",
    "RED", "REF", "REFERENCE", "REFRESH", "REG", "REGION", "REGIST", "REGISTER", "REGISTRY",
    "REGULAR", "REJECT", "REL", "RELATION", "RELATIVE", "RELEASE", "RELIABILITY", "REMOTE",
    "REMOVE", "RENAME", "RENEW", "REPEAT", "REPLACE", "REPLY", "REPORT", "REQ", "REQUEST",
    "RESET", "RESERVE", "RESIDUAL", "RESIST", "RESISTANCE", "RESOLUTION", "RESOURCE",
    "RESPONSE", "REST", "RESTART", "RESTORE", "RESULT", "RESUME", "RET", "RETURN", "REV",
    "REVERSE", "REVIEW", "REVISE", "RIGHT", "RING", "RISK", "ROLE", "ROOM", "ROT", "ROTARY",
    "ROTATE", "ROTATION", "ROUND", "ROUTE", "ROUTINE", "RUN", "SAFE", "SAFETY", "SAMPLE",
    "SAVE", "SCALE", "SCAN", "SCENE", "SCH", "SCHEDULE", "SCHEMA", "SCOPE", "SCREEN",
    "SCRIPT", "SEARCH", "SECOND", "SECONDARY", "SECRET", "SECTION", "SECURE", "SECURITY",
    "SEE", "SEEK", "SELECT", "SELECTION", "SELF", "SELL", "SEM", "SEND", "SENSOR", "SENT",
    "SEQ", "SEQUENCE", "SER", "SERIAL", "SERIES", "SERV", "SERVE", "SERVER", "SERVICE",
    "SET", "SETTING", "SETTLE", "SETUP", "SEVEN", "SEX", "SHAPE", "SHARE", "SHIFT", "SHIP",
    "SHOP", "SHORT", "SHOT", "SHOW", "SHUT", "SIDE", "SIGNAL", "SIGN", "SILENT", "SIM",
    "SIMPLE", "SIMULATION", "SINGLE", "SITE", "SIZE", "SKIP", "SLAVE", "SLEEP", "SLOT",
    "SLOW", "SMALL", "SMART", "SMOKE", "SOFT", "SOFTWARE", "SOL", "SOLAR", "SOLID",
    "SOLUTION", "SOLVE", "SOME", "SON", "SORT", "SOURCE", "SPACE", "SPARE", "SPEAKER",
    "SPECIAL", "SPECIFIC", "SPEED", "SPEND", "SPLIT", "SPOT", "SPRING", "STAB", "STABLE",
    "STAFF", "STAGE", "STALL", "STAND", "STANDARD", "START", "STATE", "STATEMENT",
    "STATIC", "STATUS", "STEP", "STOP", "STORAGE", "STORE", "STR", "STREAM", "STREET",
    "STRENGTH", "STRICT", "STRING", "STRIP", "STRUCT", "STRUCTURE", "STUDY", "STYLE",
    "SUB", "SUBJECT", "SUBMIT", "SUBNET", "SUBSCRIBE", "SUBSCRIPTION", "SUCCESS", "SUM",
    "SUMMARY", "SUN", "SUPER", "SUPERVISE", "SUPERVISOR", "SUPPLY", "SUPPORT", "SURE",
    "SURFACE", "SURGE", "SURVEY", "SW", "SWAP", "SWITCH", "SYM", "SYMBOL", "SYNC",
    "SYNCHRONIZE", "SYSTEM", "TABLE", "TAG", "TAIL", "TAKE", "TALK", "TANK", "TASK",
    "TEAM", "TEMP", "TEMPERATURE", "TEMPLATE", "TEN", "TERM", "TERMINAL", "TERMINATE",
    "TEST", "TEXT", "THAN", "THANK", "THE", "THEORY", "THREE", "THRESHOLD", "THROUGH",
    "TIME", "TIMER", "TIMESTAMP", "TITLE", "TO", "TODAY", "TOKEN", "TOLERANCE", "TON",
    "TOOL", "TOP", "TOTAL", "TOUCH", "TOWER", "TRACK", "TRADE", "TRAFFIC", "TRAIN",
    "TRAINING", "TRANSFER", "TRANSFORM", "TRANSMIT", "TRANSMITTER", "TRANSPORT", "TRAP",
    "TRAVEL", "TREND", "TRIGGER", "TRIP", "TRUE", "TRUNK", "TRUST", "TRY", "TUBE",
    "TURN", "TWO", "TYPE", "UNIT", "UNIVERSAL", "UNIVERSE", "UNIVERSITY", "UNLOCK",
    "UNPACK", "UNTIL", "UP", "UPDATE", "UPLOAD", "UPPER", "UPS", "UPSTREAM", "URGENCY",
    "URGENT", "URL", "USAGE", "USE", "USED", "USER", "UTIL", "UTILITY", "VACUUM", "VAL",
    "VALID", "VALIDATE", "VALUE", "VALVE", "VAN", "VAR", "VARIABLE", "VARIANT", "VARIATION",
    "VARIOUS", "VECTOR", "VEL", "VELOCITY", "VENDOR", "VENT", "VENTILATION", "VERIFY",
    "VERSION", "VERTICAL", "VERY", "VESSEL", "VIA", "VIDEO", "VIEW", "VIRTUAL", "VIS",
    "VISIBLE", "VISION", "VISIT", "VISITOR", "VOLT", "VOLTAGE", "VOL", "VOLUME", "WAIT",
    "WAKE", "WALK", "WAN", "WARNING", "WASH", "WATCH", "WATER", "WATT", "WAVE", "WAY",
    "WEATHER", "WEB", "WEEK", "WEIGHT", "WELCOME", "WEST", "WHAT", "WHEN", "WHERE",
    "WHICH", "WHITE", "WHO", "WHY", "WIDE", "WIDTH", "WIFI", "WIND", "WINDOW", "WIRE",
    "WIRELESS", "WITH", "WITHOUT", "WORK", "WORKER", "WORKFLOW", "WORKING", "WORKSHOP",
    "WORLD", "WRITE", "WRONG", "YEAR", "YELLOW", "YES", "YIELD", "ZONE",
    # کلمات ترکیبی اشتباه رایج
    "ROUTINE", "CHECKED", "INSTALLED", "REPLACED", "TESTED", "CALIBRATED", "CLEANED",
    "MAINTAINED", "SERVICED", "REPAIRED", "OVERHAULED", "INSPECTED", "VERIFIED",
    "ADJUSTED", "ALIGNED", "BALANCED", "LUBRICATED", "GREASED", "OILED", "PAINTED",
    "COATED", "INSULATED", "HEATED", "COOLED", "VENTILATED", "EXHAUSTED", "DRAINED",
    "FILLED", "EMPTY", "LOADED", "UNLOADED", "STARTED", "STOPPED", "RUNNING", "STOPPED",
    "TRIPPED", "RESETTED", "LOCKED", "UNLOCKED", "OPENED", "CLOSED", "ISOLATED",
    "BYPASSED", "CONNECTED", "DISCONNECTED", "WIRED", "CABLED", "TERMINATED", "GROUNDED",
    "EARTHED", "SHORTED", "BURNED", "MELTED", "BROKEN", "CRACKED", "LEAKED", "CORRODED",
    "RUSTED", "WORN", "DAMAGED", "FAILED", "MALFUNCTIONED", "ERROR", "FAULTY", "DEFECTIVE",
    "ABNORMAL", "IRREGULAR", "UNSTABLE", "UNSAFE", "DANGEROUS", "HAZARDOUS", "RISKY",
    "CRITICAL", "SEVERE", "MAJOR", "MINOR", "NEGLIGIBLE", "INSIGNIFICANT", "TRIVIAL",
    "IMPORTANT", "SIGNIFICANT", "ESSENTIAL", "VITAL", "CRUCIAL", "KEY", "PRIMARY",
    "SECONDARY", "TERTIARY", "AUXILIARY", "STANDBY", "RESERVE", "BACKUP", "REDUNDANT",
    "DUPLICATE", "TRIPLE", "QUADRUPLE", "MULTIPLE", "SINGLE", "DOUBLE", "TRIPLE",
    "QUAD", "DUAL", "BI", "TRI", "MULTI", "POLY", "OMNI", "UNI", "NON", "UN", "IN",
    "IM", "IR", "IL", "DIS", "MIS", "UNDER", "OVER", "SUPER", "SUB", "INTER", "INTRA",
    "TRANS", "CROSS", "ANTI", "PRO", "CON", "COM", "CO", "COL", "COR", "SYN", "SYM",
    "HOMO", "HETERO", "ISO", "ANISO", "MACRO", "MICRO", "NANO", "PICO", "FEMTO", "ATTO",
    "ZEPTO", "YOCTO", "KILO", "MEGA", "GIGA", "TERA", "PETA", "EXA", "ZETTA", "YOTTA"
]

def is_valid_tag(text):
    if not isinstance(text, str):
        return False
    
    text = text.strip()
    if not text:
        return False
    
    upper_text = text.upper()
    
    # 1. حذف مواردی که کاملاً در لیست سیاه هستند
    if upper_text in BLACKLIST_WORDS:
        return False
    
    # 2. حذف جملات طولانی (تگ‌ها معمولاً کوتاه هستند)
    if len(text.split()) > 4: 
        return False
        
    # 3. حذف متن‌های فارسی خالص (تگ‌ها لاتین هستند)
    if re.search(r'[\u0600-\u06FF]', text) and not re.search(r'[A-Za-z0-9\-/]', text):
        return False
    
    # 4. الگوی اصلی تگ‌های مهندسی
    # باید حداقل یک حرف و یک عدد داشته باشد (مثلا P-201 یا LT-100)
    # استثنا: تگ‌های خاصی مثل GA, GD که ممکن است بدون عدد باشند اما در لیست مجاز قرار می‌گیرند
    has_letter = bool(re.search(r'[A-Za-z]', text))
    has_digit = bool(re.search(r'\d', text))
    
    # اگر نه حرف دارد نه عدد، احتمالا کلمه بی‌معنی است (مگر اینکه در لیست سفید باشد)
    if not has_letter or not has_digit:
        # چک کردن لیست سفید برای موارد خاص تک کلمه‌ای (اختیاری)
        # فعلا فرض می‌کنیم تگ معتبر باید حرف و عدد داشته باشد
        return False

    # 5. حذف کلماتی که با افعال یا صفات شروع می‌شوند و شبیه تگ نیستند
    # مثلا "Installed P-201" -> باید فقط P-201 استخراج شود
    # این کار در مرحله Extract انجام می‌شود
    
    return True

def extract_tags_from_cell(cell_value):
    if not isinstance(cell_value, str):
        return []
    
    found_tags = []
    
    # الگوی پیدا کردن تگ‌ها در میان متن
    # الگو: (حروف اختصاری) + (خط تیره اختیاری) + (اعداد) + (حرف اختیاری)
    # مثال‌ها: P-201, P201, LT-100A, XV-500, GA-2001
    # این الگو سعی می‌کند تگ‌ها را از دل جملات بیرون بکشد
    
    # الگوی جامع برای تگ‌های صنعتی
    # شامل: پیشوند (2-5 حرف)، خط تیره اختیاری، شماره (1-6 رقم)، پسوند اختیاری (حرف یا ترکیب حروف)
    pattern = r'\b([A-Z]{2,5}-?\d{2,6}[A-Z]?)(?:\s|,|\.|$|/)'
    
    # جستجو در متن بزرگ شده
    matches = re.findall(pattern, cell_value.upper())
    
    for match in matches:
        tag = match.strip()
        if is_valid_tag(tag):
            found_tags.append(tag)
            
    # همچنین چک کنیم اگر کل سلول خودش یک تگ تمیز است
    if is_valid_tag(cell_value.strip()):
        # اگر الگوی بالا چیزی پیدا نکرد ولی کل سلول تگ بود
        if not found_tags: 
             found_tags.append(cell_value.strip().upper())
             
    return found_tags

all_raw_tags = []
for col in cols_to_check:
    for val in df[col].dropna():
        extracted = extract_tags_from_cell(str(val))
        all_raw_tags.extend(extracted)

# پردازش نهایی: نرمال‌سازی، اولویت‌بندی و بسط بازه‌ها
tag_map = {} # کلید: تگ بدون خط تیره، مقدار: لیست فرمت‌های موجود

for tag in all_raw_tags:
    # نرمال‌سازی برای کلید یکتا (حذف خط تیره)
    key = tag.replace('-', '')
    
    if key not in tag_map:
        tag_map[key] = []
    tag_map[key].append(tag)

final_tags = set()

for key, variations in tag_map.items():
    # 1. اولویت GA بر P
    # اگر هم GA-xxxx و هم P-xxxx داشتیم، GA را نگه دار
    has_ga = any(v.startswith('GA') for v in variations)
    
    selected = []
    if has_ga:
        selected = [v for v in variations if v.startswith('GA')]
    else:
        selected = variations
    
    # حذف تکراری‌های دقیق
    unique_selected = list(set(selected))
    
    # 2. بسط بازه‌ها (Range Expansion)
    for t in unique_selected:
        # چک کردن الگوی بازه: PREFIX-NUM/NUM
        range_match = re.match(r'^([A-Z]+)(\d+)/(\d+)$', t)
        if range_match:
            prefix = range_match.group(1)
            start = int(range_match.group(2))
            end = int(range_match.group(3))
            width = len(range_match.group(2)) # حفظ صفرهای اولیه
            
            for i in range(start, end + 1):
                num_str = str(i).zfill(width)
                final_tags.add(f"{prefix}{num_str}")
        else:
            final_tags.add(t)

# ساخت خروجی
result_list = sorted(list(final_tags))
result_df = pd.DataFrame(result_list, columns=['Tag'])

print(f"تعداد تگ‌های معتبر استخراج شده: {len(result_df)}")
print("نمونه تگ‌ها:", result_df.head(10)['Tag'].tolist())

# ذخیره فایل tag.xlsx
result_df.to_excel(output_file, index=False)
print(f"فایل {output_file} با موفقیت ساخته شد.")

# آپدیت فایل اصلی
try:
    with pd.ExcelWriter(input_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        result_df.to_excel(writer, sheet_name='Tags', index=False)
    print("شیت Tags در فایل Daily Report.xlsx بروزرسانی شد.")
except Exception as e:
    print(f"خطا در بروزرسانی فایل اصلی: {e}")

print("پایان عملیات.")