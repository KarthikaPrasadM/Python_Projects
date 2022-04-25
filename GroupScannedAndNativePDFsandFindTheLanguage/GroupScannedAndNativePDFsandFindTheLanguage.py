# importing required modules
import os
# pip install pytest-shutil
import shutil
# pip install fitz
import fitz
# pip install langdetect
# pip install tesseract, important
from langdetect import detect
from langdetect import detect_langs

# Get PDFs file location into the variable "path"
path = input("Enter PDF file location: ")

# Get all the items in the directory as a list "ListOfAllItems"
ListOfAllItems = os.listdir(path)

# Iterate through each item's in the list "ListOfAllItems" to find the PDFs.
# Create a new list "ListOfPDFs" contains PDFs only.
ListOfPDFs = []
for item in ListOfAllItems:
    if item.endswith(".pdf") or item.endswith(".PDF"):
        ListOfPDFs.append(item)

# Create a new folder "Output"
ScannedPDFsPathlist = []
ScannedPDFsPath = path + r"\Non Scannable PDFs"
if "Non Scannable PDFs" not in ListOfAllItems:
    os.mkdir(ScannedPDFsPath)
else:
    pass
# Copy pdfs to another location
for eachpdf in ListOfPDFs:
    os.chdir(path)
    FilePath = os.path.abspath(eachpdf)
    shutil.copy(FilePath, ScannedPDFsPath)

detectedlanglist = []
# Create a new folder "Output"
OutputPath = path + "\Output"
if "Output" not in ListOfAllItems:
    os.mkdir(OutputPath)
else:
    pass

LanguageDict = {
    "ab": "Abkhazian",
    "aa": "Afar",
    "af": "Afrikaans",
    "ak": "Akan",
    "sq": "Albanian",
    "am": "Amharic",
    "ar": "Arabic",
    "an": "Aragonese",
    "hy": "Armenian",
    "as": "Assamese",
    "av": "Avaric",
    "ae": "Avestan",
    "ay": "Aymara",
    "az": "Azerbaijani",
    "bm": "Bambara",
    "ba": "Bashkir",
    "eu": "Basque",
    "be": "Belarusian",
    "bn": "Bengali",
    "bi": "Bislama",
    "bs": "Bosnian",
    "br": "Breton",
    "bg": "Bulgarian",
    "my": "Burmese",
    "ca": "Catalan, Valencian",
    "ch": "Chamorro",
    "ce": "Chechen",
    "ny": "Chichewa, Chewa, Nyanja",
    "zh": "Chinese",
    "cu": "Church Slavic, Old Slavonic, Church Slavonic, Old Bulgarian, Old Church Slavonic",
    "cv": "Chuvash",
    "kw": "Cornish",
    "co": "Corsican",
    "cr": "Cree",
    "hr": "Croatian",
    "cs": "Czech",
    "da": "Danish",
    "dv": "Divehi, Dhivehi, Maldivian",
    "nl": "Dutch, Flemish",
    "dz": "Dzongkha",
    "en": "English",
    "eo": "Esperanto",
    "et": "Estonian",
    "ee": "Ewe",
    "fo": "Faroese",
    "fj": "Fijian",
    "fi": "Finnish",
    "fr": "French",
    "fy": "Western Frisian",
    "ff": "Fulah",
    "gd": "Gaelic, Scottish Gaelic",
    "gl": "Galician",
    "lg": "Ganda",
    "ka": "Georgian",
    "de": "German",
    "el": "Greek",
    "kl": "Kalaallisut, Greenlandic",
    "gn": "Guarani",
    "gu": "Gujarati",
    "ht": "Haitian, Haitian Creole",
    "ha": "Hausa",
    "he": "Hebrew",
    "hz": "Herero",
    "hi": "Hindi",
    "ho": "Hiri Motu",
    "hu": "Hungarian",
    "is": "Icelandic",
    "io": "Ido",
    "ig": "Igbo",
    "id": "Indonesian",
    "ia": "Interlingua (International Auxiliary Language Association)",
    "ie": "Interlingue, Occidental",
    "iu": "Inuktitut",
    "ik": "Inupiaq",
    "ga": "Irish",
    "it": "Italian",
    "ja": "Japanese",
    "jv": "Javanese",
    "kn": "Kannada",
    "kr": "Kanuri",
    "ks": "Kashmiri",
    "kk": "Kazakh",
    "km": "Central Khmer",
    "ki": "Kikuyu, Gikuyu",
    "rw": "Kinyarwanda",
    "ky": "Kirghiz, Kyrgyz",
    "kv": "Komi",
    "kg": "Kongo",
    "ko": "Korean",
    "kj": "Kuanyama, Kwanyama",
    "ku": "Kurdish",
    "lo": "Lao",
    "la": "Latin",
    "lv": "Latvian",
    "li": "Limburgan, Limburger, Limburgish",
    "ln": "Lingala",
    "lt": "Lithuanian",
    "lu": "Luba-Katanga",
    "lb": "Luxembourgish",
    "mk": "Macedonian",
    "mg": "Malagasy",
    "ms": "Malay",
    "ml": "Malayalam",
    "mt": "Maltese",
    "gv": "Manx",
    "mi": "Maori",
    "mr": "Marathi",
    "mh": "Marshallese",
    "mn": "Mongolian",
    "na": "Nauru",
    "nv": "Navajo, Navaho",
    "nd": "North Ndebele",
    "nr": "South Ndebele",
    "ng": "Ndonga",
    "ne": "Nepali",
    "no": "Norwegian",
    "nb": "Norwegian Bokmål",
    "nn": "Norwegian Nynorsk",
    "ii": "Sichuan Yi, Nuosu",
    "oc": "Occitan",
    "oj": "Ojibwa",
    "or": "Oriya",
    "om": "Oromo",
    "os": "Ossetian, Ossetic",
    "pi": "Pali",
    "ps": "Pashto, Pushto",
    "fa": "Persian",
    "pl": "Polish",
    "pt": "Portuguese",
    "pa": "Panjabi",
    "qu": "Quechua",
    "ro": "Romanian, Moldavian, Moldovan",
    "rm": "Romansh",
    "rn": "Rundi",
    "ru": "Russian",
    "se": "Northern Sami",
    "sm": "Samoan",
    "sg": "Sango",
    "sa": "Sanskrit",
    "sc": "Sardinian",
    "sr": "Serbian",
    "sn": "Shona",
    "sd": "Sindhi",
    "si": "Sinhala",
    "sk": "Slovak",
    "sl": "Slovenian",
    "so": "Somali",
    "st": "Spanish, Castilian",
    "su": "Sundanese",
    "sw": "Swahili",
    "ss": "Swati",
    "sv": "Swedish",
    "tl": "Tagalog",
    "ty": "Tahitian",
    "tg": "Tajik",
    "ta": "Tamil",
    "tt": "Tatar",
    "te": "Telugu",
    "th": "Thai",
    "bo": "Tibetan",
    "ti": "Tigrinya",
    "to": "Tonga (Tonga Islands)",
    "ts": "Tsonga",
    "tn": "Tswana",
    "tr": "Turkish",
    "tk": "Turkmen",
    "tw": "Twi",
    "ug": "Uighur, Uyghur",
    "uk": "Ukrainian",
    "ur": "Urdu",
    "uz": "Uzbek",
    "ve": "Venda",
    "vi": "Vietnamese",
    "vo": "Volapük",
    "wa": "Walloon",
    "cy": "Welsh",
    "wo": "Wolof",
    "xh": "Xhosa",
    "yi": "Yiddish",
    "yo": "Yoruba",
    "za": "Zhuang, Chuang",
    "zu": "Zulu"
}

for PDF in ListOfPDFs:
    try:
        os.chdir(path)

        # text extraction from pdf using fitz module.
        ExtractedData = ""
        doc = fitz.open(PDF)
        for page in doc:
            ExtractedData += page.get_text()

        # Detect the language type
        LanguageType = detect(ExtractedData)

        # Detection accuracy
        DetectionAccuracy = detect_langs(ExtractedData)

        Language = LanguageDict[LanguageType]

        # mode
        mode = 0o666
        # Path
        LanguageFolder = os.path.join(OutputPath, Language)

        listOfFolders = os.listdir(OutputPath)

        if Language not in listOfFolders:
            # with mode 0o666
            os.mkdir(LanguageFolder, mode)
        else:
            pass
        FilePath = os.path.abspath(PDF)

        shutil.move(FilePath, LanguageFolder)

        # print(pdf)
        detectedlanglist.append(PDF)

    except:
        pass

scannablepdfslist = []
os.chdir(OutputPath)
langfolders = os.listdir(OutputPath)
for langfolder in langfolders:
    abspath = os.path.abspath(langfolder)
    listofpdfs = os.listdir(abspath)
    for i in listofpdfs:
        scannablepdfslist.append(i)

print(scannablepdfslist)

for detele in ListOfPDFs:
    if detele in scannablepdfslist:
        os.chdir(ScannedPDFsPath)
        file_path = os.path.abspath(detele)
        os.remove(file_path)
