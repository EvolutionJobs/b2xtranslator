using System;
using b2xtranslator.DocFileFormat;
using b2xtranslator.CommonTranslatorLib;
using System.Xml;
using b2xtranslator.OpenXmlLib;
using System.Globalization;

namespace b2xtranslator.WordprocessingMLMapping
{
    public class LanguageIdMapping : PropertiesMapping, IMapping<LanguageId>
    {
        public enum LanguageType
        {
            Default,
            EastAsian,
            Complex
        }

        private XmlElement _parent;
        private LanguageType _type;

        public LanguageIdMapping(XmlWriter writer, LanguageType type)
            : base(writer)
        {
            this._type = type;
        }

        public LanguageIdMapping(XmlElement parentElement, LanguageType type)
            : base(null)
        {
            this._nodeFactory = parentElement.OwnerDocument;
            this._parent = parentElement;
            this._type = type;
        }

        public void Apply(LanguageId lid)
        {
            if (lid.Code != LanguageId.LanguageCode.Nothing)
            {
                string langcode = "en-US";

                try
                {
                    var ci = new CultureInfo((int)lid.Code);
                    langcode = ci.ToString();
                }
                catch (Exception) 
                {
                    //langcode = getLanguageCode(lid);
                }

                XmlAttribute att;
                switch (this._type)
                {
                    case LanguageType.Default:
                        att = this._nodeFactory.CreateAttribute("w", "val", OpenXmlNamespaces.WordprocessingML);
                        break;
                    case LanguageType.EastAsian:
                        att = this._nodeFactory.CreateAttribute("w", "eastAsia", OpenXmlNamespaces.WordprocessingML);
                        break;
                    case LanguageType.Complex:
                        att = this._nodeFactory.CreateAttribute("w", "bidi", OpenXmlNamespaces.WordprocessingML);
                        break;
                    default:
                        att = this._nodeFactory.CreateAttribute("w", "val", OpenXmlNamespaces.WordprocessingML);
                        break;
                }
                att.Value = langcode;

                if (this._writer != null)
                {
                    att.WriteTo(this._writer);
                }
                else if (this._parent != null)
                {
                    this._parent.Attributes.Append(att);
                }
            }
        }

        [Obsolete("Don't use this method because it's not complete")]
        private string getLanguageCode(LanguageId lid)
        {
            switch (lid.Code)
            {
                case LanguageId.LanguageCode.Afrikaans:
                    return "af-ZA";
                case LanguageId.LanguageCode.Albanian:
                    return "sq-AL";
                case LanguageId.LanguageCode.Amharic:
                    return "am-ET";
                case LanguageId.LanguageCode.ArabicAlgeria:
                    return "ar-DZ";
                case LanguageId.LanguageCode.ArabicBahrain:
                    return "ar-BH";
                case LanguageId.LanguageCode.ArabicEgypt:
                    return "ar-EG";
                case LanguageId.LanguageCode.ArabicIraq:
                    return "ar-IQ";
                case LanguageId.LanguageCode.ArabicJordan:
                    return "ar-JO";
                case LanguageId.LanguageCode.ArabicKuwait:
                    return "ar-KW";
                case LanguageId.LanguageCode.ArabicLebanon:
                    return "ar-LB";
                case LanguageId.LanguageCode.ArabicLibya:
                    return "ar-LY";
                case LanguageId.LanguageCode.ArabicMorocco:
                    return "ar-MA";
                case LanguageId.LanguageCode.ArabicOman:
                    return "ar-OM";
                case LanguageId.LanguageCode.ArabicQatar:
                    return "ar-QA";
                case LanguageId.LanguageCode.ArabicSaudiArabia:
                    return "ar-SA";
                case LanguageId.LanguageCode.ArabicSyria:
                    return "ar-SY";
                case LanguageId.LanguageCode.ArabicTunisia:
                    return "ar-TN";
                case LanguageId.LanguageCode.ArabicUAE:
                    return "ar-AE";
                case LanguageId.LanguageCode.ArabicYemen:
                    return "ar-YE";
                case LanguageId.LanguageCode.Armenian:
                    return "hy-AM";
                case LanguageId.LanguageCode.Assamese:
                    return "as-IN";
                case LanguageId.LanguageCode.AzeriCyrillic:
                    return "az-AZ-cyrl";
                case LanguageId.LanguageCode.AzeriLatin:
                    return "az-AZ-latn";
                case LanguageId.LanguageCode.Basque:
                    return "eu-ES";
                case LanguageId.LanguageCode.Belarusian:
                    return "be-BY";
                case LanguageId.LanguageCode.Bengali:
                    return "bn-IN";
                case LanguageId.LanguageCode.BengaliBangladesh:
                    return "bn-BD";
                case LanguageId.LanguageCode.Bulgarian:
                    return "bg-BG";
                case LanguageId.LanguageCode.Burmese:
                    return "my-MM";
                case LanguageId.LanguageCode.Catalan:
                    return "ca-ES";
                case LanguageId.LanguageCode.Cherokee:
                    //there is no iso code fpr cherokee
                case LanguageId.LanguageCode.ChineseHongKong:
                    return "zh-HK";
                case LanguageId.LanguageCode.ChineseMacao:
                    return "zh-MO";
                case LanguageId.LanguageCode.ChinesePRC:
                    return "zh-CN";
                case LanguageId.LanguageCode.ChineseSingapore:
                    return "zh-SG";
                case LanguageId.LanguageCode.ChineseTaiwan:
                    return "zh-TW";
                case LanguageId.LanguageCode.Croatian:
                    return "hr-HR";
                case LanguageId.LanguageCode.Czech:
                    return "cs-CZ";
                case LanguageId.LanguageCode.Danish:
                    return "da-DK";
                case LanguageId.LanguageCode.Divehi:
                    return "dv-MV";
                case LanguageId.LanguageCode.DutchBelgium:
                    return "nl-BE";
                case LanguageId.LanguageCode.DutchNetherlands:
                    return "nl-NL";
                case LanguageId.LanguageCode.Edo:
                    //there is no iso 639-1 code for edo
                case LanguageId.LanguageCode.EnglishAustralia:
                    return "en-AU";
                case LanguageId.LanguageCode.EnglishBelize:
                    return "en-BZ";
                case LanguageId.LanguageCode.EnglishCanada:
                    return "en-CA";
                case LanguageId.LanguageCode.EnglishCaribbean:
                    //the caribbean sea has many english speaking countires.
                    //we use the Dominican Republic
                    return "en-DO";
                case LanguageId.LanguageCode.EnglishHongKong:
                    return "en-HK";
                case LanguageId.LanguageCode.EnglishIndia:
                    return "en-IN";
                case LanguageId.LanguageCode.EnglishIndonesia:
                    return "en-ID";
                case LanguageId.LanguageCode.EnglishIreland:
                    return "en-IE";
                case LanguageId.LanguageCode.EnglishJamaica:
                    return "en-JM";
                case LanguageId.LanguageCode.EnglishMalaysia:
                    return "en-MY";
                case LanguageId.LanguageCode.EnglishNewZealand:
                    return "en-NZ";
                case LanguageId.LanguageCode.EnglishPhilippines:
                    return "en-PH";
                case LanguageId.LanguageCode.EnglishSingapore:
                    return "en-SG";
                case LanguageId.LanguageCode.EnglishSouthAfrica:
                    return "en-ZA";
                case LanguageId.LanguageCode.EnglishTrinidadAndTobago:
                    return "en-TT";
                case LanguageId.LanguageCode.EnglishUK:
                    return "en-UK";
                case LanguageId.LanguageCode.EnglishUS:
                    return "en-US";
                case LanguageId.LanguageCode.EnglishZimbabwe:
                    return "en-ZW";
                case LanguageId.LanguageCode.Estonian:
                    return "et-EE";
                case LanguageId.LanguageCode.Faeroese:
                    return "fo-FO";
                case LanguageId.LanguageCode.Farsi:
                    //there is no iso 639-1 code for farsi
                case LanguageId.LanguageCode.Filipino:
                    //there is no iso 639-1 code for filipino
                case LanguageId.LanguageCode.Finnish:
                    return "fi-FI";
                case LanguageId.LanguageCode.FrenchBelgium:
                    return "fr-BE";
                case LanguageId.LanguageCode.FrenchCameroon:
                    return "fr-CM";
                case LanguageId.LanguageCode.FrenchCanada:
                    return "fr-CA";
                case LanguageId.LanguageCode.FrenchCongoDRC:
                    return "fr-CD";
                case LanguageId.LanguageCode.FrenchCotedIvoire:
                    return "fr-CI";
                case LanguageId.LanguageCode.FrenchFrance:
                    return "fr-FR";
                case LanguageId.LanguageCode.FrenchHaiti:
                    return "fr-HT";
                case LanguageId.LanguageCode.FrenchLuxembourg:
                    return "fr-LU";
                case LanguageId.LanguageCode.FrenchMali:
                    return "fr-ML";
                case LanguageId.LanguageCode.FrenchMonaco:
                    return "fr-MC";
                case LanguageId.LanguageCode.FrenchMorocco:
                    return "fr-MA";
                case LanguageId.LanguageCode.FrenchReunion:
                    return "fr-RE";
                case LanguageId.LanguageCode.FrenchSenegal:
                    return "fr-SN";
                case LanguageId.LanguageCode.FrenchSwitzerland:
                    return "fr-CH";
                case LanguageId.LanguageCode.FrenchWestIndies:
                    //the western caribbean sea has many french speaking countires.
                    //we use the Dominican Republic
                    return "fr-DO";
                case LanguageId.LanguageCode.FrisianNetherlands:
                    return "fy-NL";
                case LanguageId.LanguageCode.Fulfulde:
                    //there is no iso 639-1 code for fulfulde
                case LanguageId.LanguageCode.FYROMacedonian:
                    return "mk-MK";
                case LanguageId.LanguageCode.GaelicIreland:
                    return "ga-IE";
                case LanguageId.LanguageCode.GaelicScotland:
                    return "gd-UK";
                case LanguageId.LanguageCode.Galician:
                    return "gl-ES";
                case LanguageId.LanguageCode.Georgian:
                    return "ka-GE";
                case LanguageId.LanguageCode.GermanAustria:
                    return "de-AT";
                case LanguageId.LanguageCode.GermanGermany:
                    return "de-DE";
                case LanguageId.LanguageCode.GermanLiechtenstein:
                    return "de-LI";
                case LanguageId.LanguageCode.GermanLuxembourg:
                    return "de-LU";
                case LanguageId.LanguageCode.GermanSwitzerland:
                    return "de-CH";
                case LanguageId.LanguageCode.Greek:
                    return "el-GR";
                case LanguageId.LanguageCode.Guarani:
                    return "gn-BR";
                case LanguageId.LanguageCode.Gujarati:
                    return "gu-IN";
                case LanguageId.LanguageCode.Hausa:
                    return "ha-NG";
                case LanguageId.LanguageCode.Hawaiian:
                    //there is no iso 639-1 language code for hawaiian
                case LanguageId.LanguageCode.Hebrew:
                    return "he-IL";
                case LanguageId.LanguageCode.Hindi:
                    return "hi-IN";
                case LanguageId.LanguageCode.Hungarian:
                    return "hu-HU";
                case LanguageId.LanguageCode.Ibibio:
                    //there is no iso 639-1 language code for ibibio
                case LanguageId.LanguageCode.Icelandic:
                    return "is-IS";
                case LanguageId.LanguageCode.Igbo:
                    //there is no iso 639-1 language code for ibibio
                case LanguageId.LanguageCode.Indonesian:
                    return "id-ID";
                case LanguageId.LanguageCode.Inuktitut:
                    return "iu-CA";
                case LanguageId.LanguageCode.ItalianItaly:
                    return "it-IT";
                case LanguageId.LanguageCode.ItalianSwitzerland:
                    return "it-CH";
                case LanguageId.LanguageCode.Japanese:
                    return "ja-JP";
                case LanguageId.LanguageCode.Kannada:
                    return "kn-ID";
                case LanguageId.LanguageCode.Kanuri:
                    //there is no iso 639-1 language code for kanuri
                case LanguageId.LanguageCode.Kashmiri:
                    return "ks-ID";
                case LanguageId.LanguageCode.KashmiriArabic:
                    return "ks-PK";
                case LanguageId.LanguageCode.Kazakh:
                    return "kk-KZ";
                case LanguageId.LanguageCode.Khmer:
                    //there is no iso 639-1 language code for khmer
                case LanguageId.LanguageCode.Konkani:
                    //there is no iso 639-1 language code for konkani
                case LanguageId.LanguageCode.Korean:
                    return "ko-KR";
                case LanguageId.LanguageCode.Kyrgyz:
                    return "ky-KG";
                case LanguageId.LanguageCode.Lao:
                    return "lo-LA";
                case LanguageId.LanguageCode.Latin:
                    return "la";
                case LanguageId.LanguageCode.Latvian:
                    return "lv-LV";
                case LanguageId.LanguageCode.Lithuanian:
                    return "lt-LT";
                case LanguageId.LanguageCode.Malay:
                    return "ms-MY";
                case LanguageId.LanguageCode.MalayBruneiDarussalam:
                    return "ms-BN";
                case LanguageId.LanguageCode.Malayalam:
                    return "ml-ID";
                case LanguageId.LanguageCode.Maltese:
                    return "mt-MT";
                case LanguageId.LanguageCode.Manipuri:
                    //there is no iso 639-1 language code for manipuri
                case LanguageId.LanguageCode.Maori:
                    return "mi-NZ";
                case LanguageId.LanguageCode.Marathi:
                    return "mr-ID";
                case LanguageId.LanguageCode.Mongolian:
                    return "mn-MN";
                case LanguageId.LanguageCode.MongolianMongolian:
                    return "mn-MN";
                case LanguageId.LanguageCode.Nepali:
                    return "ne-NP";
                case LanguageId.LanguageCode.NepaliIndia:
                    return "ne-ID";
                case LanguageId.LanguageCode.NorwegianBokmal:
                    return "nb-NO";
                    //also possible: no-NO
                case LanguageId.LanguageCode.NorwegianNynorsk:
                    return "nn-NO";
                    //also possible: no-NO
                case LanguageId.LanguageCode.Oriya:
                    return "or-ID";
                case LanguageId.LanguageCode.Oromo:
                    //there is no iso 639-1 language code for oromo
                case LanguageId.LanguageCode.Papiamentu:
                    //there is no iso 639-1 language code for papiamentu
                case LanguageId.LanguageCode.Pashto:
                    return "ps-PK";
                case LanguageId.LanguageCode.Polish:
                    return "pl-PL";
                case LanguageId.LanguageCode.PortugueseBrazil:
                    return "pt-BR";
                case LanguageId.LanguageCode.PortuguesePortugal:
                    return "pt-PT";
                case LanguageId.LanguageCode.Punjabi:
                    return "pa-ID";
                case LanguageId.LanguageCode.PunjabiPakistan:
                    return "pa-PK";
                case LanguageId.LanguageCode.QuechuaBolivia:
                    return "qu-BO";
                case LanguageId.LanguageCode.QuechuaEcuador:
                    return "qu-EC";
                case LanguageId.LanguageCode.QuechuaPeru:
                    return "qu-PE";
                case LanguageId.LanguageCode.RhaetoRomanic:
                    return "rm-CH";
                case LanguageId.LanguageCode.RomanianMoldova:
                    return "ro-MD";
                case LanguageId.LanguageCode.RomanianRomania:
                    return "ro-RO";
                case LanguageId.LanguageCode.RussianMoldova:
                    return "ru-MD";
                case LanguageId.LanguageCode.RussianRussia:
                    return "ru-RU";
                case LanguageId.LanguageCode.SamiLappish:
                    return "se-FI";
                case LanguageId.LanguageCode.Sanskrit:
                    return "sa-ID";
                case LanguageId.LanguageCode.Sepedi:
                    //there is no iso 639-1 language code for sepedi
                case LanguageId.LanguageCode.SerbianCyrillic:
                    return "sr-YU-cyrl";
                case LanguageId.LanguageCode.SerbianLatin:
                    return "sr-YU-latn";
                case LanguageId.LanguageCode.SindhiArabic:
                    return "sd-PK";
                case LanguageId.LanguageCode.SindhiDevanagari:
                    return "sd-ID";
                case LanguageId.LanguageCode.Sinhalese:
                    return "si-ID";
                case LanguageId.LanguageCode.Slovak:
                    return "sk-SK";
                case LanguageId.LanguageCode.Slovenian:
                    return "sl-SI";
                case LanguageId.LanguageCode.Somali:
                    return "so-SO";
                case LanguageId.LanguageCode.Sorbian:
                    //there is no iso 639-1 language code for sorbian
                case LanguageId.LanguageCode.SpanishArgentina:
                    return "es-AR";
                case LanguageId.LanguageCode.SpanishBolivia:
                    return "es-BO";
                case LanguageId.LanguageCode.SpanishChile:
                    return "es-CL";
                case LanguageId.LanguageCode.SpanishColombia:
                    return "es-CO";
                case LanguageId.LanguageCode.SpanishCostaRica:
                    return "es-CR";
                case LanguageId.LanguageCode.SpanishDominicanRepublic:
                    return "es-DO";
                case LanguageId.LanguageCode.SpanishEcuador:
                    return "es-EC";
                case LanguageId.LanguageCode.SpanishElSalvador:
                    return "es-SV";
                case LanguageId.LanguageCode.SpanishGuatemala:
                    return "es-GT";
                case LanguageId.LanguageCode.SpanishHonduras:
                    return "es-HN";
                case LanguageId.LanguageCode.SpanishMexico:
                    return "es-MX";
                case LanguageId.LanguageCode.SpanishNicaragua:
                    return "es-NI";
                case LanguageId.LanguageCode.SpanishPanama:
                    return "es-PA";
                case LanguageId.LanguageCode.SpanishParaguay:
                    return "es-PY";
                case LanguageId.LanguageCode.SpanishPeru:
                    return "es-PE";
                case LanguageId.LanguageCode.SpanishPuertoRico:
                    return "es-PR";
                case LanguageId.LanguageCode.SpanishSpainModernSort:
                    return "es-ES";
                case LanguageId.LanguageCode.SpanishSpainTraditionalSort:
                    return "es-ES";
                case LanguageId.LanguageCode.SpanishUruguay:
                    return "es-UY";
                case LanguageId.LanguageCode.SpanishVenezuela:
                    return "es-VE";
                case LanguageId.LanguageCode.Sutu:
                    //there is no iso 639-1 language code for sutu
                case LanguageId.LanguageCode.Swahili:
                    //Swahili is spoken in many east african countries, so we use tansania
                    return "sw-TZ";
                case LanguageId.LanguageCode.SwedishFinland:
                    return "sv-FI";
                case LanguageId.LanguageCode.SwedishSweden:
                    return "sv-SE";
                case LanguageId.LanguageCode.Syriac:
                    //there is no iso 639-1 language code for syriac
                case LanguageId.LanguageCode.Tajik:
                    return "tg-TJ";
                case LanguageId.LanguageCode.Tamazight:
                    //there is no iso 639-1 language code for tamazight
                case LanguageId.LanguageCode.TamazightLatin:
                    //there is no iso 639-1 language code for tamazight
                case LanguageId.LanguageCode.Tamil:
                    return "ta-ID";
                case LanguageId.LanguageCode.Tatar:
                    return "tt-RU";
                case LanguageId.LanguageCode.Telugu:
                    return "te-ID";
                //case LanguageId.LanguageCode.Thai:
                    
                //case LanguageId.LanguageCode.TibetanBhutan:
                    
                //case LanguageId.LanguageCode.TibetanPRC:
                    
                //case LanguageId.LanguageCode.TigrignaEritrea:
                    
                //case LanguageId.LanguageCode.TigrignaEthiopia:
                    
                //case LanguageId.LanguageCode.Tsonga:
                    
                //case LanguageId.LanguageCode.Tswana:
                    
                //case LanguageId.LanguageCode.Turkish:
                    
                //case LanguageId.LanguageCode.Turkmen:
                    
                //case LanguageId.LanguageCode.Ukrainian:
                    
                //case LanguageId.LanguageCode.Urdu:
                    
                //case LanguageId.LanguageCode.UzbekCyrillic:
                    
                //case LanguageId.LanguageCode.UzbekLatin:
                    
                //case LanguageId.LanguageCode.Venda:
                    
                //case LanguageId.LanguageCode.Vietnamese:
                    
                //case LanguageId.LanguageCode.Welsh:
                    
                //case LanguageId.LanguageCode.Xhosa:
                    
                //case LanguageId.LanguageCode.Yi:
                    
                //case LanguageId.LanguageCode.Yiddish:
                    
                //case LanguageId.LanguageCode.Yoruba:
                    
                //case LanguageId.LanguageCode.Zulu:
                   
                default:
                    return "en-US";
                    
            }


        }
    }
}
