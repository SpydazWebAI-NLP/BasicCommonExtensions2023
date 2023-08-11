Imports System.Drawing
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Web.Script.Serialization
Imports System.Windows.Forms
Imports System.Xml.Serialization
Imports SystemExtensions.AI_SDK_EXTENSIONS.Strings.Grammar
Imports SystemExtensions.AI_SDK_EXTENSIONS.Strings.Tokens

Namespace AI_SDK_EXTENSIONS

    Namespace Strings

        Namespace Grammar
            Public Module StopWords

                ''' <summary>
                ''' Removes StopWords from sentence
                ''' ARAB/ENG/DUTCH/FRENCH/SPANISH/ITALIAN
                ''' Hopefully leaving just relevant words in the user sentence
                ''' Currently Under Revision (takes too many words)
                ''' </summary>
                ''' <param name="Userinput"></param>
                ''' <returns></returns>
                <Runtime.CompilerServices.Extension()>
                Public Function RemoveStopWords(ByRef Userinput As String) As String
                    ' Userinput = LCase(Userinput).Replace("the", "r")
                    For Each item In StopWordsENG
                        Userinput = LCase(Userinput).Replace(item, "")
                    Next
                    For Each item In StopWordsArab
                        Userinput = Userinput.Replace(item, "")
                    Next
                    For Each item In StopWordsDutch
                        Userinput = Userinput.Replace(item, "")
                    Next
                    For Each item In StopWordsFrench
                        Userinput = Userinput.Replace(item, "")
                    Next
                    For Each item In StopWordsItalian
                        Userinput = Userinput.Replace(item, "")
                    Next
                    For Each item In StopWordsSpanish
                        Userinput = Userinput.Replace(item, "")
                    Next
                    Return Userinput
                End Function

#Region "Words"

                Public StopWordsENG() As String = {"a", "as", "able", "about", "above", "according", "accordingly", "across", "actually", "after", "afterwards", "again", "against", "aint",
    "all", "allow", "allows", "almost", "alone", "along", "already", "also", "although", "always", "am", "among", "amongst", "an", "and", "another", "any",
    "anybody", "anyhow", "anyone", "anything", "anyway", "anyways", "anywhere", "apart", "appear", "appreciate", "appropriate", "are", "arent", "around",
    "as", "aside", "ask", "asking", "associated", "at", "available", "away", "awfully", "b", "be", "became", "because", "become", "becomes", "becoming",
    "been", "before", "beforehand", "behind", "being", "believe", "below", "beside", "besides", "best", "better", "between", "beyond", "both", "brief",
    "but", "by", "c", "cmon", "cs", "came", "can", "cant", "cannot", "cant", "cause", "causes", "certain", "certainly", "changes", "clearly", "co", "com",
    "come", "comes", "concerning", "consequently", "consider", "considering", "contain", "containing", "contains", "corresponding", "could", "couldnt",
    "course", "currently", "d", "definitely", "described", "despite", "did", "didnt", "different", "do", "does", "doesnt", "doing", "dont", "done", "down",
    "downwards", "during", "e", "each", "edu", "eg", "eight", "either", "else", "elsewhere", "enough", "entirely", "especially", "et", "etc", "even", "ever",
    "every", "everybody", "everyone", "everything", "everywhere", "ex", "exactly", "example", "except", "f", "far", "few", "fifth", "first", "five", "followed",
    "following", "follows", "for", "former", "formerly", "forth", "four", "from", "further", "furthermore", "g", "get", "gets", "getting", "given", "gives",
    "go", "goes", "going", "gone", "got", "gotten", "greetings", "h", "had", "hadnt", "happens", "hardly", "has", "hasnt", "have", "havent", "having", "he",
    "hes", "hello", "help", "hence", "her", "here", "heres", "hereafter", "hereby", "herein", "hereupon", "hers", "herself", "hi", "him", "himself", "his",
    "hither", "hopefully", "how", "howbeit", "however", "i", "id", "ill", "im", "ive", "ie", "if", "ignored", "immediate", "in", "inasmuch", "inc", "indeed",
    "indicate", "indicated", "indicates", "inner", "insofar", "instead", "into", "inward", "is", "isnt", "it", "itd", "itll", "its", "its", "itself", "j",
    "just", "k", "keep", "keeps", "kept", "know", "known", "knows", "l", "last", "lately", "later", "latter", "latterly", "least", "less", "lest", "let", "lets",
    "like", "liked", "likely", "little", "look", "looking", "looks", "ltd", "m", "mainly", "many", "may", "maybe", "me", "mean", "meanwhile", "merely", "might",
    "more", "moreover", "most", "mostly", "much", "must", "my", "myself", "n", "name", "namely", "nd", "near", "nearly", "necessary", "need", "needs", "neither",
    "never", "nevertheless", "new", "next", "nine", "no", "nobody", "non", "none", "noone", "nor", "normally", "not", "nothing", "novel", "now", "nowhere", "o",
    "obviously", "of", "off", "often", "oh", "ok", "okay", "old", "on", "once", "one", "ones", "only", "onto", "or", "other", "others", "otherwise", "ought", "our",
    "ours", "ourselves", "out", "outside", "over", "overall", "own", "p", "particular", "particularly", "per", "perhaps", "placed", "please", "plus", "possible",
    "presumably", "probably", "provides", "q", "que", "quite", "qv", "r", "rather", "rd", "re", "really", "reasonably", "regarding", "regardless", "regards",
    "relatively", "respectively", "right", "s", "said", "same", "saw", "say", "saying", "says", "second", "secondly", "see", "seeing", "seem", "seemed", "seeming",
    "seems", "seen", "self", "selves", "sensible", "sent", "serious", "seriously", "seven", "several", "shall", "she", "should", "shouldnt", "since", "six", "so",
    "some", "somebody", "somehow", "someone", "something", "sometime", "sometimes", "somewhat", "somewhere", "soon", "sorry", "specified", "specify", "specifying",
    "still", "sub", "such", "sup", "sure", "t", "ts", "take", "taken", "tell", "tends", "th", "than", "thank", "thanks", "thanx", "that", "thats", "thats", "the",
    "their", "theirs", "them", "themselves", "then", "thence", "there", "theres", "thereafter", "thereby", "therefore", "therein", "theres", "thereupon",
    "these", "they", "theyd", "theyll", "theyre", "theyve", "think", "third", "this", "thorough", "thoroughly", "those", "though", "three", "through",
    "throughout", "thru", "thus", "to", "together", "too", "took", "toward", "towards", "tried", "tries", "truly", "try", "trying", "twice", "two", "u", "un",
    "under", "unfortunately", "unless", "unlikely", "until", "unto", "up", "upon", "us", "use", "used", "useful", "uses", "using", "usually", "uucp", "v",
    "value", "various", "very", "via", "viz", "vs", "w", "want", "wants", "was", "wasnt", "way", "we", "wed", "well", "were", "weve", "welcome", "well",
    "went", "were", "werent", "what", "whats", "whatever", "when", "whence", "whenever", "where", "wheres", "whereafter", "whereas", "whereby", "wherein",
    "whereupon", "wherever", "whether", "which", "while", "whither", "who", "whos", "whoever", "whole", "whom", "whose", "why", "will", "willing", "wish",
    "with", "within", "without", "wont", "wonder", "would", "wouldnt", "x", "y", "yes", "yet", "you", "youd", "youll", "youre", "youve", "your", "yours",
    "yourself", "yourselves", "youll", "z", "zero"}

                Public StopWordsArab() As String = {"،", "آض", "آمينَ", "آه",
            "آهاً", "آي", "أ", "أب", "أجل", "أجمع", "أخ", "أخذ", "أصبح", "أضحى", "أقبل",
            "أقل", "أكثر", "ألا", "أم", "أما", "أمامك", "أمامكَ", "أمسى", "أمّا", "أن", "أنا", "أنت",
            "أنتم", "أنتما", "أنتن", "أنتِ", "أنشأ", "أنّى", "أو", "أوشك", "أولئك", "أولئكم", "أولاء",
            "أولالك", "أوّهْ", "أي", "أيا", "أين", "أينما", "أيّ", "أَنَّ", "أََيُّ", "أُفٍّ", "إذ", "إذا", "إذاً",
            "إذما", "إذن", "إلى", "إليكم", "إليكما", "إليكنّ", "إليكَ", "إلَيْكَ", "إلّا", "إمّا", "إن", "إنّما",
            "إي", "إياك", "إياكم", "إياكما", "إياكن", "إيانا", "إياه", "إياها", "إياهم", "إياهما", "إياهن",
            "إياي", "إيهٍ", "إِنَّ", "ا", "ابتدأ", "اثر", "اجل", "احد", "اخرى", "اخلولق", "اذا", "اربعة", "ارتدّ",
            "استحال", "اطار", "اعادة", "اعلنت", "اف", "اكثر", "اكد", "الألاء", "الألى", "الا", "الاخيرة", "الان", "الاول",
            "الاولى", "التى", "التي", "الثاني", "الثانية", "الذاتي", "الذى", "الذي", "الذين", "السابق", "الف", "اللائي",
            "اللاتي", "اللتان", "اللتيا", "اللتين", "اللذان", "اللذين", "اللواتي", "الماضي", "المقبل", "الوقت", "الى",
            "اليوم", "اما", "امام", "امس", "ان", "انبرى", "انقلب", "انه", "انها", "او", "اول", "اي", "ايار", "ايام",
            "ايضا", "ب", "بات", "باسم", "بان", "بخٍ", "برس", "بسبب", "بسّ", "بشكل", "بضع", "بطآن", "بعد", "بعض", "بك",
            "بكم", "بكما", "بكن", "بل", "بلى", "بما", "بماذا", "بمن", "بن", "بنا", "به", "بها", "بي", "بيد", "بين",
            "بَسْ", "بَلْهَ", "بِئْسَ", "تانِ", "تانِك", "تبدّل", "تجاه", "تحوّل", "تلقاء", "تلك", "تلكم", "تلكما", "تم", "تينك",
            "تَيْنِ", "تِه", "تِي", "ثلاثة", "ثم", "ثمّ", "ثمّة", "ثُمَّ", "جعل", "جلل", "جميع", "جير", "حار", "حاشا", "حاليا",
            "حاي", "حتى", "حرى", "حسب", "حم", "حوالى", "حول", "حيث", "حيثما", "حين", "حيَّ", "حَبَّذَا", "حَتَّى", "حَذارِ", "خلا",
            "خلال", "دون", "دونك", "ذا", "ذات", "ذاك", "ذانك", "ذانِ", "ذلك", "ذلكم", "ذلكما", "ذلكن", "ذو", "ذوا", "ذواتا", "ذواتي", "ذيت", "ذينك", "ذَيْنِ", "ذِه", "ذِي", "راح", "رجع", "رويدك", "ريث", "رُبَّ", "زيارة", "سبحان", "سرعان", "سنة", "سنوات", "سوف", "سوى", "سَاءَ", "سَاءَمَا", "شبه", "شخصا", "شرع", "شَتَّانَ", "صار", "صباح", "صفر", "صهٍ", "صهْ", "ضد", "ضمن", "طاق", "طالما", "طفق", "طَق", "ظلّ", "عاد", "عام", "عاما", "عامة", "عدا", "عدة", "عدد", "عدم", "عسى", "عشر", "عشرة", "علق", "على", "عليك", "عليه", "عليها", "علًّ", "عن", "عند", "عندما", "عوض", "عين", "عَدَسْ", "عَمَّا", "غدا", "غير", "ـ", "ف", "فان", "فلان", "فو", "فى", "في", "فيم", "فيما", "فيه", "فيها", "قال", "قام", "قبل", "قد", "قطّ", "قلما", "قوة", "كأنّما", "كأين", "كأيّ", "كأيّن", "كاد", "كان", "كانت", "كذا", "كذلك", "كرب", "كل", "كلا", "كلاهما", "كلتا", "كلم", "كليكما", "كليهما", "كلّما", "كلَّا", "كم", "كما", "كي", "كيت", "كيف", "كيفما", "كَأَنَّ", "كِخ", "لئن", "لا", "لات", "لاسيما", "لدن", "لدى", "لعمر", "لقاء", "لك", "لكم", "لكما", "لكن", "لكنَّما", "لكي", "لكيلا", "للامم", "لم", "لما", "لمّا", "لن", "لنا", "له", "لها", "لو", "لوكالة", "لولا", "لوما", "لي", "لَسْتَ", "لَسْتُ", "لَسْتُم", "لَسْتُمَا", "لَسْتُنَّ", "لَسْتِ", "لَسْنَ", "لَعَلَّ", "لَكِنَّ", "لَيْتَ", "لَيْسَ", "لَيْسَا", "لَيْسَتَا", "لَيْسَتْ", "لَيْسُوا", "لَِسْنَا", "ما", "ماانفك", "مابرح", "مادام", "ماذا", "مازال", "مافتئ", "مايو", "متى", "مثل", "مذ", "مساء", "مع", "معاذ", "مقابل", "مكانكم", "مكانكما", "مكانكنّ", "مكانَك", "مليار", "مليون", "مما", "ممن", "من", "منذ", "منها", "مه", "مهما", "مَنْ", "مِن", "نحن", "نحو", "نعم", "نفس", "نفسه", "نهاية", "نَخْ", "نِعِمّا", "نِعْمَ", "ها", "هاؤم", "هاكَ", "هاهنا", "هبّ", "هذا", "هذه", "هكذا", "هل", "هلمَّ", "هلّا", "هم", "هما", "هن", "هنا", "هناك", "هنالك", "هو", "هي", "هيا", "هيت", "هيّا", "هَؤلاء", "هَاتانِ", "هَاتَيْنِ", "هَاتِه", "هَاتِي", "هَجْ", "هَذا", "هَذانِ", "هَذَيْنِ", "هَذِه", "هَذِي", "هَيْهَاتَ", "و", "و6", "وا", "واحد", "واضاف", "واضافت", "واكد", "وان", "واهاً", "واوضح", "وراءَك", "وفي", "وقال", "وقالت", "وقد", "وقف", "وكان", "وكانت", "ولا", "ولم",
            "ومن", "وهو", "وهي", "ويكأنّ", "وَيْ", "وُشْكَانََ", "يكون", "يمكن", "يوم", "ّأيّان"}

                Public StopWordsItalian() As String = {"IE", "a", "abbastanza", "abbia", "abbiamo", "abbiano", "abbiate", "accidenti", "ad", "adesso", "affinche", "agl", "agli",
    "ahime", "ahimè", "ai", "al", "alcuna", "alcuni", "alcuno", "all", "alla", "alle", "allo", "allora", "altri", "altrimenti", "altro",
    "altrove", "altrui", "anche", "ancora", "anni", "anno", "ansa", "anticipo", "assai", "attesa", "attraverso", "avanti", "avemmo",
    "avendo", "avente", "aver", "avere", "averlo", "avesse", "avessero", "avessi", "avessimo", "aveste", "avesti", "avete", "aveva",
    "avevamo", "avevano", "avevate", "avevi", "avevo", "avrai", "avranno", "avrebbe", "avrebbero", "avrei", "avremmo", "avremo",
    "avreste", "avresti", "avrete", "avrà", "avrò", "avuta", "avute", "avuti", "avuto", "basta", "bene", "benissimo", "berlusconi",
    "brava", "bravo", "c", "casa", "caso", "cento", "certa", "certe", "certi", "certo", "che", "chi", "chicchessia", "chiunque", "ci",
    "ciascuna", "ciascuno", "cima", "cio", "cioe", "cioè", "circa", "citta", "città", "ciò", "co", "codesta", "codesti", "codesto",
    "cogli", "coi", "col", "colei", "coll", "coloro", "colui", "come", "cominci", "comunque", "con", "concernente", "conciliarsi",
    "conclusione", "consiglio", "contro", "cortesia", "cos", "cosa", "cosi", "così", "cui", "d", "da", "dagl", "dagli", "dai", "dal",
    "dall", "dalla", "dalle", "dallo", "dappertutto", "davanti", "degl", "degli", "dei", "del", "dell", "della", "delle", "dello",
    "dentro", "detto", "deve", "di", "dice", "dietro", "dire", "dirimpetto", "diventa", "diventare", "diventato", "dopo", "dov", "dove",
    "dovra", "dovrà", "dovunque", "due", "dunque", "durante", "e", "ebbe", "ebbero", "ebbi", "ecc", "ecco", "ed", "effettivamente", "egli",
    "ella", "entrambi", "eppure", "era", "erano", "eravamo", "eravate", "eri", "ero", "esempio", "esse", "essendo", "esser", "essere",
    "essi", "ex", "fa", "faccia", "facciamo", "facciano", "facciate", "faccio", "facemmo", "facendo", "facesse", "facessero", "facessi",
    "facessimo", "faceste", "facesti", "faceva", "facevamo", "facevano", "facevate", "facevi", "facevo", "fai", "fanno", "farai",
    "faranno", "fare", "farebbe", "farebbero", "farei", "faremmo", "faremo", "fareste", "faresti", "farete", "farà", "farò", "fatto",
    "favore", "fece", "fecero", "feci", "fin", "finalmente", "finche", "fine", "fino", "forse", "forza", "fosse", "fossero", "fossi",
    "fossimo", "foste", "fosti", "fra", "frattempo", "fu", "fui", "fummo", "fuori", "furono", "futuro", "generale", "gia", "giacche",
    "giorni", "giorno", "già", "gli", "gliela", "gliele", "glieli", "glielo", "gliene", "governo", "grande", "grazie", "gruppo", "ha",
    "haha", "hai", "hanno", "ho", "i", "ieri", "il", "improvviso", "in", "inc", "infatti", "inoltre", "insieme", "intanto", "intorno",
    "invece", "io", "l", "la", "lasciato", "lato", "lavoro", "le", "lei", "li", "lo", "lontano", "loro", "lui", "lungo", "luogo", "là",
    "ma", "macche", "magari", "maggior", "mai", "male", "malgrado", "malissimo", "mancanza", "marche", "me", "medesimo", "mediante",
    "meglio", "meno", "mentre", "mesi", "mezzo", "mi", "mia", "mie", "miei", "mila", "miliardi", "milioni", "minimi", "ministro",
    "mio", "modo", "molti", "moltissimo", "molto", "momento", "mondo", "mosto", "nazionale", "ne", "negl", "negli", "nei", "nel",
    "nell", "nella", "nelle", "nello", "nemmeno", "neppure", "nessun", "nessuna", "nessuno", "niente", "no", "noi", "non", "nondimeno",
    "nonostante", "nonsia", "nostra", "nostre", "nostri", "nostro", "novanta", "nove", "nulla", "nuovo", "o", "od", "oggi", "ogni",
    "ognuna", "ognuno", "oltre", "oppure", "ora", "ore", "osi", "ossia", "ottanta", "otto", "paese", "parecchi", "parecchie",
    "parecchio", "parte", "partendo", "peccato", "peggio", "per", "perche", "perchè", "perché", "percio", "perciò", "perfino", "pero",
    "persino", "persone", "però", "piedi", "pieno", "piglia", "piu", "piuttosto", "più", "po", "pochissimo", "poco", "poi", "poiche",
    "possa", "possedere", "posteriore", "posto", "potrebbe", "preferibilmente", "presa", "press", "prima", "primo", "principalmente",
    "probabilmente", "proprio", "puo", "pure", "purtroppo", "può", "qualche", "qualcosa", "qualcuna", "qualcuno", "quale", "quali",
    "qualunque", "quando", "quanta", "quante", "quanti", "quanto", "quantunque", "quasi", "quattro", "quel", "quella", "quelle",
    "quelli", "quello", "quest", "questa", "queste", "questi", "questo", "qui", "quindi", "realmente", "recente", "recentemente",
    "registrazione", "relativo", "riecco", "salvo", "sara", "sarai", "saranno", "sarebbe", "sarebbero", "sarei", "saremmo", "saremo",
    "sareste", "saresti", "sarete", "sarà", "sarò", "scola", "scopo", "scorso", "se", "secondo", "seguente", "seguito", "sei", "sembra",
    "sembrare", "sembrato", "sembri", "sempre", "senza", "sette", "si", "sia", "siamo", "siano", "siate", "siete", "sig", "solito",
    "solo", "soltanto", "sono", "sopra", "sotto", "spesso", "srl", "sta", "stai", "stando", "stanno", "starai", "staranno", "starebbe",
    "starebbero", "starei", "staremmo", "staremo", "stareste", "staresti", "starete", "starà", "starò", "stata", "state", "stati",
    "stato", "stava", "stavamo", "stavano", "stavate", "stavi", "stavo", "stemmo", "stessa", "stesse", "stessero", "stessi", "stessimo",
    "stesso", "steste", "stesti", "stette", "stettero", "stetti", "stia", "stiamo", "stiano", "stiate", "sto", "su", "sua", "subito",
    "successivamente", "successivo", "sue", "sugl", "sugli", "sui", "sul", "sull", "sulla", "sulle", "sullo", "suo", "suoi", "tale",
    "tali", "talvolta", "tanto", "te", "tempo", "ti", "titolo", "torino", "tra", "tranne", "tre", "trenta", "troppo", "trovato", "tu",
    "tua", "tue", "tuo", "tuoi", "tutta", "tuttavia", "tutte", "tutti", "tutto", "uguali", "ulteriore", "ultimo", "un", "una", "uno",
    "uomo", "va", "vale", "vari", "varia", "varie", "vario", "verso", "vi", "via", "vicino", "visto", "vita", "voi", "volta", "volte",
    "vostra", "vostre", "vostri", "vostro", "è"}

                Public StopWordsSpanish() As String = {"a", "actualmente", "acuerdo", "adelante", "ademas", "además", "adrede", "afirmó", "agregó", "ahi", "ahora",
    "ahí", "al", "algo", "alguna", "algunas", "alguno", "algunos", "algún", "alli", "allí", "alrededor", "ambos", "ampleamos",
    "antano", "antaño", "ante", "anterior", "antes", "apenas", "aproximadamente", "aquel", "aquella", "aquellas", "aquello",
    "aquellos", "aqui", "aquél", "aquélla", "aquéllas", "aquéllos", "aquí", "arriba", "arribaabajo", "aseguró", "asi", "así",
    "atras", "aun", "aunque", "ayer", "añadió", "aún", "b", "bajo", "bastante", "bien", "breve", "buen", "buena", "buenas", "bueno",
    "buenos", "c", "cada", "casi", "cerca", "cierta", "ciertas", "cierto", "ciertos", "cinco", "claro", "comentó", "como", "con",
    "conmigo", "conocer", "conseguimos", "conseguir", "considera", "consideró", "consigo", "consigue", "consiguen", "consigues",
    "contigo", "contra", "cosas", "creo", "cual", "cuales", "cualquier", "cuando", "cuanta", "cuantas", "cuanto", "cuantos", "cuatro",
    "cuenta", "cuál", "cuáles", "cuándo", "cuánta", "cuántas", "cuánto", "cuántos", "cómo", "d", "da", "dado", "dan", "dar", "de",
    "debajo", "debe", "deben", "debido", "decir", "dejó", "del", "delante", "demasiado", "demás", "dentro", "deprisa", "desde",
    "despacio", "despues", "después", "detras", "detrás", "dia", "dias", "dice", "dicen", "dicho", "dieron", "diferente", "diferentes",
    "dijeron", "dijo", "dio", "donde", "dos", "durante", "día", "días", "dónde", "e", "ejemplo", "el", "ella", "ellas", "ello", "ellos",
    "embargo", "empleais", "emplean", "emplear", "empleas", "empleo", "en", "encima", "encuentra", "enfrente", "enseguida", "entonces",
    "entre", "era", "eramos", "eran", "eras", "eres", "es", "esa", "esas", "ese", "eso", "esos", "esta", "estaba", "estaban", "estado",
    "estados", "estais", "estamos", "estan", "estar", "estará", "estas", "este", "esto", "estos", "estoy", "estuvo", "está", "están", "ex",
    "excepto", "existe", "existen", "explicó", "expresó", "f", "fin", "final", "fue", "fuera", "fueron", "fui", "fuimos", "g", "general",
    "gran", "grandes", "gueno", "h", "ha", "haber", "habia", "habla", "hablan", "habrá", "había", "habían", "hace", "haceis", "hacemos",
    "hacen", "hacer", "hacerlo", "haces", "hacia", "haciendo", "hago", "han", "hasta", "hay", "haya", "he", "hecho", "hemos", "hicieron",
    "hizo", "horas", "hoy", "hubo", "i", "igual", "incluso", "indicó", "informo", "informó", "intenta", "intentais", "intentamos", "intentan",
    "intentar", "intentas", "intento", "ir", "j", "junto", "k", "l", "la", "lado", "largo", "las", "le", "lejos", "les", "llegó", "lleva",
    "llevar", "lo", "los", "luego", "lugar", "m", "mal", "manera", "manifestó", "mas", "mayor", "me", "mediante", "medio", "mejor", "mencionó",
    "menos", "menudo", "mi", "mia", "mias", "mientras", "mio", "mios", "mis", "misma", "mismas", "mismo", "mismos", "modo", "momento", "mucha",
    "muchas", "mucho", "muchos", "muy", "más", "mí", "mía", "mías", "mío", "míos", "n", "nada", "nadie", "ni", "ninguna", "ningunas", "ninguno",
    "ningunos", "ningún", "no", "nos", "nosotras", "nosotros", "nuestra", "nuestras", "nuestro", "nuestros", "nueva", "nuevas", "nuevo",
    "nuevos", "nunca", "o", "ocho", "os", "otra", "otras", "otro", "otros", "p", "pais", "para", "parece", "parte", "partir", "pasada",
    "pasado", "paìs", "peor", "pero", "pesar", "poca", "pocas", "poco", "pocos", "podeis", "podemos", "poder", "podria", "podriais",
    "podriamos", "podrian", "podrias", "podrá", "podrán", "podría", "podrían", "poner", "por", "porque", "posible", "primer", "primera",
    "primero", "primeros", "principalmente", "pronto", "propia", "propias", "propio", "propios", "proximo", "próximo", "próximos", "pudo",
    "pueda", "puede", "pueden", "puedo", "pues", "q", "qeu", "que", "quedó", "queremos", "quien", "quienes", "quiere", "quiza", "quizas",
    "quizá", "quizás", "quién", "quiénes", "qué", "r", "raras", "realizado", "realizar", "realizó", "repente", "respecto", "s", "sabe",
    "sabeis", "sabemos", "saben", "saber", "sabes", "salvo", "se", "sea", "sean", "segun", "segunda", "segundo", "según", "seis", "ser",
    "sera", "será", "serán", "sería", "señaló", "si", "sido", "siempre", "siendo", "siete", "sigue", "siguiente", "sin", "sino", "sobre",
    "sois", "sola", "solamente", "solas", "solo", "solos", "somos", "son", "soy", "soyos", "su", "supuesto", "sus", "suya", "suyas", "suyo",
    "sé", "sí", "sólo", "t", "tal", "tambien", "también", "tampoco", "tan", "tanto", "tarde", "te", "temprano", "tendrá", "tendrán", "teneis",
    "tenemos", "tener", "tenga", "tengo", "tenido", "tenía", "tercera", "ti", "tiempo", "tiene", "tienen", "toda", "todas", "todavia",
    "todavía", "todo", "todos", "total", "trabaja", "trabajais", "trabajamos", "trabajan", "trabajar", "trabajas", "trabajo", "tras",
    "trata", "través", "tres", "tu", "tus", "tuvo", "tuya", "tuyas", "tuyo", "tuyos", "tú", "u", "ultimo", "un", "una", "unas", "uno", "unos",
    "usa", "usais", "usamos", "usan", "usar", "usas", "uso", "usted", "ustedes", "v", "va", "vais", "valor", "vamos", "van", "varias", "varios",
    "vaya", "veces", "ver", "verdad", "verdadera", "verdadero", "vez", "vosotras", "vosotros", "voy", "vuestra", "vuestras", "vuestro",
    "vuestros", "w", "x", "y", "ya", "yo", "z", "él", "ésa", "ésas", "ése", "ésos", "ésta", "éstas", "éste", "éstos", "última", "últimas",
    "último", "últimos"}

                Public StopWordsFrench() As String = {"a", "abord", "absolument", "afin", "ah", "ai", "aie", "ailleurs", "ainsi", "ait", "allaient", "allo", "allons",
    "allô", "alors", "anterieur", "anterieure", "anterieures", "apres", "après", "as", "assez", "attendu", "au", "aucun", "aucune",
    "aujourd", "aujourd'hui", "aupres", "auquel", "aura", "auraient", "aurait", "auront", "aussi", "autre", "autrefois", "autrement",
    "autres", "autrui", "aux", "auxquelles", "auxquels", "avaient", "avais", "avait", "avant", "avec", "avoir", "avons", "ayant", "b",
    "bah", "bas", "basee", "bat", "beau", "beaucoup", "bien", "bigre", "boum", "bravo", "brrr", "c", "car", "ce", "ceci", "cela", "celle",
    "celle-ci", "celle-là", "celles", "celles-ci", "celles-là", "celui", "celui-ci", "celui-là", "cent", "cependant", "certain",
    "certaine", "certaines", "certains", "certes", "ces", "cet", "cette", "ceux", "ceux-ci", "ceux-là", "chacun", "chacune", "chaque",
    "cher", "chers", "chez", "chiche", "chut", "chère", "chères", "ci", "cinq", "cinquantaine", "cinquante", "cinquantième", "cinquième",
    "clac", "clic", "combien", "comme", "comment", "comparable", "comparables", "compris", "concernant", "contre", "couic", "crac", "d",
    "da", "dans", "de", "debout", "dedans", "dehors", "deja", "delà", "depuis", "dernier", "derniere", "derriere", "derrière", "des",
    "desormais", "desquelles", "desquels", "dessous", "dessus", "deux", "deuxième", "deuxièmement", "devant", "devers", "devra",
    "different", "differentes", "differents", "différent", "différente", "différentes", "différents", "dire", "directe", "directement",
    "dit", "dite", "dits", "divers", "diverse", "diverses", "dix", "dix-huit", "dix-neuf", "dix-sept", "dixième", "doit", "doivent", "donc",
    "dont", "douze", "douzième", "dring", "du", "duquel", "durant", "dès", "désormais", "e", "effet", "egale", "egalement", "egales", "eh",
    "elle", "elle-même", "elles", "elles-mêmes", "en", "encore", "enfin", "entre", "envers", "environ", "es", "est", "et", "etant", "etc",
    "etre", "eu", "euh", "eux", "eux-mêmes", "exactement", "excepté", "extenso", "exterieur", "f", "fais", "faisaient", "faisant", "fait",
    "façon", "feront", "fi", "flac", "floc", "font", "g", "gens", "h", "ha", "hein", "hem", "hep", "hi", "ho", "holà", "hop", "hormis", "hors",
    "hou", "houp", "hue", "hui", "huit", "huitième", "hum", "hurrah", "hé", "hélas", "i", "il", "ils", "importe", "j", "je", "jusqu", "jusque",
    "juste", "k", "l", "la", "laisser", "laquelle", "las", "le", "lequel", "les", "lesquelles", "lesquels", "leur", "leurs", "longtemps",
    "lors", "lorsque", "lui", "lui-meme", "lui-même", "là", "lès", "m", "ma", "maint", "maintenant", "mais", "malgre", "malgré", "maximale",
    "me", "meme", "memes", "merci", "mes", "mien", "mienne", "miennes", "miens", "mille", "mince", "minimale", "moi", "moi-meme", "moi-même",
    "moindres", "moins", "mon", "moyennant", "multiple", "multiples", "même", "mêmes", "n", "na", "naturel", "naturelle", "naturelles", "ne",
    "neanmoins", "necessaire", "necessairement", "neuf", "neuvième", "ni", "nombreuses", "nombreux", "non", "nos", "notamment", "notre",
    "nous", "nous-mêmes", "nouveau", "nul", "néanmoins", "nôtre", "nôtres", "o", "oh", "ohé", "ollé", "olé", "on", "ont", "onze", "onzième",
    "ore", "ou", "ouf", "ouias", "oust", "ouste", "outre", "ouvert", "ouverte", "ouverts", "o|", "où", "p", "paf", "pan", "par", "parce",
    "parfois", "parle", "parlent", "parler", "parmi", "parseme", "partant", "particulier", "particulière", "particulièrement", "pas",
    "passé", "pendant", "pense", "permet", "personne", "peu", "peut", "peuvent", "peux", "pff", "pfft", "pfut", "pif", "pire", "plein",
    "plouf", "plus", "plusieurs", "plutôt", "possessif", "possessifs", "possible", "possibles", "pouah", "pour", "pourquoi", "pourrais",
    "pourrait", "pouvait", "prealable", "precisement", "premier", "première", "premièrement", "pres", "probable", "probante",
    "procedant", "proche", "près", "psitt", "pu", "puis", "puisque", "pur", "pure", "q", "qu", "quand", "quant", "quant-à-soi", "quanta",
    "quarante", "quatorze", "quatre", "quatre-vingt", "quatrième", "quatrièmement", "que", "quel", "quelconque", "quelle", "quelles",
    "quelqu'un", "quelque", "quelques", "quels", "qui", "quiconque", "quinze", "quoi", "quoique", "r", "rare", "rarement", "rares",
    "relative", "relativement", "remarquable", "rend", "rendre", "restant", "reste", "restent", "restrictif", "retour", "revoici",
    "revoilà", "rien", "s", "sa", "sacrebleu", "sait", "sans", "sapristi", "sauf", "se", "sein", "seize", "selon", "semblable", "semblaient",
    "semble", "semblent", "sent", "sept", "septième", "sera", "seraient", "serait", "seront", "ses", "seul", "seule", "seulement", "si",
    "sien", "sienne", "siennes", "siens", "sinon", "six", "sixième", "soi", "soi-même", "soit", "soixante", "son", "sont", "sous", "souvent",
    "specifique", "specifiques", "speculatif", "stop", "strictement", "subtiles", "suffisant", "suffisante", "suffit", "suis", "suit",
    "suivant", "suivante", "suivantes", "suivants", "suivre", "superpose", "sur", "surtout", "t", "ta", "tac", "tant", "tardive", "te",
    "tel", "telle", "tellement", "telles", "tels", "tenant", "tend", "tenir", "tente", "tes", "tic", "tien", "tienne", "tiennes", "tiens",
    "toc", "toi", "toi-même", "ton", "touchant", "toujours", "tous", "tout", "toute", "toutefois", "toutes", "treize", "trente", "tres",
    "trois", "troisième", "troisièmement", "trop", "très", "tsoin", "tsouin", "tu", "té", "u", "un", "une", "unes", "uniformement", "unique",
    "uniques", "uns", "v", "va", "vais", "vas", "vers", "via", "vif", "vifs", "vingt", "vivat", "vive", "vives", "vlan", "voici", "voilà",
    "vont", "vos", "votre", "vous", "vous-mêmes", "vu", "vé", "vôtre", "vôtres", "w", "x", "y", "z", "zut", "à", "â", "ça", "ès", "étaient",
    "étais", "était", "étant", "été", "être", "ô"}

                Public StopWordsDutch() As String = {"aan", "achte", "achter", "af", "al", "alle", "alleen", "alles", "als", "ander", "anders", "beetje",
    "behalve", "beide", "beiden", "ben", "beneden", "bent", "bij", "bijna", "bijv", "blijkbaar", "blijken", "boven", "bv",
    "daar", "daardoor", "daarin", "daarna", "daarom", "daaruit", "dan", "dat", "de", "deden", "deed", "derde", "derhalve", "dertig",
    "deze", "dhr", "die", "dit", "doe", "doen", "doet", "door", "drie", "duizend", "echter", "een", "eens", "eerst", "eerste", "eigen",
    "eigenlijk", "elk", "elke", "en", "enige", "er", "erg", "ergens", "etc", "etcetera", "even", "geen", "genoeg", "geweest", "haar",
    "haarzelf", "had", "hadden", "heb", "hebben", "hebt", "hedden", "heeft", "heel", "hem", "hemzelf", "hen", "het", "hetzelfde",
    "hier", "hierin", "hierna", "hierom", "hij", "hijzelf", "hoe", "honderd", "hun", "ieder", "iedere", "iedereen", "iemand", "iets",
    "ik", "in", "inderdaad", "intussen", "is", "ja", "je", "jij", "jijzelf", "jou", "jouw", "jullie", "kan", "kon", "konden", "kun",
    "kunnen", "kunt", "laatst", "later", "lijken", "lijkt", "maak", "maakt", "maakte", "maakten", "maar", "mag", "maken", "me", "meer",
    "meest", "meestal", "men", "met", "mevr", "mij", "mijn", "minder", "miss", "misschien", "missen", "mits", "mocht", "mochten",
    "moest", "moesten", "moet", "moeten", "mogen", "mr", "mrs", "mw", "na", "naar", "nam", "namelijk", "nee", "neem", "negen",
    "nemen", "nergens", "niemand", "niet", "niets", "niks", "noch", "nochtans", "nog", "nooit", "nu", "nv", "of", "om", "omdat",
    "ondanks", "onder", "ondertussen", "ons", "onze", "onzeker", "ooit", "ook", "op", "over", "overal", "overige", "paar", "per",
    "recent", "redelijk", "samen", "sinds", "steeds", "te", "tegen", "tegenover", "thans", "tien", "tiende", "tijdens", "tja", "toch",
    "toe", "tot", "totdat", "tussen", "twee", "tweede", "u", "uit", "uw", "vaak", "van", "vanaf", "veel", "veertig", "verder",
    "verscheidene", "verschillende", "via", "vier", "vierde", "vijf", "vijfde", "vijftig", "volgend", "volgens", "voor", "voordat",
    "voorts", "waar", "waarom", "waarschijnlijk", "wanneer", "waren", "was", "wat", "we", "wederom", "weer", "weinig", "wel", "welk",
    "welke", "werd", "werden", "werder", "whatever", "wie", "wij", "wijzelf", "wil", "wilden", "willen", "word", "worden", "wordt", "zal",
    "ze", "zei", "zeker", "zelf", "zelfde", "zes", "zeven", "zich", "zij", "zijn", "zijzelf", "zo", "zoals", "zodat", "zou", "zouden",
    "zulk", "zullen"}

#End Region

            End Module

            ''' <summary>
            ''' QuestionTypes are
            '''place           'Where is X - Returns Answer to Question (Where is becomes the subject) X is at location of Subject
            '''time            'When was - Returns Answer to Question (when was X) X was Subject
            '''person          'Who is - Returns Answer to Question - Object becomes the search term for answer replaced by Subject
            '''reason          'Why is - Returns Answer to Question
            '''description     'What is - Returns Answer to Question
            '''instruction     'How do - Returns Answer to Question
            '''Choice          'Which X, either X or Y - Returns Selected choice from input (Which becomes the Subject X and Y are both the objects)
            '''Possession      'Whose X is? - returns Question about Ownership (Whose is the subject X is the object) X belongs to object
            ''' </summary>
            Public Enum _QuestionType
                place           'Where is - Returns Answer to Question
                time            'When was - Returns Answer to Question
                person          'Who is - Returns Answer to Question
                reason          'Why is - Returns Answer to Question
                description     'What is - Returns Answer to Question
                instruction     'How do - Returns Answer to Question
                Choice          'Which X, either X or Y - Returns Selected choice from input
                Possession      'Whose X is? - returns Question about Ownwership
            End Enum

            ''' <summary>
            ''' These are the standard concept types / The future types will be a list collected fromt the
            ''' database the database will only have dynamic predicate types
            ''' </summary>
            Public Enum ConceptType

                'Thing
                CapeableOf 'Something that A can typically do is B.

                ConceptuallyRelatedTo 'The most general relation. There is some positive relationship between A and B, but ConceptNet can't determine what that relationship is based on the data. This was called "ConceptuallyRelatedTo" in
                CapableOfReceivingAction ' Subject A can be eaten, can be washed
                DefinedAs 'A and B overlap considerably in meaning, and B is a more explanatory version of A. (This is similar to TranslationOf, but within one language.)
                DesireOf 'A is a conscious entity that typically wants B. Many assertions of this type use the appropriate language's word for "person" as A
                DesirousEffectOf 'Jump / is to land
                EffectOf 'Singing / is to Entertain
                IsA 'A is a subtype or a specific instance of B; every A is a B. (We do not make the type-token distinction, because people don't usually make that distinction.) This is the hyponym relation in WordNet.
                MadeOf 'A has composition property's of
                MotivationOf 'Someone does A because they want result B; A is a step toward accomplishing the goal B.
                PartOf 'A is a part of B. This is the part meronym relation in WordNet.
                PropertyOf 'A has B as a property; A can be described as B.
                UsedFor 'A is used for B; the purpose of A is B.

                'Event
                FirstSubEventOf 'A is an event that begins with subevent B.

                LastSubEventOf 'A is an event that concludes with subevent B.
                PreRequisiteOf 'In order for A to happen, B needs to happen; B is a dependency of A.
                SubeventOf 'A and B are events, and B happens as a subevent of A.

                'Place
                LocationOf 'A is a typical location for B, or A is the inherent location of B. Some instances of this would be considered meronyms in WordNet.

            End Enum

            Public Enum CordinatingConjunction
                _for
                _and
                _nor
                _but
                _or
                _yet
                _so
            End Enum

            Public Enum ModalGroups
                Ability
                Request
                Permission
                Possibility
                Obligation
                Prohibition
                Lack_Of_necessity
                Advice
                probability
                Conditional
                Invitation
                Suggestion
                deduction
                Question
                necsessity
                prediction
                offer
                promise
                decision
                Prefference
                intention
            End Enum

            ''' <summary>
            ''' These are the different object types available these object types are the highest level
            ''' objects of which most things sub objects or combined types
            ''' </summary>
            ''' <remarks></remarks>
            Enum ObjectType
                _Location
                _Event
                _Thing
                _Animal
                _Plant
                _Person
            End Enum

            ''' <summary>
            ''' Various tools for Working with strings
            ''' </summary>
            Public Class WordTools

                'Morse Code
                Private MorseCode() As String = {".", "-"}

                Public ReadOnly AlphaBet() As String = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N",
            "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n",
            "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z"}

                Public ReadOnly CodePunctuation() As String = {"\", "#", "@", "^"}
                Public ReadOnly EncapuslationPunctuationEnd() As String = {"}", "]", ">", ")"}
                Public ReadOnly EncapuslationPunctuationStart() As String = {"{", "[", "<", "("}
                Public ReadOnly GramaticalPunctuation() As String = {".", "?", "!", ":", ";"}
                Public ReadOnly MathPunctuation() As String = {"+", "-", "=", "/", "*", "%", "PLUS", "ADD", "MINUS", "SUBTRACT", "DIVIDE", "DIFFERENCE", "TIMES", "MULTIPLY", "PERCENT", "EQUALS"}
                Public ReadOnly MoneyPunctuation() As String = {"£", "$"}

                Public ReadOnly Number() As String = {"1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20",
        "30", "40", "50", "60", "70", "80", "90", "00", "000", "0000", "00000", "000000", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen",
        "nineteen", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety", "hundred", "thousand", "million", "Billion"}

                Public ReadOnly SeperatorPunctuation() = {" ", ",", "|"}

                ''' <summary>
                ''' counts occurrences of a specific phoneme
                ''' </summary>
                ''' <param name="strIn">  </param>
                ''' <param name="strFind"></param>
                ''' <returns></returns>
                ''' <remarks></remarks>
                Public Shared Function CountOccurrences(ByRef strIn As String, ByRef strFind As String) As Integer
                    '**
                    ' Returns: the number of times a string appears in a string
                    '
                    '@rem           Example code for CountOccurrences()
                    '
                    '  ' Counts the occurrences of "ow" in the supplied string.
                    '
                    '    strTmp = "How now, brown cow"
                    '    Returns a value of 4
                    '
                    '
                    'Debug.Print "CountOccurrences(): there are " &  CountOccurrences(strTmp, "ow") &
                    '" occurrences of 'ow'" &    " in the string '" & strTmp & "'"
                    '
                    '@param        strIn Required. String.
                    '@param        strFind Required. String.
                    '@return       Long.

                    Dim lngPos As Integer
                    Dim lngWordCount As Integer

                    On Error GoTo PROC_ERR

                    lngWordCount = 1

                    ' Find the first occurrence
                    lngPos = InStr(strIn, strFind)

                    Do While lngPos > 0
                        ' Find remaining occurrences
                        lngPos = InStr(lngPos + 1, strIn, strFind)
                        If lngPos > 0 Then
                            ' Increment the hit counter
                            lngWordCount = lngWordCount + 1
                        End If
                    Loop

                    ' Return the value
                    CountOccurrences = lngWordCount

PROC_EXIT:
                    Exit Function

PROC_ERR:
                    MsgBox("Error: " & Err.Number & ". " & Err.Description, , NameOf(CountOccurrences))
                    Resume PROC_EXIT

                End Function

                ''' <summary>
                ''' This will extract the first word from a statement. Example: Message="Hello world"
                ''' Word=ExtractWord(Message) Print Word Hello Print Message world
                ''' </summary>
                ''' <param name="Statement"></param>
                ''' <returns></returns>
                ''' <remarks></remarks>
                Public Shared Function ExtractFirstWord(ByRef Statement As String) As String
                    Dim StrArr() As String = Split(Statement, " ")
                    Return StrArr(0)
                End Function

                ''' <summary>
                ''' This will extract the words within the (Text) between (word1) and (word2). Example:
                ''' Text="The rain in Spain." Print ExtractWordsBetween("the "," in ",Text) rain
                ''' </summary>
                ''' <param name="Word1"></param>
                ''' <param name="Word2"></param>
                ''' <param name="Text"> </param>
                ''' <returns></returns>
                ''' <remarks></remarks>
                Public Shared Function ExtractWordsBetween(ByRef Word1 As String, ByRef Word2 As String, ByRef Text As String) As String

                    '
                    Dim Position As Short
                    Dim Position2 As Short

                    If Word1 = "" Then Position = 1 : GoTo Lp1
                    Position = InStr(1, Text, Word1)
                    If Position = 0 Then Position = 1
                    Position = Position + Len(Word1)
Lp1:

                    If Word2 = "" Then Position2 = Len(Text) + 1 : GoTo Lp2
                    Position2 = InStr(Position, Text, Word2)
                    If Position2 = 0 Then Position2 = Len(Text) + 1
Lp2:
                    ExtractWordsBetween = Mid(Text, Position, Position2 - Position)
                    ExtractWordsBetween = Trim(ExtractWordsBetween)
                End Function

                ''' <summary>
                ''' </summary>
                ''' <param name="tmpStr">The whole string</param>
                ''' <param name="tmpDiv">The Divider chr</param>
                ''' <param name="Mode">
                ''' if mode = "F" then get the string in front of the div if mode = "B" then get the string
                ''' in back of the div
                ''' </param>
                ''' <returns></returns>
                ''' <remarks>Example: String_GetString("Red=Green","=","F") will return "Red"</remarks>
                Public Shared Function ExtractWordsEitherSide(ByRef tmpStr As String, ByRef tmpDiv As String, ByRef Mode As String) As String

                    ' Get strings on either side of special character. Here is a function that will get the
                    ' string on each side of a string that is really two strings, with a special char in the
                    ' center. Return values: Errors Err1 = wrong mode ("F" & "B" ok) Err2 = No Div, did not
                    ' found a divider Err3 = No F, did not find any string in front of the div Err4 = No B,
                    ' did not find any string in back of the div Err5 = No String
                    Dim tmpRetStr As String

                    Dim tmpErr As Short
                    tmpErr = 0
                    If Len(tmpStr) = 0 Then
                        tmpErr = 5
                        GoTo GetStringErr
                    End If

                    ' *** test for chr in tmpDiv
                    If Len(tmpDiv) = 0 Then
                        tmpErr = 2
                        GoTo GetStringErr
                    End If

                    ' *** test for div in string
                    If InStr(1, tmpStr, tmpDiv) = 0 Then
                        tmpErr = 2
                        GoTo GetStringErr
                    End If

                    ' *** process "F"
                    If Mode = "F" Then
                        tmpRetStr = Left(tmpStr, InStr(1, tmpStr, tmpDiv) - 1)
                        If Len(tmpRetStr) > 0 Then
                            ExtractWordsEitherSide = tmpRetStr
                            Exit Function
                        Else
                            tmpErr = 3
                            GoTo GetStringErr
                        End If
                    End If

                    If Mode = "B" Then
                        tmpRetStr = Right(tmpStr, Len(tmpStr) - (InStr(1, tmpStr, tmpDiv)))
                        If Len(tmpRetStr) > 0 Then
                            ExtractWordsEitherSide = tmpRetStr
                            Exit Function
                        Else
                            tmpErr = 4
                            GoTo GetStringErr
                        End If

                    End If
                    tmpErr = 1
                    GoTo GetStringErr
                    Exit Function
GetStringErr:
                    ExtractWordsEitherSide = "Error " & tmpErr
                End Function

                ' Generate a random number based on the upper and lower bounds of the array,
                'then use that to return the item.
                Public Shared Function FetchRandomItem(Of t)(ByRef theArray() As t) As t

                    Dim randNumberGenerator As New Random
                    Randomize()
                    Dim index As Integer = randNumberGenerator.Next(theArray.GetLowerBound(0),
                                                        theArray.GetUpperBound(0) + 1)

                    Return theArray(index)

                End Function

                Public Shared Function FormatText(ByRef Text As String) As String
                    Dim FormatTextResponse As String = ""
                    'FORMAT USERINPUT
                    'turn to uppercase for searching the db
                    Text = LTrim(Text)
                    Text = RTrim(Text)
                    Text = UCase(Text)

                    FormatTextResponse = Text
                    Return FormatTextResponse
                End Function

                Public Shared Function GetBetween(ByRef sSearch As String,
                   ByRef sStart As String,
                   ByRef sStop As String) As String
                    'Use GetBetween to Parse a String Between Two Strings
                    'This tips shows how to use GetBetween to parse and extract a string between two strings.
                    ' GetBetween returns the extracted string, if found.
                    GetBetween = ""
                    Dim lSearch As Integer = 1
                    lSearch = InStr(lSearch, sSearch, sStart, 1)
                    'InStr(lSearch, sSearch, sStart)

                    If lSearch > 0 Then
                        lSearch = lSearch + Len(sStart)
                        Dim lTemp As Long
                        lTemp = InStr(lSearch, sSearch, sStop)
                        If lTemp > lSearch Then
                            GetBetween = Mid$(sSearch, lSearch, lTemp - lSearch)
                        End If
                    End If
                End Function

                ''' <summary>
                ''' Advanced search String pattern Wildcard denotes which position 1st =1 or 2nd =2 Send
                ''' Original input &gt; Search pattern to be used &gt; Wildcard requred SPattern = "WHAT
                ''' COLOUR DO YOU LIKE * OR *" Textstr = "WHAT COLOUR DO YOU LIKE red OR black" ITEM_FOUND =
                ''' = SearchPattern(USERINPUT, SPattern, 1) ---- RETURNS RED ITEM_FOUND = =
                ''' SearchPattern(USERINPUT, SPattern, 1) ---- RETURNS black
                ''' </summary>
                ''' <param name="TextSTR"> TextStr Required. String.</param>
                ''' <param name="SPattern">SPattern Required. String.</param>
                ''' <param name="Wildcard">Wildcard Required. Integer.</param>
                ''' <returns></returns>
                ''' <remarks>* in search pattern</remarks>
                Public Shared Function SearchPattern(ByRef TextSTR As String, ByRef SPattern As String, ByRef Wildcard As Short) As String
                    Dim SearchP2 As String
                    Dim SearchP1 As String
                    Dim TextStrp3 As String
                    Dim TextStrp4 As String
                    SearchPattern = ""
                    SearchP2 = ""
                    SearchP1 = ""
                    TextStrp3 = ""
                    TextStrp4 = ""
                    If TextSTR Like SPattern = True Then
                        Select Case Wildcard
                            Case 1
                                Call SplitPhrase(SPattern, "*", SearchP1, SearchP2)
                                TextSTR = Replace(TextSTR, SearchP1, "", 1, -1, CompareMethod.Text)

                                SearchP2 = Replace(SearchP2, "*", "", 1, -1, CompareMethod.Text)
                                Call SplitPhrase(TextSTR, SearchP2, TextStrp3, TextStrp4)

                                TextSTR = TextStrp3

                            Case 2
                                Call SplitPhrase(SPattern, "*", SearchP1, SearchP2)
                                SPattern = Replace(SPattern, SearchP1, " ", 1, -1, CompareMethod.Text)
                                TextSTR = Replace(TextSTR, SearchP1, " ", 1, -1, CompareMethod.Text)

                                Call SplitPhrase(SearchP2, "*", SearchP1, SearchP2)
                                Call SplitPhrase(TextSTR, SearchP1, TextStrp3, TextStrp4)

                                TextSTR = TextStrp4

                        End Select

                        SearchPattern = TextSTR
                        LTrim(SearchPattern)
                        RTrim(SearchPattern)
                    Else
                    End If

                End Function

                ''' <summary>
                ''' SPLITS THE GIVEN PHRASE UP INTO TWO PARTS by dividing word SplitPhrase(Userinput, "and",
                ''' Firstp, SecondP)
                ''' </summary>
                ''' <param name="PHRASE">      Sentence to be divided</param>
                ''' <param name="DIVIDINGWORD">String: Word to divide sentence by</param>
                ''' <param name="FIRSTPART">   String: firstpart of sentence to be populated</param>
                ''' <param name="SECONDPART">  String: Secondpart of sentence to be populated</param>
                ''' <remarks></remarks>
                Public Shared Sub SplitPhrase(ByVal PHRASE As String, ByRef DIVIDINGWORD As String, ByRef FIRSTPART As String, ByRef SECONDPART As String)
                    Dim POS As Short
                    POS = InStr(PHRASE, DIVIDINGWORD)
                    If (POS > 0) Then
                        FIRSTPART = Trim(Left(PHRASE, POS - 1))
                        SECONDPART = Trim(Right(PHRASE, Len(PHRASE) - POS - Len(DIVIDINGWORD) + 1))
                    Else
                        FIRSTPART = ""
                        SECONDPART = PHRASE
                    End If
                End Sub

            End Class

            ''' <summary>
            ''' TO USE THE PROGRAM CALL THE FUNCTION PORTERALGORITHM. THE WORD
            ''' TO BE STEMMED SHOULD BE PASSED AS THE ARGUEMENT ARGUEMENT. THE STRING
            ''' RETURNED BY THE FUNCTION IS THE STEMMED WORD
            ''' Porter Stemmer. It follows the algorithm definition
            ''' presented in :
            '''   Porter, 1980, An algorithm for suffix stripping, Program, Vol. 14,
            '''   no. 3, pp 130-137,
            ''' </summary>
            Public Class WordStemmer

                '   (http://www.tartarus.org/~martin/PorterStemmer)

                'Author : Navonil Mustafee
                'Brunel University - student
                'Algorithm Implemented as part for assignment on document visualization

                Private Shared Function porterAlgorithmStep1(str As String) As String

                    On Error Resume Next

                    'STEP 1A
                    '
                    '    SSES -> SS                         caresses  ->  caress
                    '    IES  -> I                          ponies    ->  poni
                    '                                       ties      ->  ti
                    '    SS   -> SS                         caress    ->  caress
                    '    S    ->                            cats      ->  cat

                    'declaring local variables
                    Dim i As Byte
                    Dim j As Byte
                    Dim step1a(3, 1) As String

                    'initializing contents of 2D array
                    step1a(0, 0) = "sses"
                    step1a(0, 1) = "ss"
                    step1a(1, 0) = "ies"
                    step1a(1, 1) = "i"
                    step1a(2, 0) = "ss"
                    step1a(2, 1) = "ss"
                    step1a(3, 0) = "s"
                    step1a(3, 1) = ""

                    'checking word
                    For i = 0 To 3 Step 1
                        If porterEndsWith(str, step1a(i, 0)) Then
                            str = porterTrimEnd(str, Len(step1a(i, 0)))
                            str = porterAppendEnd(str, step1a(i, 1))
                            Exit For
                        End If
                    Next i

                    '--------------------------------------------------------------------------------------------------------

                    'STEP 1B
                    '
                    '   If
                    '       (m>0) EED -> EE                     feed      ->  feed
                    '                                           agreed    ->  agree
                    '   Else
                    '       (*v*) ED  ->                        plastered ->  plaster
                    '                                           bled      ->  bled
                    '       (*v*) ING ->                        motoring  ->  motor
                    '                                           sing      ->  sing
                    '
                    'If the second or third of the rules in Step 1b is successful, the following
                    'is done:
                    '
                    '    AT -> ATE                       conflat(ed)  ->  conflate
                    '    BL -> BLE                       troubl(ed)   ->  trouble
                    '    IZ -> IZE                       siz(ed)      ->  size
                    '    (*d and not (*L or *S or *Z))
                    '       -> single letter
                    '                                    hopp(ing)    ->  hop
                    '                                    tann(ed)     ->  tan
                    '                                    fall(ing)    ->  fall
                    '                                    hiss(ing)    ->  hiss
                    '                                    fizz(ed)     ->  fizz
                    '    (m=1 and *o) -> E               fail(ing)    ->  fail
                    '                                    fil(ing)     ->  file
                    '
                    'The rule to map to a single letter causes the removal of one of the double
                    'letter pair. The -E is put back on -AT, -BL and -IZ, so that the suffixes
                    '-ATE, -BLE and -IZE can be recognised later. This E may be removed in step
                    '4.

                    'declaring local variables
                    Dim m As Byte
                    Dim temp As String
                    Dim second_third_success As Boolean

                    'initializing contents of 2D array
                    second_third_success = False

                    '(m>0) EED -> EE..else..(*v*) ED  ->(*v*) ING  ->
                    If porterEndsWith(str, "eed") Then

                        'counting the number of m's
                        temp = porterTrimEnd(str, Len("eed"))
                        m = porterCountm(temp)

                        If m > 0 Then
                            str = porterTrimEnd(str, Len("eed"))
                            str = porterAppendEnd(str, "ee")
                        End If

                    ElseIf porterEndsWith(str, "ed") Then

                        'trim and check for vowel
                        temp = porterTrimEnd(str, Len("ed"))

                        If porterContainsVowel(temp) Then
                            str = porterTrimEnd(str, Len("ed"))
                            second_third_success = True
                        End If

                    ElseIf porterEndsWith(str, "ing") Then

                        'trim and check for vowel
                        temp = porterTrimEnd(str, Len("ing"))

                        If porterContainsVowel(temp) Then
                            str = porterTrimEnd(str, Len("ing"))
                            second_third_success = True
                        End If

                    End If

                    'If the second or third of the rules in Step 1b is SUCCESSFUL, the following
                    'is done:
                    '
                    '    AT -> ATE                       conflat(ed)  ->  conflate
                    '    BL -> BLE                       troubl(ed)   ->  trouble
                    '    IZ -> IZE                       siz(ed)      ->  size
                    '    (*d and not (*L or *S or *Z))
                    '       -> single letter
                    '                                    hopp(ing)    ->  hop
                    '                                    tann(ed)     ->  tan
                    '                                    fall(ing)    ->  fall
                    '                                    hiss(ing)    ->  hiss
                    '                                    fizz(ed)     ->  fizz
                    '    (m=1 and *o) -> E               fail(ing)    ->  fail
                    '                                    fil(ing)     ->  file

                    If second_third_success = True Then             'If the second or third of the rules in Step 1b is SUCCESSFUL

                        If porterEndsWith(str, "at") Then           'AT -> ATE
                            str = porterTrimEnd(str, Len("at"))
                            str = porterAppendEnd(str, "ate")
                        ElseIf porterEndsWith(str, "bl") Then       'BL -> BLE
                            str = porterTrimEnd(str, Len("bl"))
                            str = porterAppendEnd(str, "ble")
                        ElseIf porterEndsWith(str, "iz") Then       'IZ -> IZE
                            str = porterTrimEnd(str, Len("iz"))
                            str = porterAppendEnd(str, "ize")
                        ElseIf porterEndsDoubleConsonent(str) Then  '(*d and not (*L or *S or *Z))-> single letter
                            If Not (porterEndsWith(str, "l") Or porterEndsWith(str, "s") Or porterEndsWith(str, "z")) Then
                                str = porterTrimEnd(str, 1)
                            End If
                        ElseIf porterCountm(str) = 1 Then                           '(m=1 and *o) -> E
                            If porterEndsCVC(str) Then
                                str = porterAppendEnd(str, "e")
                            End If
                        End If

                    End If

                    '--------------------------------------------------------------------------------------------------------
                    '
                    'STEP 1C
                    '
                    '    (*v*) Y -> I                    happy        ->  happi
                    '                                    sky          ->  sky

                    If porterEndsWith(str, "y") Then

                        'trim and check for vowel
                        temp = porterTrimEnd(str, 1)

                        If porterContainsVowel(temp) Then
                            str = porterTrimEnd(str, Len("y"))
                            str = porterAppendEnd(str, "i")
                        End If

                    End If

                    'retuning the word
                    porterAlgorithmStep1 = str

                End Function

                Private Shared Function porterAlgorithmStep2(str As String) As String

                    On Error Resume Next

                    'STEP 2
                    '
                    '    (m>0) ATIONAL ->  ATE           relational     ->  relate
                    '    (m>0) TIONAL  ->  TION          conditional    ->  condition
                    '                                    rational       ->  rational
                    '    (m>0) ENCI    ->  ENCE          valenci        ->  valence
                    '    (m>0) ANCI    ->  ANCE          hesitanci      ->  hesitance
                    '    (m>0) IZER    ->  IZE           digitizer      ->  digitize
                    'Also,
                    '    (m>0) BLI    ->   BLE           conformabli    ->  conformable
                    '
                    '    (m>0) ALLI    ->  AL            radicalli      ->  radical
                    '    (m>0) ENTLI   ->  ENT           differentli    ->  different
                    '    (m>0) ELI     ->  E             vileli        - >  vile
                    '    (m>0) OUSLI   ->  OUS           analogousli    ->  analogous
                    '    (m>0) IZATION ->  IZE           vietnamization ->  vietnamize
                    '    (m>0) ATION   ->  ATE           predication    ->  predicate
                    '    (m>0) ATOR    ->  ATE           operator       ->  operate
                    '    (m>0) ALISM   ->  AL            feudalism      ->  feudal
                    '    (m>0) IVENESS ->  IVE           decisiveness   ->  decisive
                    '    (m>0) FULNESS ->  FUL           hopefulness    ->  hopeful
                    '    (m>0) OUSNESS ->  OUS           callousness    ->  callous
                    '    (m>0) ALITI   ->  AL            formaliti      ->  formal
                    '    (m>0) IVITI   ->  IVE           sensitiviti    ->  sensitive
                    '    (m>0) BILITI  ->  BLE           sensibiliti    ->  sensible
                    'Also,
                    '    (m>0) LOGI    ->  LOG           apologi        -> apolog
                    '
                    'The test for the string S1 can be made fast by doing a program switch on
                    'the penultimate letter of the word being tested. This gives a fairly even
                    'breakdown of the possible values of the string S1. It will be seen in fact
                    'that the S1-strings in step 2 are presented here in the alphabetical order
                    'of their penultimate letter. Similar techniques may be applied in the other
                    'steps.

                    'declaring local variables
                    Dim step2(20, 1) As String
                    Dim i As Byte
                    Dim temp As String

                    'initializing contents of 2D array
                    step2(0, 0) = "ational"
                    step2(0, 1) = "ate"
                    step2(1, 0) = "tional"
                    step2(1, 1) = "tion"
                    step2(2, 0) = "enci"
                    step2(2, 1) = "ence"
                    step2(3, 0) = "anci"
                    step2(3, 1) = "ance"
                    step2(4, 0) = "izer"
                    step2(4, 1) = "ize"
                    step2(5, 0) = "bli"
                    step2(5, 1) = "ble"
                    step2(6, 0) = "alli"
                    step2(6, 1) = "al"
                    step2(7, 0) = "entli"
                    step2(7, 1) = "ent"
                    step2(8, 0) = "eli"
                    step2(8, 1) = "e"
                    step2(9, 0) = "ousli"
                    step2(9, 1) = "ous"
                    step2(10, 0) = "ization"
                    step2(10, 1) = "ize"
                    step2(11, 0) = "ation"
                    step2(11, 1) = "ate"
                    step2(12, 0) = "ator"
                    step2(12, 1) = "ate"
                    step2(13, 0) = "alism"
                    step2(13, 1) = "al"
                    step2(14, 0) = "iveness"
                    step2(14, 1) = "ive"
                    step2(15, 0) = "fulness"
                    step2(15, 1) = "ful"
                    step2(16, 0) = "ousness"
                    step2(16, 1) = "ous"
                    step2(17, 0) = "aliti"
                    step2(17, 1) = "al"
                    step2(18, 0) = "iviti"
                    step2(18, 1) = "ive"
                    step2(19, 0) = "biliti"
                    step2(19, 1) = "ble"
                    step2(20, 0) = "logi"
                    step2(20, 1) = "log"

                    'checking word
                    For i = 0 To 20 Step 1
                        If porterEndsWith(str, step2(i, 0)) Then
                            temp = porterTrimEnd(str, Len(step2(i, 0)))
                            If porterCountm(temp) > 0 Then
                                str = porterTrimEnd(str, Len(step2(i, 0)))
                                str = porterAppendEnd(str, step2(i, 1))
                            End If
                            Exit For
                        End If
                    Next i

                    'retuning the word
                    porterAlgorithmStep2 = str

                End Function

                Private Shared Function porterAlgorithmStep3(str As String) As String

                    On Error Resume Next

                    'STEP 3
                    '
                    '    (m>0) ICATE ->  IC              triplicate     ->  triplic
                    '    (m>0) ATIVE ->                  formative      ->  form
                    '    (m>0) ALIZE ->  AL              formalize      ->  formal
                    '    (m>0) ICITI ->  IC              electriciti    ->  electric
                    '    (m>0) ICAL  ->  IC              electrical     ->  electric
                    '    (m>0) FUL   ->                  hopeful        ->  hope
                    '    (m>0) NESS  ->                  goodness       ->  good

                    'declaring local variables
                    Dim i As Byte
                    Dim temp As String
                    Dim step3(6, 1) As String

                    'initializing contents of 2D array
                    step3(0, 0) = "icate"
                    step3(0, 1) = "ic"
                    step3(1, 0) = "ative"
                    step3(1, 1) = ""
                    step3(2, 0) = "alize"
                    step3(2, 1) = "al"
                    step3(3, 0) = "iciti"
                    step3(3, 1) = "ic"
                    step3(4, 0) = "ical"
                    step3(4, 1) = "ic"
                    step3(5, 0) = "ful"
                    step3(5, 1) = ""
                    step3(6, 0) = "ness"
                    step3(6, 1) = ""

                    'checking word
                    For i = 0 To 6 Step 1
                        If porterEndsWith(str, step3(i, 0)) Then
                            temp = porterTrimEnd(str, Len(step3(i, 0)))
                            If porterCountm(temp) > 0 Then
                                str = porterTrimEnd(str, Len(step3(i, 0)))
                                str = porterAppendEnd(str, step3(i, 1))
                            End If
                            Exit For
                        End If
                    Next i

                    'retuning the word
                    porterAlgorithmStep3 = str

                End Function

                Private Shared Function porterAlgorithmStep4(str As String) As String

                    On Error Resume Next

                    'STEP 4
                    '
                    '    (m>1) AL    ->                  revival        ->  reviv
                    '    (m>1) ANCE  ->                  allowance      ->  allow
                    '    (m>1) ENCE  ->                  inference      ->  infer
                    '    (m>1) ER    ->                  airliner       ->  airlin
                    '    (m>1) IC    ->                  gyroscopic     ->  gyroscop
                    '    (m>1) ABLE  ->                  adjustable     ->  adjust
                    '    (m>1) IBLE  ->                  defensible     ->  defens
                    '    (m>1) ANT   ->                  irritant       ->  irrit
                    '    (m>1) EMENT ->                  replacement    ->  replac
                    '    (m>1) MENT  ->                  adjustment     ->  adjust
                    '    (m>1) ENT   ->                  dependent      ->  depend
                    '    (m>1 and (*S or *T)) ION ->     adoption       ->  adopt
                    '    (m>1) OU    ->                  homologou      ->  homolog
                    '    (m>1) ISM   ->                  communism      ->  commun
                    '    (m>1) ATE   ->                  activate       ->  activ
                    '    (m>1) ITI   ->                  angulariti     ->  angular
                    '    (m>1) OUS   ->                  homologous     ->  homolog
                    '    (m>1) IVE   ->                  effective      ->  effect
                    '    (m>1) IZE   ->                  bowdlerize     ->  bowdler
                    '
                    'The suffixes are now removed. All that remains is a little tidying up.

                    'declaring local variables
                    Dim i As Byte
                    Dim temp As String
                    Dim step4(18) As String

                    'initializing contents of 2D array
                    step4(0) = "al"
                    step4(1) = "ance"
                    step4(2) = "ence"
                    step4(3) = "er"
                    step4(4) = "ic"
                    step4(5) = "able"
                    step4(6) = "ible"
                    step4(7) = "ant"
                    step4(8) = "ement"
                    step4(9) = "ment"
                    step4(10) = "ent"
                    step4(11) = "ion"
                    step4(12) = "ou"
                    step4(13) = "ism"
                    step4(14) = "ate"
                    step4(15) = "iti"
                    step4(16) = "ous"
                    step4(17) = "ive"
                    step4(18) = "ize"

                    'checking word
                    For i = 0 To 18 Step 1

                        If porterEndsWith(str, step4(i)) Then

                            temp = porterTrimEnd(str, Len(step4(i)))

                            If porterCountm(temp) > 1 Then

                                If porterEndsWith(str, "ion") Then
                                    If porterEndsWith(temp, "s") Or porterEndsWith(temp, "t") Then
                                        str = porterTrimEnd(str, Len(step4(i)))
                                        str = porterAppendEnd(str, "")
                                    End If
                                Else
                                    str = porterTrimEnd(str, Len(step4(i)))
                                    str = porterAppendEnd(str, "")
                                End If

                            End If

                            Exit For

                        End If

                    Next i

                    'retuning the word
                    porterAlgorithmStep4 = str

                End Function

                Private Shared Function porterAlgorithmStep5(str As String) As String

                    On Error Resume Next

                    'STEP 5a
                    '
                    '    (m>1) E     ->                  probate        ->  probat
                    '                                    rate           ->  rate
                    '    (m=1 and not *o) E ->           cease          ->  ceas
                    '
                    'STEP 5b
                    '
                    '    (m>1 and *d and *L) -> single letter
                    '                                    controll       ->  control
                    '                                    roll           ->  roll

                    'declaring local variables
                    Dim i As Byte
                    Dim temp As String

                    'Step5a
                    If porterEndsWith(str, "e") Then            'word ends with e
                        temp = porterTrimEnd(str, 1)
                        If porterCountm(temp) > 1 Then          'm>1
                            str = porterTrimEnd(str, 1)
                        ElseIf porterCountm(temp) = 1 Then      'm=1
                            If Not porterEndsCVC(temp) Then     'not *o
                                str = porterTrimEnd(str, 1)
                            End If
                        End If
                    End If

                    '--------------------------------------------------------------------------------------------------------
                    '
                    'Step5b
                    If porterCountm(str) > 1 Then
                        If porterEndsDoubleConsonent(str) And porterEndsWith(str, "l") Then
                            str = porterTrimEnd(str, 1)
                        End If
                    End If

                    'retuning the word
                    porterAlgorithmStep5 = str

                End Function

                Private Shared Function porterAppendEnd(str As String, ends As String) As String

                    On Error Resume Next

                    'returning the appended string
                    porterAppendEnd = str + ends

                End Function

                Private Shared Function porterContains(str As String, present As String) As Boolean

                    On Error Resume Next

                    'checking whether strr contains present
                    porterContains = If(InStr(str, present) = 0, False, True)

                End Function

                Private Shared Function porterContainsVowel(str As String) As Boolean

                    'checking word to see if vowels are present

                    Dim pattern As String

                    If Len(str) >= 0 Then

                        'find out the CVC pattern
                        pattern = returnCVCpattern(str)

                        'check to see if the return pattern contains a vowel
                        porterContainsVowel = If(InStr(pattern, "v") = 0, False, True)
                    Else
                        porterContainsVowel = False
                    End If

                End Function

                Private Shared Function porterCountm(str As String) As Byte

                    On Error Resume Next

                    'A \consonant\ in a word is a letter other than A, E, I, O or U, and other
                    'than Y preceded by a consonant. (The fact that the term `consonant' is
                    'defined to some extent in terms of itself does not make it ambiguous.) So in
                    'TOY the consonants are T and Y, and in SYZYGY they are S, Z and G. If a
                    'letter is not a consonant it is a \vowel\.

                    'declaring local variables
                    Dim chars() As Byte
                    Dim const_vowel As String
                    Dim i As Byte
                    Dim m As Byte
                    Dim flag As Boolean
                    Dim pattern As String

                    'initializing
                    const_vowel = ""
                    m = 0
                    flag = False

                    If Not Len(str) = 0 Then

                        'find out the CVC pattern
                        pattern = returnCVCpattern(str)

                        'converting const_vowel to byte array
                        chars = System.Text.Encoding.Unicode.GetBytes(pattern)

                        'counting the number of m's...
                        For i = 0 To UBound(chars) Step 1
                            If Chr(chars(i)) = "v" Or flag = True Then
                                flag = True
                                If Chr(chars(i)) = "c" Then
                                    m = m + 1
                                    flag = False
                                End If
                            End If
                        Next i

                    End If

                    porterCountm = m

                End Function

                Private Shared Function porterEndsCVC(str As String) As Boolean

                    On Error Resume Next

                    '*o  - the stem ends cvc, where the second c is not W, X or Y (e.g. -WIL, -HOP).

                    'declaring local variables
                    Dim chars() As Byte
                    Dim const_vowel As String
                    Dim i As Byte
                    Dim pattern As String

                    'check to see if atleast 3 characters are present
                    If Len(str) >= 3 Then

                        'converting string to byte array

                        chars = System.Text.Encoding.Unicode.GetBytes(str)

                        'find out the CVC pattern
                        pattern = returnCVCpattern(str)

                        'we need to check only the last three characters
                        pattern = Right(pattern, 3)

                        'check to see if the letters in str match the sequence cvc
                        porterEndsCVC = If(pattern = "cvc", If(Not (Chr(chars(UBound(chars))) = "w" Or Chr(chars(UBound(chars))) = "x" Or Chr(chars(UBound(chars))) = "y"), True, False), False)
                    Else

                        porterEndsCVC = False

                    End If

                End Function

                Private Shared Function porterEndsDoubleConsonent(str As String) As Boolean

                    On Error Resume Next

                    'checking whether word ends with a double consonant (e.g. -TT, -SS).

                    'declaring local variables
                    Dim holds_ends As String
                    Dim hold_third_last As String
                    Dim chars() As Byte

                    'first check whether the size of the word is >= 2
                    If Len(str) >= 2 Then

                        'extract 2 characters from right of str
                        holds_ends = Right(str, 2)

                        'converting string to byte array
                        chars = System.Text.Encoding.Unicode.GetBytes(holds_ends)

                        'checking if both the characters are same
                        If chars(0) = chars(1) Then

                            'check for double consonent
                            If holds_ends = "aa" Or holds_ends = "ee" Or holds_ends = "ii" Or holds_ends = "oo" Or holds_ends = "uu" Then

                                porterEndsDoubleConsonent = False
                            Else

                                'if the second last character is y, and there are atleast three letters in str
                                If holds_ends = "yy" And Len(str) > 2 Then

                                    'extracting the third last character
                                    hold_third_last = Right(str, 3)
                                    hold_third_last = Left(str, 1)

                                    porterEndsDoubleConsonent = If(Not (hold_third_last = "a" Or hold_third_last = "e" Or hold_third_last = "i" Or hold_third_last = "o" Or hold_third_last = "u"), False, True)
                                Else

                                    porterEndsDoubleConsonent = True

                                End If

                            End If
                        Else

                            porterEndsDoubleConsonent = False

                        End If
                    Else

                        porterEndsDoubleConsonent = False

                    End If

                End Function

                Private Shared Function porterEndsWith(str As String, ends As String) As Boolean

                    On Error Resume Next

                    'declaring local variables
                    Dim length_str As Byte
                    Dim length_ends As Byte
                    Dim hold_ends As String

                    'finding the length of the string
                    length_str = Len(str)
                    length_ends = Len(ends)

                    'if length of str is greater than the length of length_ends, only then proceed..else return false
                    If length_ends >= length_str Then

                        porterEndsWith = False
                    Else

                        'extract characters from right of str
                        hold_ends = Right(str, length_ends)

                        'comparing to see whether hold_ends=ends
                        porterEndsWith = If(StrComp(hold_ends, ends) = 0, True, False)

                    End If

                End Function

                Private Shared Function porterTrimEnd(str As String, length As Byte) As String

                    On Error Resume Next

                    'returning the trimmed string
                    porterTrimEnd = Left(str, Len(str) - length)

                End Function

                Private Shared Function returnCVCpattern(str As String) As String

                    'local variables
                    Dim chars() As Byte
                    Dim const_vowel As String = ""
                    Dim i As Byte

                    'converting string to byte array
                    chars = System.Text.Encoding.Unicode.GetBytes(str)

                    'checking each character to see if it is a consonent or a vowel. also inputs the information in const_vowel
                    For i = 0 To UBound(chars) Step 1

                        If Chr(chars(i)) = "a" Or Chr(chars(i)) = "e" Or Chr(chars(i)) = "i" Or Chr(chars(i)) = "o" Or Chr(chars(i)) = "u" Then
                            const_vowel = const_vowel + "v"
                        ElseIf Chr(chars(i)) = "y" Then
                            'if y is not the first character, only then check the previous character
                            'check to see if previous character is a consonent
                            const_vowel = If(i > 0, If(Not (Chr(chars(i - 1)) = "a" Or Chr(chars(i - 1)) = "e" Or Chr(chars(i - 1)) = "i" Or Chr(chars(i - 1)) = "o" Or Chr(chars(i - 1)) = "u"), const_vowel + "v", const_vowel + "c"), const_vowel + "c")
                        Else
                            const_vowel = const_vowel + "c"
                        End If

                    Next i

                    returnCVCpattern = const_vowel

                End Function

                Public Shared Function StemWord(str As String) As String

                    'only strings greater than 2 are stemmed
                    If Len(Trim(str)) > 0 Then
                        str = porterAlgorithmStep1(str)
                        str = porterAlgorithmStep2(str)
                        str = porterAlgorithmStep3(str)
                        str = porterAlgorithmStep4(str)
                        str = porterAlgorithmStep5(str)
                    End If

                    'End of Porter's algorithm.........returning the word
                    StemWord = str

                End Function

            End Class

            ''' <summary>
            ''' This is a list of Known SentenceStructures stored in db ; Each sentence which has
            ''' been sucessfully anylasied for parts of speech; the structure is saved; as a list of string
            ''' </summary>
            Public Structure Taglist
                Public Sentence As String
                Public TagList As List(Of String)

                Public Function ToJson() As String
                    Dim Converter As New JavaScriptSerializer
                    Return Converter.Serialize(Me)

                End Function

            End Structure

            ''' <summary>
            ''' Detailed Word information
            ''' </summary>
            ''' <remarks></remarks>
            Public Structure Word

                Public Function ToJson() As String
                    Dim Converter As New JavaScriptSerializer
                    Return Converter.Serialize(Me)

                End Function

                Public Function Join(ByVal Word1 As Word, ByRef Word2 As Word) As Word
                    Join = New Word
                    Join.POS &= Word1.POS & " " & Word2.POS
                    Join.Word &= Word1.Word & " " & Word2.Word
                    Return Word1
                End Function

                Public frequecy As Integer
                Public POS As String
                Public Position As Integer
                Public Word As String
                Public Catagory As String

                Public Overrides Function ToString() As String
                    Dim Str As String = ""
                    Str &= frequecy.ToString() & vbCrLf
                    Str &= POS.ToString() & vbCrLf
                    Str &= Position.ToString() & vbCrLf
                    Str &= Me.Word.ToString() & vbCrLf

                    Return Str

                End Function

                'Word
                'Position in sentence
                'Part of Speech Tag
                'Frequency of Word
            End Structure

            ''' <summary>
            ''' Used to denote tense
            ''' </summary>
            ''' <remarks></remarks>
            Public Enum tense
                Past = 0
                Present = 1
                Future = 2
                Unknown = 3
            End Enum

            ''' <summary>
            ''' Used to output infomation from the Class
            ''' </summary>
            Public Structure PosSentence

                Public Function ToJson() As String
                    Dim Converter As New JavaScriptSerializer
                    Return Converter.Serialize(Me)

                End Function

                Public Overrides Function ToString() As String
                    Dim Str As String = ""

                    Str &= vbNewLine & vbNewLine & "ACTIONS :" & vbNewLine & Me._Actions.ToJson
                    Str &= vbNewLine & vbNewLine & "NOUN_PHRASES :" & vbNewLine & Me._NounPhrases.ToJson
                    Str &= vbNewLine & vbNewLine & "PEOPLE :" & vbNewLine & Me._People.ToJson
                    Str &= vbNewLine & vbNewLine & "NUMBER_OF_WORDS :" & vbNewLine & Me._NumberOfWords.ToJson
                    Str &= vbNewLine & vbNewLine & "SUBJECTS :" & vbNewLine & Me._Subjects.ToJson
                    Str &= vbNewLine & vbNewLine & "VERB_PHRASES :" & vbNewLine & Me._VerbPhrases.ToJson
                    Select Case Me._tense
                        Case tense.Past
                            Str &= vbNewLine & vbNewLine & "TENSE :" & "Past"
                        Case tense.Present
                            Str &= vbNewLine & vbNewLine & "TENSE :" & "Present"
                        Case tense.Future
                            Str &= vbNewLine & vbNewLine & "TENSE :" & "Future"

                    End Select
                    Str &= vbNewLine & vbNewLine & "WORDS :" & vbNewLine & Me._Words.ToJson

                    Return Str
                End Function

                Dim _Actions As List(Of Word)
                Dim _NounPhrases As List(Of Word)
                Dim _NumberOfWords As Integer
                Dim _People As List(Of Word)
                Dim _Subjects As List(Of Word)
                Dim _tense As tense
                Dim _VerbPhrases As List(Of Word)
                Dim _Words As List(Of Word)

            End Structure

            ''' <summary>
            ''' Used to hold catagorical data retrieved from the DB
            ''' </summary>
            Public Structure CatagoricalWord
                Dim WordType As String
                Dim Pos As String
                Dim Word As String
                Dim Catagory As String

                Public Function GetRandomPattern(ByRef Patterns As List(Of CatagoricalWord)) As CatagoricalWord
                    Dim rnd = New Random()
                    If Patterns.Count > 0 Then

                        Return Patterns(rnd.Next(0, Patterns.Count))
                    Else
                        Return Nothing
                    End If
                End Function

            End Structure

            Public Module PartsOfSpeechExtensions

                <Runtime.CompilerServices.Extension()>
                Public Function GetRandomObjectFromList(ByRef Patterns As List(Of Object)) As Object
                    Dim rnd = New Random()
                    If Patterns.Count > 0 Then

                        Return Patterns(rnd.Next(0, Patterns.Count))
                    Else
                        Return Nothing
                    End If
                End Function

                ''' <summary>
                ''' Returns frequencys for words
                ''' </summary>
                ''' <param name="_Text"></param>
                ''' <param name="Delimiter"></param>
                ''' <returns></returns>
                <Runtime.CompilerServices.Extension()>
                Public Function GetWordFrequecys(ByVal _Text As String, ByVal Delimiter As String) As List(Of WordFrequecys)
                    Dim Words As New WordFrequecys
                    Dim ListOfWordFrequecys As New List(Of WordFrequecys)
                    Dim WordList As List(Of String) = _Text.Split(Delimiter).ToList
                    Dim groups = WordList.GroupBy(Function(value) value)
                    For Each grp In groups
                        Words.word = grp(0)
                        Words.frequency = grp.Count
                        ListOfWordFrequecys.Add(Words)
                    Next
                    Return ListOfWordFrequecys
                End Function

                ''' <summary>
                ''' THE STRING
                ''' RETURNED BY THE FUNCTION IS THE STEMMED WORD
                ''' Porter Stemmer.
                ''' </summary>
                ''' <param name="Word"></param>
                ''' <returns></returns>
                <Runtime.CompilerServices.Extension()>
                Public Function StemWord(ByRef Word As String) As String
                    Return WordStemmer.StemWord(Word)
                End Function

                ''' <summary>
                ''' counts occurrences of a specific phoneme
                ''' </summary>
                ''' <param name="strIn">  </param>
                ''' <param name="strFind"></param>
                ''' <returns></returns>
                ''' <remarks></remarks>
                <Runtime.CompilerServices.Extension()>
                Public Function CountOccurrences(ByRef strIn As String, ByRef strFind As String) As Integer
                    '**
                    ' Returns: the number of times a string appears in a string
                    '
                    '@rem           Example code for CountOccurrences()
                    '
                    '  ' Counts the occurrences of "ow" in the supplied string.
                    '
                    '    strTmp = "How now, brown cow"
                    '    Returns a value of 4
                    '
                    '
                    'Debug.Print "CountOccurrences(): there are " &  CountOccurrences(strTmp, "ow") &
                    '" occurrences of 'ow'" &    " in the string '" & strTmp & "'"
                    '
                    '@param        strIn Required. String.
                    '@param        strFind Required. String.
                    '@return       Long.

                    Dim lngPos As Integer
                    Dim lngWordCount As Integer

                    On Error GoTo PROC_ERR

                    lngWordCount = 1

                    ' Find the first occurrence
                    lngPos = InStr(strIn, strFind)

                    Do While lngPos > 0
                        ' Find remaining occurrences
                        lngPos = InStr(lngPos + 1, strIn, strFind)
                        If lngPos > 0 Then
                            ' Increment the hit counter
                            lngWordCount = lngWordCount + 1
                        End If
                    Loop

                    ' Return the value
                    CountOccurrences = lngWordCount

PROC_EXIT:
                    Exit Function

PROC_ERR:
                    MsgBox("Error: " & Err.Number & ". " & Err.Description, , NameOf(CountOccurrences))
                    Resume PROC_EXIT

                End Function

                ''' <summary>
                ''' each charcter can be defined as a particular token enabling for removal of unwanted
                ''' token types
                ''' </summary>
                ''' <param name="CharStr"></param>
                ''' <returns></returns>
                <Runtime.CompilerServices.Extension()>
                Public Function GetTokenType(ByRef CharStr As String) As TokenType
                    For Each item In SeperatorPunctuation
                        If CharStr = item Then Return TokenType.SeperatorPunctuation
                    Next
                    For Each item In GramaticalPunctuation
                        If CharStr = item Then Return TokenType.GramaticalPunctuation
                    Next
                    For Each item In EncapuslationPunctuationStart
                        If CharStr = item Then Return TokenType.EncapuslationPunctuationStart
                    Next
                    For Each item In EncapuslationPunctuationEnd
                        If CharStr = item Then Return TokenType.EncapuslationPunctuationEnd
                    Next
                    For Each item In MoneyPunctuation
                        If CharStr = item Then Return TokenType.MoneyPunctuation
                    Next
                    For Each item In MathPunctuation
                        If CharStr = item Then Return TokenType.MathPunctuation
                    Next
                    For Each item In CodePunctuation
                        If CharStr = item Then Return TokenType.CodePunctuation
                    Next
                    For Each item In AlphaBet
                        If CharStr = item Then Return TokenType.AlphaBet
                    Next
                    For Each item In Number
                        If CharStr = item Then Return TokenType.Number
                    Next
                    Return TokenType.Ignore
                End Function

            End Module

            Public Structure WordFrequecys

                Public Overrides Function ToString() As String
                    Dim Str As String = ""
                    Str = Str & "Word :" & word & vbCrLf
                    Str = Str & "frequency :" & frequency & vbCrLf
                    Return Str
                End Function

                Public frequency As Integer
                Public word As String
            End Structure

        End Namespace

        Namespace Tokens

            ''' <summary>
            ''' Tokenizer For NLP Techniques
            ''' </summary>
            <ComClass(Tokenizer.ClassId, Tokenizer.InterfaceId, Tokenizer.EventsId)>
            Public Class Tokenizer
                Public Const ClassId As String = "2899E490-7702-401C-BAB3-38FF97BC1AC9"
                Public Const EventsId As String = "CD994307-F53E-401A-AC6D-3CFDD86FD6F1"
                Public Const InterfaceId As String = "8B1145F1-5D13-4059-829B-B531310144B5"

                'Tokenizer
                ''' <summary>
                ''' Returns Characters in String as list
                ''' </summary>
                ''' <param name="InputStr"></param>
                ''' <returns></returns>
                Public Function Tokenizer(ByRef InputStr As String) As List(Of String)
                    Tokenizer = New List(Of String)
                    InputStr = GetValidTokens(InputStr)

                    Dim Endstr As Integer = InputStr.Length
                    For i = 1 To Endstr
                        Tokenizer.Add(InputStr(i - 1))
                    Next
                End Function

                ''' <summary>
                ''' Return Tokens in string divided by seperator
                ''' </summary>
                ''' <param name="InputStr"></param>
                ''' <param name="Divider">  </param>
                ''' <returns></returns>
                Public Function Tokenizer(ByRef InputStr As String, ByRef Divider As String) As List(Of String)
                    Tokenizer = New List(Of String)
                    InputStr = GetValidTokens(InputStr)
                    Dim Tokens() As String = InputStr.Split(Divider)

                    For Each item In Tokens
                        Tokenizer.Add(item)
                    Next
                End Function

                ''' <summary>
                ''' Gets Frequency of token
                ''' </summary>
                ''' <param name="Token"></param>
                ''' <param name="InputStr"></param>
                ''' <returns></returns>
                Public Shared Function GetTokenFrequency(ByRef Token As String, ByRef InputStr As String) As Integer
                    GetTokenFrequency = 0
                    InputStr = GetValidTokens(InputStr)
                    If InputStr.Contains(Token) = True Then
                        For Each item In GetWordFrequecys(InputStr, " ")
                            If item.word = Token Then
                                GetTokenFrequency = item.frequency
                            End If
                        Next
                    End If
                End Function

                Private Shared Function GetWordFrequecys(ByVal _Text As String, ByVal Delimiter As String) As List(Of WordFrequecys)
                    Dim Words As New WordFrequecys
                    Dim TempArray() As String = _Text.Split(Delimiter)
                    Dim WordList As New List(Of String)
                    Dim ListOfWordFrequecys As New List(Of WordFrequecys)
                    For Each word As String In TempArray
                        WordList.Add(word)
                    Next
                    Dim groups = WordList.GroupBy(Function(value) value)
                    For Each grp In groups
                        Words.word = grp(0)
                        Words.frequency = grp.Count
                        ListOfWordFrequecys.Add(Words)
                    Next
                    Return ListOfWordFrequecys
                End Function

                ''' <summary>
                ''' each charcter can be defined as a particular token enabling for removal of unwanted token types
                ''' </summary>
                ''' <param name="CharStr"></param>
                ''' <returns></returns>
                Public Shared Function GetTokenType(ByRef CharStr As String) As TokenType
                    For Each item In SeperatorPunctuation
                        If CharStr = item Then Return TokenType.SeperatorPunctuation
                    Next
                    For Each item In GramaticalPunctuation
                        If CharStr = item Then Return TokenType.GramaticalPunctuation
                    Next
                    For Each item In EncapuslationPunctuationStart
                        If CharStr = item Then Return TokenType.EncapuslationPunctuationStart
                    Next
                    For Each item In EncapuslationPunctuationEnd
                        If CharStr = item Then Return TokenType.EncapuslationPunctuationEnd
                    Next
                    For Each item In MoneyPunctuation
                        If CharStr = item Then Return TokenType.MoneyPunctuation
                    Next
                    For Each item In MathPunctuation
                        If CharStr = item Then Return TokenType.MathPunctuation
                    Next
                    For Each item In CodePunctuation
                        If CharStr = item Then Return TokenType.CodePunctuation
                    Next
                    For Each item In AlphaBet
                        If CharStr = item Then Return TokenType.AlphaBet
                    Next
                    For Each item In Number
                        If CharStr = item Then Return TokenType.Number
                    Next
                    Return TokenType.Ignore
                End Function

                ''' <summary>
                ''' Returns valid tokens only tokens that are not defined are removed
                ''' </summary>
                ''' <param name="InputStr"></param>
                ''' <returns></returns>
                Private Shared Function GetValidTokens(ByRef InputStr As String) As String
                    Dim EndStr As Integer = InputStr.Length
                    Dim CharStr As String = ""
                    For i = 0 To EndStr - 1
                        If GetTokenType(InputStr(i)) <> TokenType.Ignore Then
                            CharStr = CharStr.AddSuffix(InputStr(i))
                        Else

                        End If
                    Next
                    Return CharStr
                End Function

                Public Shared Function AlphanumericOnly(ByRef Str As String) As String
                    Str = Str.GetValidTokens
                    Str = RemoveTokenType(Str, TokenType.CodePunctuation)
                    Str = RemoveTokenType(Str, TokenType.EncapuslationPunctuationEnd)
                    Str = RemoveTokenType(Str, TokenType.EncapuslationPunctuationStart)
                    Str = RemoveTokenType(Str, TokenType.MathPunctuation)
                    Str = RemoveTokenType(Str, TokenType.MoneyPunctuation)
                    Str = RemoveTokenType(Str, TokenType.GramaticalPunctuation)
                    Str = Str.Remove(",")
                    Str = Str.Remove("|")
                    Str = Str.Remove("_")
                    Return Str
                End Function

                ''' <summary>
                ''' Removes Tokens From String by Type
                ''' </summary>
                ''' <param name="UserStr"></param>
                ''' <param name="nType">  </param>
                ''' <returns></returns>
                Public Shared Function RemoveTokenType(ByRef UserStr As String, ByRef nType As TokenType) As String

                    Select Case nType
                        Case TokenType.GramaticalPunctuation
                            For Each item In GramaticalPunctuation
                                If UCase(UserStr).Contains(UCase(item)) = True Then
                                    UserStr = UCase(UserStr).Remove(UCase(item))
                                End If
                            Next
                        Case TokenType.AlphaBet
                            For Each item In AlphaBet
                                If UCase(UserStr).Contains(UCase(item)) = True Then
                                    UserStr = UCase(UserStr).Remove(UCase(item))
                                End If
                            Next
                        Case TokenType.CodePunctuation
                            For Each item In CodePunctuation
                                If UCase(UserStr).Contains(UCase(item)) = True Then
                                    UserStr = UCase(UserStr).Remove(UCase(item))
                                End If
                            Next
                        Case TokenType.EncapuslationPunctuationEnd
                            For Each item In EncapuslationPunctuationEnd
                                If UCase(UserStr).Contains(UCase(item)) = True Then
                                    UserStr = UCase(UserStr).Remove(UCase(item))
                                End If
                            Next
                        Case TokenType.EncapuslationPunctuationStart
                            For Each item In EncapuslationPunctuationStart
                                If UCase(UserStr).Contains(UCase(item)) = True Then
                                    UserStr = UCase(UserStr).Remove(UCase(item))
                                End If
                            Next
                        Case TokenType.Ignore
                        Case TokenType.MathPunctuation
                            For Each item In MathPunctuation
                                If UCase(UserStr).Contains(UCase(item)) = True Then
                                    UserStr = UCase(UserStr).Remove(UCase(item))
                                End If
                            Next
                        Case TokenType.MoneyPunctuation
                            For Each item In MoneyPunctuation
                                If UCase(UserStr).Contains(UCase(item)) = True Then
                                    UserStr = UCase(UserStr).Remove(UCase(item))
                                End If
                            Next
                        Case TokenType.Number
                            For Each item In Number
                                If UCase(UserStr).Contains(UCase(item)) = True Then
                                    UserStr = UCase(UserStr).Remove(UCase(item))
                                End If
                            Next
                        Case TokenType.SeperatorPunctuation
                            For Each item In SeperatorPunctuation
                                If UCase(UserStr).Contains(UCase(item)) = True Then
                                    UserStr = UCase(UserStr).Remove(UCase(item))
                                End If
                            Next

                    End Select
                    Return UserStr
                End Function

                'Form Extensions
                ''' <summary>
                ''' Counts tokens in string
                ''' </summary>
                ''' <param name="InputStr"></param>
                ''' <param name="Delimiter"></param>
                ''' <returns></returns>
                Public Shared Function CountTokens(ByRef InputStr As String, ByRef Delimiter As String) As Integer
                    Dim Words() As String = Split(InputStr, Delimiter)
                    Return Words.Count
                End Function

                ''' <summary>
                ''' Checks if input contains Ecapuslation Punctuation
                ''' </summary>
                ''' <param name="Userinput"></param>
                ''' <returns></returns>
                Public Function ContainsEncapsulated(ByRef Userinput As String) As Boolean
                    Dim Start = False
                    Dim Ending = False
                    ContainsEncapsulated = False
                    For Each item In EncapuslationPunctuationStart
                        If Userinput.Contains(item) = True Then Start = True
                    Next
                    For Each item In EncapuslationPunctuationEnd
                        If Userinput.Contains(item) = True Then Ending = True
                    Next
                    If Start And Ending = True Then
                        ContainsEncapsulated = True
                    End If
                End Function

                ''' <summary>
                ''' Gets encapsulated items found in userinput
                ''' </summary>
                ''' <param name="USerinput"></param>
                ''' <returns></returns>
                Public Function GetEncapsulated(ByRef Userinput As String) As List(Of String)
                    GetEncapsulated = New List(Of String)
                    Do Until ContainsEncapsulated(Userinput) = False
                        GetEncapsulated.Add(ExtractEncapsulated(Userinput))
                    Loop
                End Function

                ''' <summary>
                ''' Extracts first Encapsulated located in string
                ''' </summary>
                ''' <param name="Userinput"></param>
                ''' <returns></returns>
                Public Function ExtractEncapsulated(ByRef Userinput As String) As String
                    ExtractEncapsulated = Userinput
                    If ContainsEncapsulated(ExtractEncapsulated) = True Then
                        If ExtractEncapsulated.Contains("(") = True And ExtractEncapsulated.Contains(")") = True Then
                            ExtractEncapsulated = ExtractEncapsulated.ExtractStringBetween("(", ")")
                        End If
                        If Userinput.Contains("[") = True And Userinput.Contains("]") = True Then
                            ExtractEncapsulated = ExtractEncapsulated.ExtractStringBetween("[", "]")
                        End If
                        If Userinput.Contains("{") = True And Userinput.Contains("}") = True Then
                            ExtractEncapsulated = ExtractEncapsulated.ExtractStringBetween("{", "}")
                        End If
                        If Userinput.Contains("<") = True And Userinput.Contains(">") = True Then
                            ExtractEncapsulated = ExtractEncapsulated.ExtractStringBetween("<", ">")
                        End If
                    End If
                End Function

            End Class

            ''' <summary>
            ''' Extensions for Creating a Term Document Matrix
            ''' </summary>
            Public Module TermDocumentMatrix

                ''' <summary>
                ''' Ngrams are Words grouped in to sets, they can be as long as the senence permits
                ''' </summary>
                ''' <param name="Str">String to be aynalized</param>
                ''' <param name="Ngrams">Size of token component Group (Must be &gt; 0)</param>
                ''' <returns></returns>
                <Runtime.CompilerServices.Extension()>
                Public Function GetNgrms(ByRef Str As String, ByRef Ngrams As Integer) As List(Of String)
                    Dim NgramArr() As String = Split(Str, " ")
                    Dim Length As Integer = NgramArr.Count
                    Dim lst As New List(Of String)
                    Dim Str2 As String = ""
                    If Length - Ngrams > 0 Then

                        For i = 0 To Length - Ngrams
                            Str2 = ""
                            Dim builder As New System.Text.StringBuilder()
                            builder.Append(Str2)
                            For j = 0 To Ngrams - 1
                                builder.Append(NgramArr(i + j) & " ")
                            Next
                            Str2 = builder.ToString()
                            lst.Add(Str2)
                        Next
                    Else

                    End If
                    Return lst
                End Function

                'Document Matrix
                ''' <summary>
                ''' TF-IDF also offsets this value by the frequency of the term in the entire document set,
                ''' a value called Inverse Document Frequency. in this case a single document TF(t) =
                ''' (Number of times term t appears in a document) / (Total number of terms in the document)
                ''' IDF(t) = log_e(Total number of documents / Number of documents with term t in it). Value
                ''' = TF * IDF
                '''
                ''' TF-IDF is computed for each term in each document. Typically, you will be interested
                ''' either in one term in particular (like a search engine). or you would be interested in
                ''' the terms with the highest TF-IDF in a specific document.
                ''' </summary>
                ''' <param name="Word"></param>
                ''' <param name="Document"></param>
                ''' <returns></returns>
                <System.Runtime.CompilerServices.Extension>
                Public Function TermFrequencyInverseDocumentFrequency(ByRef Word As String, ByRef Document As String) As Integer
                    'TF * IDF
                    Dim DF = 1

                    Dim tf As Integer = Document.TermFrequency(Word)

                    Return DF / tf
                End Function

                <System.Runtime.CompilerServices.Extension>
                Public Function TermFrequency(ByRef Userinput As String, ByRef Term As String) As Double
                    Return Term.GetTokenFrequency(Userinput)
                End Function

                ''' <summary>
                ''' Counts tokens in string
                ''' </summary>
                ''' <param name="Userinput"></param>
                ''' <param name="Delimiter"></param>
                ''' <returns></returns>
                <System.Runtime.CompilerServices.Extension>
                Public Function CountTokens(ByRef Userinput As String, ByRef Delimiter As String) As Integer
                    Dim Words() As String = Split(Userinput, Delimiter)
                    Return Words.Count
                End Function

                ''' <summary>
                ''' TF-IDF also offsets this value by the frequency of the term in the entire document set,
                ''' a value called Inverse Document Frequency. TF(t) = (Number of times term t appears in a
                ''' document) / (Total number of terms in the document) IDF(t) = log_e(Total number of documents
                ''' / Number of documents with term t in it). Value = TF * IDF
                '''
                ''' TF-IDF is computed for each term in each document. Typically, you will be interested
                ''' either in one term in particular (like a search engine). or you would be interested in
                ''' the terms with the highest TF-IDF in a specific document.
                ''' </summary>
                ''' <param name="Word"></param>
                ''' <param name="Document"></param>
                ''' <param name="Documents"></param>
                ''' <returns></returns>
                <System.Runtime.CompilerServices.Extension>
                Public Function TermFrequencyInverseDocumentFrequency(ByRef Word As String, ByRef Document As String,
                                                      ByRef Documents As List(Of String)) As Integer
                    'TF * IDF

                    Dim tf As Integer = Document.TermFrequency(Word)
                    Dim IDF As Integer = Word.InverseTermDocumentFrequency(Documents)

                    Return tf * IDF
                End Function

                ''' <summary>
                ''' TDF(t) = Total number of documents / Number of documents with term t in it
                ''' </summary>
                ''' <param name="Word"></param>
                ''' <param name="Documents"></param>
                ''' <returns></returns>
                <System.Runtime.CompilerServices.Extension>
                Public Function TermDocumentFrequency(ByRef Word As String, ByRef Documents As List(Of String)) As Integer
                    TermDocumentFrequency = 0
                    Dim Freqs As New List(Of Integer)
                    Dim Count As Integer = 0
                    For Each item In Documents
                        Freqs.Add(item.TermFrequency(Word))
                        If item.Contains(Word) = True Then
                            Count += 1
                        End If
                    Next
                    TermDocumentFrequency = Documents.Count / Count
                End Function

                ''' <summary>
                ''' IDF: Inverse Document Frequency, which measures how important a term Is. While computing TF,
                ''' all terms are considered equally important. However it Is known that certain terms, such
                ''' As "is", "of", And "that", may appear a lot Of times but have little importance. Thus we
                ''' need To weigh down the frequent terms While scale up the rare ones, by computing the
                ''' following: IDF(t) = log_e(Total number of documents / Number of documents with term t in it).
                ''' </summary>
                ''' <param name="Userinput"></param>
                ''' <param name="Term"></param>
                ''' <returns></returns>
                <System.Runtime.CompilerServices.Extension>
                Public Function InverseTermFrequency(ByRef Userinput As String, ByRef Term As String) As Double
                    Return Math.Log(GetTokenFrequency(Term, Userinput) / Userinput.CountTokens(" "))
                End Function

                ''' <summary>
                ''' IDF(t) = log_e(Total number of documents / Number of documents with term t in it).
                ''' </summary>
                ''' <param name="Word"></param>
                ''' <param name="Documents"></param>
                ''' <returns></returns>
                <System.Runtime.CompilerServices.Extension>
                Public Function InverseTermDocumentFrequency(ByRef Word As String, ByRef Documents As List(Of String)) As Integer
                    InverseTermDocumentFrequency = 0
                    Dim Freqs As New List(Of Integer)
                    Dim Count As Integer = 0
                    For Each item In Documents
                        Freqs.Add(item.TermFrequency(Word))
                        If item.Contains(Word) = True Then
                            Count += 1
                        End If
                    Next
                    InverseTermDocumentFrequency = Math.Log(Documents.Count / Count)
                End Function

            End Module

            ''' <summary>
            ''' Recognized Tokens
            ''' </summary>
            Public Enum TokenType
                GramaticalPunctuation
                EncapuslationPunctuationStart
                EncapuslationPunctuationEnd
                MoneyPunctuation
                MathPunctuation
                CodePunctuation
                AlphaBet
                Number
                SeperatorPunctuation
                Ignore
            End Enum

            ''' <summary>
            ''' Extension Methods for tokeninzer Fucntions
            ''' </summary>
            Public Module Tokenizer_Extensions

                ''' <summary>
                ''' Checks if string is a reserved VBscipt Keyword
                ''' </summary>
                ''' <param name="keyword"></param>
                ''' <returns></returns>
                <Runtime.CompilerServices.Extension()>
                Function IsReservedWord(ByVal keyword As String) As Boolean
                    Dim IsReserved = False
                    Select Case LCase(keyword)
                        Case "and" : IsReserved = True
                        Case "as" : IsReserved = True
                        Case "boolean" : IsReserved = True
                        Case "byref" : IsReserved = True
                        Case "byte" : IsReserved = True
                        Case "byval" : IsReserved = True
                        Case "call" : IsReserved = True
                        Case "case" : IsReserved = True
                        Case "class" : IsReserved = True
                        Case "const" : IsReserved = True
                        Case "currency" : IsReserved = True
                        Case "debug" : IsReserved = True
                        Case "dim" : IsReserved = True
                        Case "do" : IsReserved = True
                        Case "double" : IsReserved = True
                        Case "each" : IsReserved = True
                        Case "else" : IsReserved = True
                        Case "elseif" : IsReserved = True
                        Case "empty" : IsReserved = True
                        Case "end" : IsReserved = True
                        Case "endif" : IsReserved = True
                        Case "enum" : IsReserved = True
                        Case "eqv" : IsReserved = True
                        Case "event" : IsReserved = True
                        Case "exit" : IsReserved = True
                        Case "false" : IsReserved = True
                        Case "for" : IsReserved = True
                        Case "function" : IsReserved = True
                        Case "get" : IsReserved = True
                        Case "goto" : IsReserved = True
                        Case "if" : IsReserved = True
                        Case "imp" : IsReserved = True
                        Case "implements" : IsReserved = True
                        Case "in" : IsReserved = True
                        Case "integer" : IsReserved = True
                        Case "is" : IsReserved = True
                        Case "let" : IsReserved = True
                        Case "like" : IsReserved = True
                        Case "long" : IsReserved = True
                        Case "loop" : IsReserved = True
                        Case "lset" : IsReserved = True
                        Case "me" : IsReserved = True
                        Case "mod" : IsReserved = True
                        Case "new" : IsReserved = True
                        Case "next" : IsReserved = True
                        Case "not" : IsReserved = True
                        Case "nothing" : IsReserved = True
                        Case "null" : IsReserved = True
                        Case "on" : IsReserved = True
                        Case "option" : IsReserved = True
                        Case "optional" : IsReserved = True
                        Case "or" : IsReserved = True
                        Case "paramarray" : IsReserved = True
                        Case "preserve" : IsReserved = True
                        Case "private" : IsReserved = True
                        Case "public" : IsReserved = True
                        Case "raiseevent" : IsReserved = True
                        Case "redim" : IsReserved = True
                        Case "rem" : IsReserved = True
                        Case "resume" : IsReserved = True
                        Case "rset" : IsReserved = True
                        Case "select" : IsReserved = True
                        Case "set" : IsReserved = True
                        Case "shared" : IsReserved = True
                        Case "single" : IsReserved = True
                        Case "static" : IsReserved = True
                        Case "stop" : IsReserved = True
                        Case "sub" : IsReserved = True
                        Case "then" : IsReserved = True
                        Case "to" : IsReserved = True
                        Case "true" : IsReserved = True
                        Case "type" : IsReserved = True
                        Case "typeof" : IsReserved = True
                        Case "until" : IsReserved = True
                        Case "variant" : IsReserved = True
                        Case "wend" : IsReserved = True
                        Case "while" : IsReserved = True
                        Case "with" : IsReserved = True
                        Case "xor" : IsReserved = True
                    End Select
                    Return IsReserved
                End Function

                Public ReadOnly AlphaBet() As String = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N",
        "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n",
        "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z"}

                Public ReadOnly CodePunctuation() As String = {"\", "#", "@", "^"}
                Public ReadOnly EncapuslationPunctuationEnd() As String = {"}", "]", ">", ")"}
                Public ReadOnly EncapuslationPunctuationStart() As String = {"{", "[", "<", "("}
                Public ReadOnly GramaticalPunctuation() As String = {".", "?", "!", ":", ";"}
                Public ReadOnly MathPunctuation() As String = {"+", "-", "=", "/", "*", "%", "PLUS", "ADD", "MINUS", "SUBTRACT", "DIVIDE", "DIFFERENCE", "TIMES", "MULTIPLY", "PERCENT", "EQUALS"}
                Public ReadOnly MoneyPunctuation() As String = {"£", "$"}

                Public ReadOnly Number() As String = {"1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20",
    "30", "40", "50", "60", "70", "80", "90", "00", "000", "0000", "00000", "000000", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen",
    "nineteen", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety", "hundred", "thousand", "million", "Billion"}

                Public ReadOnly SeperatorPunctuation() = {" ", ",", "|"}

                ' these are the common word delimiters
                Public Delimiters() As Char = {CType(" ", Char), CType(".", Char),
                    CType(",", Char), CType("?", Char),
                    CType("!", Char), CType(";", Char),
                    CType(":", Char), Chr(10), Chr(13), vbTab}

                'Form Extensions
                ''' <summary>
                ''' Counts tokens in string
                ''' </summary>
                ''' <param name="InputStr"></param>
                ''' <param name="Delimiter"></param>
                ''' <returns></returns>
                <System.Runtime.CompilerServices.Extension>
                Public Function CountTokens(ByRef InputStr As String, ByRef Delimiter As String) As Integer
                    Dim Words() As String = Split(InputStr, Delimiter)
                    Return Words.Count
                End Function

                <Runtime.CompilerServices.Extension()>
                Public Function AlphanumericOnly(ByRef Str As String) As String
                    Str = Str.GetValidTokens
                    Str = RemoveTokenType(Str, TokenType.CodePunctuation)
                    Str = RemoveTokenType(Str, TokenType.EncapuslationPunctuationEnd)
                    Str = RemoveTokenType(Str, TokenType.EncapuslationPunctuationStart)
                    Str = RemoveTokenType(Str, TokenType.MathPunctuation)
                    Str = RemoveTokenType(Str, TokenType.MoneyPunctuation)
                    Str = RemoveTokenType(Str, TokenType.GramaticalPunctuation)
                    Str = Str.Remove(",")
                    Str = Str.Remove("|")
                    Str = Str.Remove("_")
                    Return Str
                End Function

                ''' <summary>
                ''' Removes Tokens From String by Type
                ''' </summary>
                ''' <param name="UserStr"></param>
                ''' <param name="nType"></param>
                ''' <returns></returns>
                <Runtime.CompilerServices.Extension()>
                Public Function RemoveTokenType(ByRef UserStr As String, ByRef nType As TokenType) As String

                    Select Case nType
                        Case TokenType.GramaticalPunctuation
                            For Each item In GramaticalPunctuation
                                If UCase(UserStr).Contains(UCase(item)) = True Then
                                    UserStr = UCase(UserStr).Remove(UCase(item))
                                End If
                            Next
                        Case TokenType.AlphaBet
                            For Each item In AlphaBet
                                If UCase(UserStr).Contains(UCase(item)) = True Then
                                    UserStr = UCase(UserStr).Remove(UCase(item))
                                End If
                            Next
                        Case TokenType.CodePunctuation
                            For Each item In CodePunctuation
                                If UCase(UserStr).Contains(UCase(item)) = True Then
                                    UserStr = UCase(UserStr).Remove(UCase(item))
                                End If
                            Next
                        Case TokenType.EncapuslationPunctuationEnd
                            For Each item In EncapuslationPunctuationEnd
                                If UCase(UserStr).Contains(UCase(item)) = True Then
                                    UserStr = UCase(UserStr).Remove(UCase(item))
                                End If
                            Next
                        Case TokenType.EncapuslationPunctuationStart
                            For Each item In EncapuslationPunctuationStart
                                If UCase(UserStr).Contains(UCase(item)) = True Then
                                    UserStr = UCase(UserStr).Remove(UCase(item))
                                End If
                            Next
                        Case TokenType.Ignore
                        Case TokenType.MathPunctuation
                            For Each item In MathPunctuation
                                If UCase(UserStr).Contains(UCase(item)) = True Then
                                    UserStr = UCase(UserStr).Remove(UCase(item))
                                End If
                            Next
                        Case TokenType.MoneyPunctuation
                            For Each item In MoneyPunctuation
                                If UCase(UserStr).Contains(UCase(item)) = True Then
                                    UserStr = UCase(UserStr).Remove(UCase(item))
                                End If
                            Next
                        Case TokenType.Number
                            For Each item In Number
                                If UCase(UserStr).Contains(UCase(item)) = True Then
                                    UserStr = UCase(UserStr).Remove(UCase(item))
                                End If
                            Next
                        Case TokenType.SeperatorPunctuation
                            For Each item In SeperatorPunctuation
                                If UCase(UserStr).Contains(UCase(item)) = True Then
                                    UserStr = UCase(UserStr).Remove(UCase(item))
                                End If
                            Next

                    End Select
                    Return UserStr
                End Function

                ''' <summary>
                ''' Returns Characters in String as list
                ''' </summary>
                ''' <param name="InputStr"></param>
                ''' <returns></returns>
                <Runtime.CompilerServices.Extension()>
                Public Function Tokenizer(ByRef InputStr As String) As List(Of String)
                    Tokenizer = New List(Of String)
                    InputStr = GetValidTokens(InputStr)

                    Dim Endstr As Integer = InputStr.Length
                    For i = 0 To Endstr
                        Tokenizer.Add(InputStr(i))
                    Next
                End Function

                ''' <summary>
                ''' Return Tokens in string divided by seperator
                ''' </summary>
                ''' <param name="InputStr"></param>
                ''' <param name="Divider"></param>
                ''' <returns></returns>
                <Runtime.CompilerServices.Extension()>
                Public Function Tokenizer(ByRef InputStr As String, ByRef Divider As String) As List(Of String)
                    Tokenizer = New List(Of String)
                    InputStr = GetValidTokens(InputStr)
                    Dim Tokens() As String = InputStr.Split(Divider)

                    For Each item In Tokens
                        Tokenizer.Add(item)
                    Next
                End Function

                <Runtime.CompilerServices.Extension()>
                Public Function GetTokenFrequency(ByRef Token As String, ByRef InputStr As String) As Integer
                    GetTokenFrequency = 0
                    InputStr = GetValidTokens(InputStr)
                    If InputStr.Contains(Token) = True Then
                        For Each item In GetWordFrequecys(InputStr, " ")
                            If item.word = Token Then
                                GetTokenFrequency = item.frequency
                            End If
                        Next
                    End If
                End Function

                Private Function GetWordFrequecys(ByVal _Text As String, ByVal Delimiter As String) As List(Of WordFrequecys)
                    Dim Words As New WordFrequecys
                    Dim TempArray() As String = _Text.Split(Delimiter)
                    Dim WordList As New List(Of String)
                    Dim ListOfWordFrequecys As New List(Of WordFrequecys)
                    For Each word As String In TempArray
                        WordList.Add(word)
                    Next
                    Dim groups = WordList.GroupBy(Function(value) value)
                    For Each grp In groups
                        Words.word = grp(0)
                        Words.frequency = grp.Count
                        ListOfWordFrequecys.Add(Words)
                    Next
                    Return ListOfWordFrequecys
                End Function

                ''' <summary>
                ''' each charcter can be defined as a particular token enabling for removal of unwanted
                ''' token types
                ''' </summary>
                ''' <param name="CharStr"></param>
                ''' <returns></returns>
                <Runtime.CompilerServices.Extension()>
                Public Function GetTokenType(ByRef CharStr As String) As TokenType
                    For Each item In SeperatorPunctuation
                        If CharStr = item Then Return TokenType.SeperatorPunctuation
                    Next
                    For Each item In GramaticalPunctuation
                        If CharStr = item Then Return TokenType.GramaticalPunctuation
                    Next
                    For Each item In EncapuslationPunctuationStart
                        If CharStr = item Then Return TokenType.EncapuslationPunctuationStart
                    Next
                    For Each item In EncapuslationPunctuationEnd
                        If CharStr = item Then Return TokenType.EncapuslationPunctuationEnd
                    Next
                    For Each item In MoneyPunctuation
                        If CharStr = item Then Return TokenType.MoneyPunctuation
                    Next
                    For Each item In MathPunctuation
                        If CharStr = item Then Return TokenType.MathPunctuation
                    Next
                    For Each item In CodePunctuation
                        If CharStr = item Then Return TokenType.CodePunctuation
                    Next
                    For Each item In AlphaBet
                        If CharStr = item Then Return TokenType.AlphaBet
                    Next
                    For Each item In Number
                        If CharStr = item Then Return TokenType.Number
                    Next
                    Return TokenType.Ignore
                End Function

                ''' <summary>
                ''' Returns valid tokens only tokens that are not defined are removed
                ''' </summary>
                ''' <param name="InputStr"></param>
                ''' <returns></returns>
                <Runtime.CompilerServices.Extension()>
                Public Function GetValidTokens(ByRef InputStr As String) As String
                    Dim EndStr As Integer = InputStr.Length
                    Dim CharStr As String = ""
                    For i = 0 To EndStr - 1
                        If GetTokenType(InputStr(i)) <> TokenType.Ignore Then
                            CharStr = CharStr.AddSuffix(InputStr(i))
                        Else

                        End If
                    Next
                    Return CharStr
                End Function

                Public ReadOnly CoOrdinatingConjunction() As String = {" and ", " but ", " or ", " for ", " nor ", " so ", " yet "}

                Public ReadOnly DependantMarker() As String = {" because ", " before ", " since ", " while ", " although ", " If "", until ", " when ", " after ", " as ", " as if "}

                'Punctuation
                Public ReadOnly EndPunctuation() As String = {".", "?", "!"}

                Public ReadOnly IndependantMarker() As String = {" therefore ", " moreover ", " thus ", " consequently ", " however ", " also ", " nevertheless "}

                Public ReadOnly InternalPunctuation() As String = {",", ";", ":"}

            End Module

        End Namespace

        Namespace Tools

            Public Class ObjectSerializer

                '<ComClass(SDK_Propositional.ClassId, SDK_Propositional.InterfaceId, SDK_Propositional.EventsId)>
                Public Const ClassId As String = "2828E720-7705-401C-BAB3-38FBA7BC1AC9"

                Public Const EventsId As String = "CDB77327-FA5E-401A-ACAD-3CF80B6BD6F1"
                Public Const InterfaceId As String = "8B7A55F8-5D13-4059-82AB-B53131B14BB5"

                ''' <summary>
                ''' Used to extract Object from XML
                ''' </summary>
                ''' <param name="FileName">XML FILENAME</param>
                ''' <param name="myObj">OBJECT RETURNED</param>
                ''' <returns>True If Sucessfull</returns>
                Public Shared Function ObjectDeserializer(ByRef FileName As String, ByRef myObj As Object) As Boolean
                    Try

                        Dim ObjItem As Object
                        Dim string_reader As New StringReader(FileName)
                        Dim xml_serializer As New XmlSerializer(myObj.GetType)
                        ObjItem = xml_serializer.Deserialize(string_reader)
                        ObjectDeserializer = True
                    Catch ex As Exception
                        ObjectDeserializer = False
                    End Try
                End Function

                ''' <summary>
                ''' Serializes object to XML
                ''' </summary>
                ''' <param name="myObj">Presented Object</param>
                ''' <returns>True if Sucessfull</returns>
                Public Shared Function ObjectSerializer(ByRef myObj As Object) As Boolean
                    Try

                        Dim string_writer As New StringWriter()
                        Dim xml_serializer As New Xml.Serialization.XmlSerializer(myObj.GetType)
                        xml_serializer.Serialize(string_writer, myObj)
                        ObjectSerializer = True
                    Catch ex As Exception
                        ObjectSerializer = False
                    End Try
                End Function

                ''' <summary>
                ''' Serialize object
                ''' </summary>
                ''' <param name="Item"></param>
                ''' <returns></returns>
                Public Shared Function ToJson(ByRef Item As Object) As String
                    Dim Converter As New JavaScriptSerializer
                    Return Converter.Serialize(Item)

                End Function

            End Class

            Public Class WebRegEx
                Dim Emoticon_string = "(?:[<>]?[:;=8][\-o\*\']?[\)\]\(\[dDpP/\:\}\{@\|\\]|[\)\]\(\[dDpP/\:\}\{@\|\\][\-o\*\']?[:;=8][<>]?|<[/\\]?3|\(?\(?\#?[>\-\^\*\+o\~][\_\.\|oO\,][<\-\^\*\+o\~][\#\;]?\)?\)?)"
                Dim PhoneNumbers As String = "(?:(?:\+?[01][\-\s.]*)?(?:[\(]?\d{3}[\-\s.\)]*)?\d{3}[\-\s.]*\d{4})"
                Dim TwitterUSerName As String = "(?:@[\w_]+)"
                Dim TwitterHashTags As String = "(?:\#+[\w_]+[\w\'_\-]*[\w_]+)"
                Dim httpGETinfo As String = "(?:\/\w+\?(?:\;?\w+\=\w+)+)"
                Dim remainingWordTypes As String = "(?:[a-z][a-z'\-_]+[a-z])|(?:[+\-]?\d+[,/.:-]\d+[+\-]?)|(?:[\w_]+)|(?:\.(?:\s*\.){1,})|(?:\S)"

                Public Shared QuotedText As String = "("".+ "")"
                Public Shared DotALL As String = "+(.|\n)*"
                Public Shared HtmlTagR As String = "<\w+>"
                Public Shared HtmlTagL As String = "<w+>"
                Public Shared HtmlTagA As String = "<.*?>"
                Public Shared HttpUrl As String = "(https?://[0-9a-zA-Z\_/-/.]*)"

                Public Shared Function RegExFindHttpUrls(ByRef Text As String) As List(Of String)
                    Dim Tag As String = "(https?://[0-9a-zA-Z\_/-/.]*)"
                    Return RegExSearch(Text, Tag)
                End Function

                Public Shared Function RegExFindQuotedTexts(ByRef Text As String) As List(Of String)
                    Dim Tag As String = "("".+ "")"
                    Return RegExSearch(Text, Tag)
                End Function

                ''' <summary>
                ''' Using RegEx for Searching
                ''' </summary>
                ''' <param name="Text">TobeSearched</param>
                ''' <returns></returns>
                Public Shared Function RegExFindDigits(ByRef Text As String) As List(Of String)
                    Dim Tag As String = "\d+"
                    Return RegExSearch(Text, Tag)
                End Function

                ''' <summary>
                ''' Strip HTML tags.Removes tags from text
                ''' </summary>
                Public Shared Function RemoveTags(ByVal html As String) As String
                    ' Remove HTML tags.
                    Return Regex.Replace(html, "<.*?>", "")
                End Function

                ''' <summary>
                ''' Removes just the outer braces from text
                ''' </summary>
                ''' <param name="html"></param>
                ''' <returns></returns>
                Public Shared Function RemoveBraces(ByVal html As String) As String
                    ' Remove HTML tags.
                    html = Regex.Replace(html, "<", "")
                    html = Regex.Replace(html, ">", "")
                    Return html
                End Function

                ''' <summary>
                ''' removes both sides of the tag from the text
                ''' </summary>
                ''' <param name="html">text to be cleands</param>
                ''' <param name="TagId"> ie Head</param>
                ''' <returns></returns>
                Public Shared Function RemoveTags(ByRef html As String, ByVal TagId As String) As String
                    ' Remove HTML tags.
                    html = Regex.Replace(html, "<" & TagId & ">", "")

                    Return html
                End Function

                ''' <summary>
                ''' Main Searcher
                ''' </summary>
                ''' <param name="Text">to be searched </param>
                ''' <param name="Pattern">RegEx Search String</param>
                ''' <returns></returns>
                Public Shared Function RegExSearch(ByRef Text As String, Pattern As String) As List(Of String)
                    Dim Searcher As New Regex(Pattern)
                    Dim iMatch As Match = Searcher.Match(Text)
                    Dim iMatches As New List(Of String)
                    Do While iMatch.Success
                        iMatches.Add(iMatch.Value)
                        iMatch = iMatch.NextMatch
                    Loop

                    Return iMatches
                End Function

                ''' <summary>
                ''' Find Between
                ''' Start ........
                ''' ....................
                ''' ....................Stop
                ''' </summary>
                ''' <param name="Text"></param>
                ''' <param name="StartStr"></param>
                ''' <param name="EndStr"></param>
                ''' <returns></returns>
                Public Shared Function RegExFindBetween(ByRef Text As String, ByRef StartStr As String, ByRef EndStr As String) As List(Of String)
                    Dim Tag = StartStr & "+(.|\n)*" & EndStr
                    Dim Searcher As New Regex(Tag)
                    Dim iMatch As Match = Searcher.Match(Text)
                    Dim iMatches As New List(Of String)
                    Do While iMatch.Success
                        iMatches.Add(iMatch.Value)
                        iMatch = iMatch.NextMatch
                    Loop
                    Return iMatches
                End Function

                ''' <summary>
                ''' Using RegEx for Searching Searches for (tag)
                ''' </summary>
                ''' <param name="Text">TobeSearched</param>
                ''' <returns></returns>
                Public Shared Function RegExFindTagLeft(ByRef Text As String) As List(Of String)
                    Dim Tag As String = "<\w+>"
                    Return RegExSearch(Text, Tag)
                End Function

                ''' <summary>
                ''' Using RegEx for Searching Searches for (/tag)
                ''' </summary>
                ''' <param name="Text">TobeSearched</param>
                ''' <returns></returns>
                Public Shared Function RegExFindTagsRight(ByRef Text As String) As List(Of String)
                    Dim Tag As String = "<\/\w+>"
                    Return RegExSearch(Text, Tag)
                End Function

                ''' <summary>
                ''' Searches for (......Tag)
                ''' </summary>
                ''' <param name="Text"></param>
                ''' <param name="iTag"></param>
                ''' <returns></returns>
                Public Shared Function RegExFindTagsBetweenRight(ByRef Text As String, ByRef iTag As String) As List(Of String)
                    Dim Tag As String = "<.*?" & iTag & ">"
                    Return RegExSearch(Text, Tag)
                End Function

                ''' <summary>
                ''' Searches for (tag......)
                ''' </summary>
                ''' <param name="Text"></param>
                ''' <param name="iTag"></param>
                ''' <returns></returns>
                Public Shared Function RegExFindTagsBetweenLeft(ByRef Text As String, ByRef iTag As String) As List(Of String)
                    Dim Tag As String = "<" & iTag & ".*?>"
                    Return RegExSearch(Text, Tag)
                End Function

                ''' <summary>
                ''' Searches for (tag......
                '''               ........
                '''               ........tag)
                '''
                ''' </summary>
                ''' <param name="Text"></param>
                ''' <param name="iTag"></param>
                ''' <returns></returns>
                Public Shared Function RegExFindTagsBetweenBoth(ByRef Text As String, ByRef iTag As String) As List(Of String)
                    '  Dim Tag As String = "<" & iTag & ".*?/" & iTag & ">"
                    Dim Tag As String = "<" & iTag & "+(.|\n)*" & iTag & ">"
                    Return RegExSearch(Text, Tag)
                End Function

                ''' <summary>
                ''' Searches for (....Tag....)
                ''' </summary>
                ''' <param name="Text"></param>
                ''' <param name="iTag"></param>
                ''' <returns></returns>
                Public Shared Function RegExFindTag(ByRef Text As String, ByRef iTag As String) As List(Of String)
                    Dim Tag As String = "<.*" & iTag & ".*>"
                    Return RegExSearch(Text, Tag)
                End Function

                ''' <summary>
                ''' Searches for
                ''' Searches for (tag...........)
                ''' Searches for (...........Tag)
                ''' Searches for (.....Tag... ..)
                ''' Searches for (tag.........tag)
                ''' </summary>
                ''' <param name="Text"></param>
                ''' <param name="Tag"></param>
                ''' <returns></returns>
                Public Shared Function FindByTag(ByRef Text As String, ByRef Tag As String) As List(Of String)
                    Dim result2 = New List(Of String)

                    result2.AddRange(RegExFindTag(Text, Tag))
                    result2.AddRange(RegExFindTagsBetweenBoth(Text, Tag))
                    result2.AddRange(RegExFindTagsBetweenLeft(Text, Tag))
                    result2.AddRange(RegExFindTagsBetweenRight(Text, Tag))

                    Return result2
                End Function

            End Class

            ''' <summary>
            ''' Colors Words in RichText Box
            ''' </summary>
            Public Class SyntaxHighlighter

                ''' <summary>
                ''' Returns Propercase Sentence
                ''' </summary>
                ''' <param name="TheString">String to be formatted</param>
                ''' <returns></returns>
                Private Shared Function ProperCase(ByRef TheString As String) As String
                    ProperCase = UCase(Left(TheString, 1))

                    For i = 2 To Len(TheString)

                        ProperCase = If(Mid(TheString, i - 1, 1) = " ", ProperCase & UCase(Mid(TheString, i, 1)), ProperCase & LCase(Mid(TheString, i, 1)))
                    Next i
                End Function

                Public Shared SyntaxTerms() As String = ({"SPYDAZ", "ABS", "ACCESS", "ADDITEM", "ADDNEW", "ALIAS", "AND", "ANY", "APP", "APPACTIVATE", "APPEND", "APPENDCHUNK", "ARRANGE", "AS", "ASC", "ATN", "BASE", "BEEP", "BEGINTRANS", "BINARY", "BYVAL", "CALL", "CASE", "CCUR", "CDBL", "CHDIR", "CHDRIVE", "CHR", "CHR$", "CINT", "CIRCLE", "CLEAR", "CLIPBOARD", "CLNG", "CLOSE", "CLS", "COMMAND", "
COMMAND$", "COMMITTRANS", "COMPARE", "CONST", "CONTROL", "CONTROLS", "COS", "CREATEDYNASET", "CSNG", "CSTR", "CURDIR$", "CURRENCY", "CVAR", "CVDATE", "DATA", "DATE", "DATE$", "DATESERIAL", "DATEVALUE", "DAY", "
DEBUG", "DECLARE", "DEFCUR", "CEFDBL", "DEFINT", "DEFLNG", "DEFSNG", "DEFSTR", "DEFVAR", "DELETE", "DIM", "DIR", "DIR$", "DO", "DOEVENTS", "DOUBLE", "DRAG", "DYNASET", "EDIT", "ELSE", "ELSEIF", "END", "ENDDOC", "ENDIF", "
ENVIRON$", "EOF", "EQV", "ERASE", "ERL", "ERR", "ERROR", "ERROR$", "EXECUTESQL", "EXIT", "EXP", "EXPLICIT", "FALSE", "FIELDSIZE", "FILEATTR", "FILECOPY", "FILEDATETIME", "FILELEN", "FIX", "FOR", "FORM", "FORMAT", "
FORMAT$", "FORMS", "FREEFILE", "FUNCTION", "GET", "GETATTR", "GETCHUNK", "GETDATA", "DETFORMAT", "GETTEXT", "GLOBAL", "GOSUB", "GOTO", "HEX", "HEX$", "HIDE", "HOUR", "IF", "IMP", "INPUT", "INPUT$", "INPUTBOX", "INPUTBOX$", "
INSTR", "INT", "INTEGER", "IS", "ISDATE", "ISEMPTY", "ISNULL", "ISNUMERIC", "KILL", "LBOUND", "LCASE", "LCASE$", "LEFT", "LEFT$", "LEN", "LET", "LIB", "LIKE", "LINE", "LINKEXECUTE", "LINKPOKE", "LINKREQUEST", "
LINKSEND", "LOAD", "LOADPICTURE", "LOC", "LOCAL", "LOCK", "LOF", "LOG", "LONG", "LOOP", "LSET", "LTRIM",
            "LTRIM$", "ME", "MID", "MID$", "MINUTE", "MKDIR", "MOD", "MONTH", "MOVE", "MOVEFIRST", "MOVELAST", "MOVENEXT", "MOVEPREVIOUS",
            "MOVERELATIVE", "MSGBOX", "NAME", "NEW", "NEWPAGE", "NEXT", "NEXTBLOCK", "NOT", "NOTHING",
            "NOW", "NULL", "OCT", "OCT$", "ON", "OPEN", "OPENDATABASE", "OPTION", "OR", "OUTPUT", "POINT", "PRESERVE", "PRINT",
            "PRINTER", "PRINTFORM", "PRIVATE", "PSET", "PUT", "PUBLIC", "QBCOLOR", "RANDOM", "RANDOMIZE", "READ", "REDIM", "REFRESH",
            "REGISTERDATABASE", "REM", "REMOVEITEM", "RESET", "RESTORE", "RESUME", "RETURN", "RGB", "RIGHT", "RIGHT$", "RMDIR", "RND",
            "ROLLBACK", "RSET", "RTRIM", "RTRIM$", "SAVEPICTURE", "SCALE", "SECOND", "SEEK", "SELECT", "SENDKEYS", "SET", "SETATTR",
            "SETDATA", "SETFOCUS", "SETTEXT", "SGN", "SHARED",
            "SHELL", "SHOW", "SIN", "SINGLE", "SPACE", "SPACE$", "SPC", "SQR",
            "STATIC", "STEP", "STOP", "STR", "STR$", "STRCOMP", "STRING", "STRING$", "SUB",
            "SYSTEM", "TAB", "TAN", "TEXT", "TEXTHEIGHT", "TEXTWIDTH", "THEN", "TIME", "TIME$",
            "TIMER", "TIMESERIAL", "TIMEVALUE", "TO", "TRIM",
            "TRIM$", "TRUE", "TYPE", "TYPEOF", "UBOUND", "UCASE", "UCASE$", "UNLOAD",
            "UNLOCK", "UNTIL", "UPDATE", "USING", "VAL", "VARIANT", "VARTYPE", "WEEKDAY", "WEND", "WHILE", "WIDTH",
            "WRITE", "XOR", "YEAR", "ZORDER", "ADDHANDLER", "ADDRESSOF", "ALIAS", "AND", "ANDALSO", "AS", "BYREF",
            "BOOLEAN", "BYTE", "BYVAL", "CALL", "CASE", "CATCH", "CBOOL", "CBYTE", "CCHAR", "CDATE",
            "CDEC", "CDBL", "CHAR", "CINT", "CLASS", "CLNG", "COBJ", "CONST", "CONTINUE", "CSBYTE",
            "CSHORT", "CSNG", "CSTR", "CTYPE", "CUINT", "CULNG", "CUSHORT", "DATE", "DECIMAL", "DECLARE",
            "DEFAULT", "DELEGATE", "DIM", "DIRECTCAST", "DOUBLE", "DO", "EACH", "ELSE", "ELSEIF", "END",
            "ENDIF", "ENUM", "ERASE", "ERROR", "EVENT", "EXIT", "FALSE", "FINALLY", "FOR", "FRIEND",
            "FUNCTION", "GET", "GETTYPE", "GLOBAL", "GOSUB", "GOTO", "HANDLES", "IF", "IMPLEMENTS",
            "IMPORTS", "IN", "INHERITS", "INTEGER", "INTERFACE", "IS", "ISNOT", "LET", "LIB", "LIKE",
            "LONG", "LOOP", "ME", "MOD", "MODULE", "MUSTINHERIT", "MUSTOVERRIDE", "MYBASE", "MYCLASS",
            "NAMESPACE", "NARROWING", "NEW", "NEXT", "NOT", "NOTHING", "NOTINHERITABLE", "NOTOVERRIDABLE",
            "OBJECT", "ON", "OF", "OPERATOR", "OPTION", "OPTIONAL", "OR", "ORELSE", "OVERLOADS",
            "OVERRIDABLE", "OVERRIDES", "PARAMARRAY", "PARTIAL", "PRIVATE", "PROPERTY", "PROTECTED",
            "PUBLIC", "RAISEEVENT", "READONLY", "REDIM", "REM", "REMOVEHANDLER", "RESUME", "RETURN",
            "SBYTE", "SELECT", "SET", "SHADOWS", "SHARED", "SHORT", "SINGLE", "STATIC", "STEP", "STOP",
            "STRING", "STRUCTURE", "SUB", "SYNCLOCK", "THEN", "THROW", "TO", "TRUE", "TRY", "TRYCAST",
            "TYPEOF", "WEND", "VARIANT", "UINTEGER", "ULONG", "USHORT", "USING", "WHEN", "WHILE", "WIDENING",
            "WITH", "WITHEVENTS", "WRITEONLY",
            "XOR", "#CONST", "#ELSE", "#ELSEIF", "#END", "#IF"})

                Private Shared indexOfSearchText As Integer = 0

                Private Shared start As Integer = 0

                Private mGrammar As New List(Of String)

                Public Shared Sub ColorSearchTerm(ByRef SearchStr As String, Rtb As RichTextBox)
                    ColorSearchTerm(SearchStr, Rtb, Color.CadetBlue)
                End Sub

                Public Shared Sub ColorSearchTerm(ByRef SearchStr As String, Rtb As RichTextBox, ByRef MyColor As Color)
                    Dim startindex As Integer = 0
                    start = 0
                    indexOfSearchText = 0

                    If SearchStr <> "" Then

                        SearchStr = SearchStr & " "
                        If SearchStr.Length > 0 Then
                            startindex = FindText(Rtb, ProperCase(SearchStr), start, Rtb.Text.Length)
                        End If
                        If SearchStr.Length > 0 And startindex = 0 Then
                            startindex = FindText(Rtb, LCase(SearchStr), start, Rtb.Text.Length)
                        End If
                        If SearchStr.Length > 0 And startindex = 0 Then
                            startindex = FindText(Rtb, UCase(SearchStr), start, Rtb.Text.Length)
                        End If
                        If SearchStr.Length > 0 And startindex = 0 Then
                            startindex = FindText(Rtb, SearchStr, start, Rtb.Text.Length)
                        End If
                        ' If string was found in the RichTextBox, highlight it
                        If startindex >= 0 Then
                            ' Set the highlight color as red
                            Rtb.SelectionColor = MyColor

                            ' Find the end index. End Index = number of characters in textbox
                            Dim endindex As Integer = SearchStr.Length
                            ' Highlight the search string

                            Rtb.Select(startindex, endindex)
                            Rtb.SelectedText.ToUpper()
                            ' mark the start position after the position of last search string
                            start = startindex + endindex

                        End If
                    Else
                    End If
                    Rtb.Select(Rtb.TextLength, Rtb.TextLength)
                End Sub

                ''' <summary>
                ''' Searches For Internal Syntax
                ''' </summary>
                ''' <param name="Rtb"></param>
                ''' <remarks></remarks>
                Public Shared Sub SearchSyntax(ByRef Rtb As RichTextBox)
                    'Searches Basic Syntax
                    For Each wrd As String In SyntaxTerms
                        start = 0
                        indexOfSearchText = 0
                        ColorSearchTerm(wrd, Rtb)
                    Next
                    For Each wrd As String In SyntaxTerms
                        start = 0
                        indexOfSearchText = 0
                        ColorSearchTerm(wrd, Rtb)
                    Next
                End Sub

                ''' <summary>
                ''' Searches For Internal Syntax
                ''' </summary>
                ''' <param name="Rtb"></param>
                ''' <remarks></remarks>
                Public Shared Sub SearchSyntax(ByRef Rtb As RichTextBox, ByRef Terms As List(Of String))
                    'Searches Basic Syntax
                    For Each wrd As String In SyntaxTerms
                        start = 0
                        indexOfSearchText = 0
                        ColorSearchTerm(wrd, Rtb)
                    Next
                    For Each wrd As String In Terms
                        start = 0
                        indexOfSearchText = 0
                        ColorSearchTerm(UCase(wrd), Rtb)
                    Next
                End Sub

                ''' <summary>
                ''' Searches for Specific Word to colorize (Blue)
                ''' </summary>
                ''' <param name="Rtb"></param>
                ''' <param name="SearchStr"></param>
                ''' <remarks></remarks>
                Public Shared Sub SearchSyntax(ByRef Rtb As RichTextBox, ByRef SearchStr As String)
                    start = 0
                    indexOfSearchText = 0
                    ColorSearchTerm(SearchStr, Rtb)
                End Sub

                ''' <summary>
                ''' Searches for Specfic word to colorize specified color
                ''' </summary>
                ''' <param name="Rtb"></param>
                ''' <param name="SearchStr"></param>
                ''' <param name="MyColor"></param>
                ''' <remarks></remarks>
                Public Shared Sub SearchSyntax(ByRef Rtb As RichTextBox, ByRef SearchStr As String, MyColor As Color)

                    start = 0
                    indexOfSearchText = 0
                    ColorSearchTerm(SearchStr, Rtb, MyColor)

                End Sub

                Private Shared Function FindText(ByRef Rtb As RichTextBox, SearchStr As String,
                                                                ByVal searchStart As Integer, ByVal searchEnd As Integer) As Integer

                    ' Unselect the previously searched string
                    If searchStart > 0 AndAlso searchEnd > 0 AndAlso indexOfSearchText >= 0 Then
                        Rtb.Undo()
                    End If

                    ' Set the return value to -1 by default.
                    Dim retVal As Integer = -1

                    ' A valid starting index should be specified. if indexOfSearchText = -1, the end of search
                    If searchStart >= 0 AndAlso indexOfSearchText >= 0 Then

                        ' A valid ending index
                        If searchEnd > searchStart OrElse searchEnd = -1 Then

                            ' Find the position of search string in RichTextBox
                            indexOfSearchText = Rtb.Find(SearchStr, searchStart, searchEnd, RichTextBoxFinds.WholeWord)

                            ' Determine whether the text was found in richTextBox1.
                            If indexOfSearchText <> -1 Then
                                ' Return the index to the specified search text.
                                retVal = indexOfSearchText
                            End If

                        End If
                    End If
                    Return retVal

                End Function

                Public Shared Sub HighlightInternalSyntax(ByRef sender As RichTextBox)

                    For Each Word As String In SyntaxTerms
                        HighlightWord(sender, Word)
                        HighlightWord(sender, LCase(Word))
                        HighlightWord(sender, ProperCase(Word))
                    Next

                End Sub

                Public Shared Sub HighlightWord(ByRef sender As RichTextBox, ByRef word As String)

                    Dim index As Integer = 0
                    While index < sender.Text.LastIndexOf(word)
                        'find
                        sender.Find(word, index, sender.TextLength, RichTextBoxFinds.WholeWord)
                        'select and color
                        sender.SelectionColor = Color.OrangeRed
                        index = sender.Text.IndexOf(word, index) + 1
                    End While
                End Sub

            End Class

        End Namespace

        Public Module temp

            ''' <summary>
            '''     A string extension method that query if '@this' satisfy the specified pattern.
            ''' </summary>
            ''' <param name="this">The @this to act on.</param>
            ''' <param name="pattern">The pattern to use. Use '*' as wildcard string.</param>
            ''' <returns>true if '@this' satisfy the specified pattern, false if not.</returns>
            <System.Runtime.CompilerServices.Extension>
            Public Function IsLike(this As String, pattern As String) As Boolean
                ' Turn the pattern into regex pattern, and match the whole string with ^$
                Dim regexPattern As String = "^" + Regex.Escape(pattern) + "$"

                ' Escape special character ?, #, *, [], and [!]
                regexPattern = regexPattern.Replace("\[!", "[^").Replace("\[", "[").Replace("\]", "]").Replace("\?", ".").Replace("\*", ".*").Replace("\#", "\d")

                Return Regex.IsMatch(this, regexPattern)
            End Function

        End Module
        ''' <summary>
        ''' Basic String Extensions
        ''' </summary>
        Public Module Text_Extensions

            ''' <summary>
            ''' Returns from index to end of file
            ''' </summary>
            ''' <param name="Str">String</param>
            ''' <param name="indx">Index</param>
            ''' <returns></returns>
            <Runtime.CompilerServices.Extension()>
            Public Function GetSlice(ByRef Str As String, ByRef indx As Integer) As String
                If indx <= Str.Length Then
                    Str.Substring(indx, Str.Length)
                    Return Str(indx)
                Else
                End If
                Return Nothing
            End Function

            <Runtime.CompilerServices.Extension()>
            Public Function FormatJsonOutput(ByVal jsonString As String) As String
                Dim stringBuilder = New StringBuilder()
                Dim escaping As Boolean = False
                Dim inQuotes As Boolean = False
                Dim indentation As Integer = 0

                For Each character As Char In jsonString

                    If escaping Then
                        escaping = False
                        stringBuilder.Append(character)
                    Else

                        If character = "\"c Then
                            escaping = True
                            stringBuilder.Append(character)
                        ElseIf character = """"c Then
                            inQuotes = Not inQuotes
                            stringBuilder.Append(character)
                        ElseIf Not inQuotes Then

                            If character = ","c Then
                                stringBuilder.Append(character)
                                stringBuilder.Append(vbCrLf)
                                stringBuilder.Append(vbTab, indentation)
                            ElseIf character = "["c OrElse character = "{"c Then
                                stringBuilder.Append(character)
                                stringBuilder.Append(vbCrLf)
                                stringBuilder.Append(vbTab, System.Threading.Interlocked.Increment(indentation))
                            ElseIf character = "]"c OrElse character = "}"c Then
                                stringBuilder.Append(vbCrLf)
                                stringBuilder.Append(vbTab, System.Threading.Interlocked.Decrement(indentation))
                                stringBuilder.Append(character)
                            ElseIf character = ":"c Then
                                stringBuilder.Append(character)
                                stringBuilder.Append(vbTab)
                            ElseIf Not Char.IsWhiteSpace(character) Then
                                stringBuilder.Append(character)
                            End If
                        Else
                            stringBuilder.Append(character)
                        End If
                    End If
                Next

                Return stringBuilder.ToString()
            End Function

            ''' <summary>
            ''' Checks if string is a reserved VBscipt Keyword
            ''' </summary>
            ''' <param name="keyword"></param>
            ''' <returns></returns>
            <Runtime.CompilerServices.Extension()>
            Function IsReservedWord(ByVal keyword As String) As Boolean
                Dim IsReserved = False
                Select Case LCase(keyword)
                    Case "and" : IsReserved = True
                    Case "as" : IsReserved = True
                    Case "boolean" : IsReserved = True
                    Case "byref" : IsReserved = True
                    Case "byte" : IsReserved = True
                    Case "byval" : IsReserved = True
                    Case "call" : IsReserved = True
                    Case "case" : IsReserved = True
                    Case "class" : IsReserved = True
                    Case "const" : IsReserved = True
                    Case "currency" : IsReserved = True
                    Case "debug" : IsReserved = True
                    Case "dim" : IsReserved = True
                    Case "do" : IsReserved = True
                    Case "double" : IsReserved = True
                    Case "each" : IsReserved = True
                    Case "else" : IsReserved = True
                    Case "elseif" : IsReserved = True
                    Case "empty" : IsReserved = True
                    Case "end" : IsReserved = True
                    Case "endif" : IsReserved = True
                    Case "enum" : IsReserved = True
                    Case "eqv" : IsReserved = True
                    Case "event" : IsReserved = True
                    Case "exit" : IsReserved = True
                    Case "false" : IsReserved = True
                    Case "for" : IsReserved = True
                    Case "function" : IsReserved = True
                    Case "get" : IsReserved = True
                    Case "goto" : IsReserved = True
                    Case "if" : IsReserved = True
                    Case "imp" : IsReserved = True
                    Case "implements" : IsReserved = True
                    Case "in" : IsReserved = True
                    Case "integer" : IsReserved = True
                    Case "is" : IsReserved = True
                    Case "let" : IsReserved = True
                    Case "like" : IsReserved = True
                    Case "long" : IsReserved = True
                    Case "loop" : IsReserved = True
                    Case "lset" : IsReserved = True
                    Case "me" : IsReserved = True
                    Case "mod" : IsReserved = True
                    Case "new" : IsReserved = True
                    Case "next" : IsReserved = True
                    Case "not" : IsReserved = True
                    Case "nothing" : IsReserved = True
                    Case "null" : IsReserved = True
                    Case "on" : IsReserved = True
                    Case "option" : IsReserved = True
                    Case "optional" : IsReserved = True
                    Case "or" : IsReserved = True
                    Case "paramarray" : IsReserved = True
                    Case "preserve" : IsReserved = True
                    Case "private" : IsReserved = True
                    Case "public" : IsReserved = True
                    Case "raiseevent" : IsReserved = True
                    Case "redim" : IsReserved = True
                    Case "rem" : IsReserved = True
                    Case "resume" : IsReserved = True
                    Case "rset" : IsReserved = True
                    Case "select" : IsReserved = True
                    Case "set" : IsReserved = True
                    Case "shared" : IsReserved = True
                    Case "single" : IsReserved = True
                    Case "static" : IsReserved = True
                    Case "stop" : IsReserved = True
                    Case "sub" : IsReserved = True
                    Case "then" : IsReserved = True
                    Case "to" : IsReserved = True
                    Case "true" : IsReserved = True
                    Case "type" : IsReserved = True
                    Case "typeof" : IsReserved = True
                    Case "until" : IsReserved = True
                    Case "variant" : IsReserved = True
                    Case "wend" : IsReserved = True
                    Case "while" : IsReserved = True
                    Case "with" : IsReserved = True
                    Case "xor" : IsReserved = True
                End Select
                Return IsReserved
            End Function

            ''' <summary>
            ''' Convert object to Json String
            ''' </summary>
            ''' <param name="Item"></param>
            ''' <returns></returns>
            <Runtime.CompilerServices.Extension()>
            Public Function ToJson(ByRef Item As Object) As String
                Dim Converter As New JavaScriptSerializer
                Return Converter.Serialize(Item)

            End Function

            ''' <summary>
            ''' DETECT IF STATMENT IS AN IF/THEN DETECT IF STATMENT IS AN IF/THEN -- -RETURNS PARTS DETIFTHEN
            ''' = DETECTLOGIC(USERINPUT, "IF", "THEN", IFPART, THENPART)
            ''' </summary>
            ''' <param name="userinput"></param>
            ''' <param name="LOGICA">"IF", can also be replace by "IT CAN BE SAID THAT</param>
            ''' <param name="LOGICB">"THEN" can also be replaced by "it must follow that"</param>
            ''' <param name="IFPART">supply empty string to be used to hold part</param>
            ''' <param name="THENPART">supply empty string to be used to hold part</param>
            ''' <returns>true/false</returns>
            ''' <remarks></remarks>
            <System.Runtime.CompilerServices.Extension()>
            Public Function DetectLOGIC(ByRef userinput As String, ByRef LOGICA As String, ByRef LOGICB As String, ByRef IFPART As String, ByRef THENPART As String) As Boolean
                If InStr(1, userinput, LOGICA, 1) > 0 And InStr(1, userinput, " " & LOGICB & " ", 1) > 0 Then
                    'SPLIT USER INPUT
                    Call SplitPhrase(userinput, " " & LOGICB & " ", IFPART, THENPART)

                    IFPART = Replace(IFPART, LOGICA, "", 1, -1, CompareMethod.Text)
                    THENPART = Replace(THENPART, " " & LOGICB & " ", "", 1, -1, CompareMethod.Text)
                    DetectLOGIC = True
                Else
                    DetectLOGIC = False
                End If
            End Function

            ''' <summary>
            ''' Returns a delimited string from the list.
            ''' </summary>
            ''' <param name="ls"></param>
            ''' <param name="delimiter"></param>
            ''' <returns></returns>
            <System.Runtime.CompilerServices.Extension>
            Public Function ToDelimitedString(ls As List(Of String), delimiter As String) As String
                Dim sb As New StringBuilder
                For Each buf As String In ls
                    sb.Append(buf)
                    sb.Append(delimiter)
                Next
                Return sb.ToString.Trim(CChar(delimiter))
            End Function

            ''' <summary>
            ''' Gets the string before the given string parameter.
            ''' </summary>
            ''' <param name="value">The default value.</param>
            ''' <param name="x">The given string parameter.</param>
            ''' <returns></returns>
            ''' <remarks>Unlike GetBetween and GetAfter, this does not Trim the result.</remarks>
            <System.Runtime.CompilerServices.Extension>
            Public Function GetBefore(value As String, x As String) As String
                Dim xPos = value.IndexOf(x, StringComparison.Ordinal)
                Return If(xPos = -1, [String].Empty, value.Substring(0, xPos))
            End Function

            ''' <summary>
            ''' Gets the string between the given string parameters.
            ''' </summary>
            ''' <param name="value">The source value.</param>
            ''' <param name="x">The left string sentinel.</param>
            ''' <param name="y">The right string sentinel</param>
            ''' <returns></returns>
            ''' <remarks>Unlike GetBefore, this method trims the result</remarks>
            <System.Runtime.CompilerServices.Extension>
            Public Function GetBetween(value As String, x As String, y As String) As String
                Dim xPos = value.IndexOf(x, StringComparison.Ordinal)
                Dim yPos = value.LastIndexOf(y, StringComparison.Ordinal)
                If xPos = -1 OrElse xPos = -1 Then
                    Return [String].Empty
                End If
                Dim startIndex = xPos + x.Length
                Return If(startIndex >= yPos, [String].Empty, value.Substring(startIndex, yPos - startIndex).Trim())
            End Function

            ''' <summary>
            ''' Gets the string after the given string parameter.
            ''' </summary>
            ''' <param name="value">The default value.</param>
            ''' <param name="x">The given string parameter.</param>
            ''' <returns></returns>
            ''' <remarks>Unlike GetBefore, this method trims the result</remarks>
            <System.Runtime.CompilerServices.Extension>
            Public Function GetAfter(value As String, x As String) As String
                Dim xPos = value.LastIndexOf(x, StringComparison.Ordinal)
                If xPos = -1 Then
                    Return [String].Empty
                End If
                Dim startIndex = xPos + x.Length
                Return If(startIndex >= value.Length, [String].Empty, value.Substring(startIndex).Trim())
            End Function

            <System.Runtime.CompilerServices.Extension>
            Public Function GetWordsBetween(ByRef InputStr As String, ByRef StartStr As String, ByRef StopStr As String)
                Return InputStr.ExtractStringBetween(StartStr, StopStr)
            End Function

            ''' <summary>
            ''' extracts string between defined strings
            ''' </summary>
            ''' <param name="value">base sgtring</param>
            ''' <param name="strStart">Start string</param>
            ''' <param name="strEnd">End string</param>
            ''' <returns></returns>
            <System.Runtime.CompilerServices.Extension()>
            Public Function ExtractStringBetween(ByVal value As String, ByVal strStart As String, ByVal strEnd As String) As String
                If Not String.IsNullOrEmpty(value) Then
                    Dim i As Integer = value.IndexOf(strStart)
                    Dim j As Integer = value.IndexOf(strEnd)
                    Return value.Substring(i, j - i)
                Else
                    Return value
                End If
            End Function

            ''' <summary>
            ''' gets the last word
            ''' </summary>
            ''' <param name="InputStr"></param>
            ''' <returns></returns>
            <System.Runtime.CompilerServices.Extension()>
            Public Function GetSuffix(ByRef InputStr As String) As String
                Dim TempArr() As String = Split(InputStr, " ")
                Dim Count As Integer = TempArr.Count - 1
                Return TempArr(Count)
            End Function

            ''' <summary>
            ''' Returns the first Word
            ''' </summary>
            ''' <param name="Statement"></param>
            ''' <returns></returns>
            <System.Runtime.CompilerServices.Extension()>
            Public Function GetPrefix(ByRef Statement As String) As String
                Dim StrArr() As String = Split(Statement, " ")
                Return StrArr(0)
            End Function

            ''' <summary>
            ''' Advanced search String pattern Wildcard denotes which position 1st =1 or 2nd =2 Send
            ''' Original input &gt; Search pattern to be used &gt; Wildcard requred SPattern = "WHAT
            ''' COLOUR DO YOU LIKE * OR *" Textstr = "WHAT COLOUR DO YOU LIKE red OR black" ITEM_FOUND =
            ''' = SearchPattern(USERINPUT, SPattern, 1) ---- RETURNS RED ITEM_FOUND = =
            ''' SearchPattern(USERINPUT, SPattern, 1) ---- RETURNS black
            ''' </summary>
            ''' <param name="TextSTR">
            ''' TextStr Required. String.EG: "WHAT COLOUR DO YOU LIKE red OR black"
            ''' </param>
            ''' <param name="SPattern">
            ''' SPattern Required. String.EG: "WHAT COLOUR DO YOU LIKE * OR *"
            ''' </param>
            ''' <param name="Wildcard">Wildcard Required. Integer.EG: 1st =1 or 2nd =2</param>
            ''' <returns></returns>
            ''' <remarks>* in search pattern</remarks>
            <Runtime.CompilerServices.Extension()>
            Public Function SearchPattern(ByRef TextSTR As String, ByRef SPattern As String, ByRef Wildcard As Short) As String
                Dim SearchP2 As String
                Dim SearchP1 As String
                Dim TextStrp3 As String
                Dim TextStrp4 As String
                SearchPattern = ""
                SearchP2 = ""
                SearchP1 = ""
                TextStrp3 = ""
                TextStrp4 = ""
                If TextSTR Like SPattern = True Then
                    Select Case Wildcard
                        Case 1
                            Call SplitPhrase(SPattern, "*", SearchP1, SearchP2)
                            TextSTR = Replace(TextSTR, SearchP1, "", 1, -1, CompareMethod.Text)

                            SearchP2 = Replace(SearchP2, "*", "", 1, -1, CompareMethod.Text)
                            Call SplitPhrase(TextSTR, SearchP2, TextStrp3, TextStrp4)

                            TextSTR = TextStrp3

                        Case 2
                            Call SplitPhrase(SPattern, "*", SearchP1, SearchP2)
                            SPattern = Replace(SPattern, SearchP1, " ", 1, -1, CompareMethod.Text)
                            TextSTR = Replace(TextSTR, SearchP1, " ", 1, -1, CompareMethod.Text)

                            Call SplitPhrase(SearchP2, "*", SearchP1, SearchP2)
                            Call SplitPhrase(TextSTR, SearchP1, TextStrp3, TextStrp4)

                            TextSTR = TextStrp4

                    End Select

                    SearchPattern = TextSTR
                    LTrim(SearchPattern)
                    RTrim(SearchPattern)
                Else
                End If

            End Function

            ''' <summary>
            ''' SPLITS THE GIVEN PHRASE UP INTO TWO PARTS by dividing word SplitPhrase(Userinput, "and",
            ''' Firstp, SecondP)
            ''' </summary>
            ''' <param name="PHRASE">Sentence to be divided</param>
            ''' <param name="DIVIDINGWORD">String: Word to divide sentence by</param>
            ''' <param name="FIRSTPART">String: firstpart of sentence to be populated</param>
            ''' <param name="SECONDPART">String: Secondpart of sentence to be populated</param>
            ''' <remarks></remarks>
            <Runtime.CompilerServices.Extension()>
            Public Sub SplitPhrase(ByVal PHRASE As String, ByRef DIVIDINGWORD As String, ByRef FIRSTPART As String, ByRef SECONDPART As String)
                Dim POS As Short
                POS = InStr(PHRASE, DIVIDINGWORD)
                If (POS > 0) Then
                    FIRSTPART = Trim(Left(PHRASE, POS - 1))
                    SECONDPART = Trim(Right(PHRASE, Len(PHRASE) - POS - Len(DIVIDINGWORD) + 1))
                Else
                    FIRSTPART = ""
                    SECONDPART = PHRASE
                End If
            End Sub

            <Runtime.CompilerServices.Extension()>
            Public Function RemoveFullStop(ByRef MESSAGE As String) As String
Loop1:
                If Right(MESSAGE, 1) = "." Then MESSAGE = Left(MESSAGE, Len(MESSAGE) - 1) : GoTo Loop1
                Return MESSAGE
            End Function

            ''' <summary>
            ''' Split string to List of strings
            ''' </summary>
            ''' <param name="Str">base string</param>
            ''' <param name="Seperator">to be seperated by</param>
            ''' <returns></returns>
            <System.Runtime.CompilerServices.Extension()>
            Public Function SplitToList(ByRef Str As String, ByVal Seperator As String) As List(Of String)
                Dim lst As New List(Of String)
                If Str <> "" = True And Seperator <> "" Then

                    Dim Found() As String = Str.Split(Seperator)
                    For Each item In Found
                        lst.Add(item)
                    Next
                Else

                End If
                Return lst
            End Function

            ''' <summary>
            ''' Counts the vowels used (AEIOU)
            ''' </summary>
            ''' <param name="InputString"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            <Runtime.CompilerServices.Extension()>
            Public Function VowelCount(ByVal InputString As String) As Integer
                Dim v(9) As String 'Declare an array  of 10 elements 0 to 9
                Dim vcount As Short 'This variable will contain number of vowels
                Dim flag As Integer
                Dim strLen As Integer
                Dim i As Integer
                v(0) = "a" 'First element of array is assigned small a
                v(1) = "i"
                v(2) = "o"
                v(3) = "u"
                v(4) = "e"
                v(5) = "A" 'Sixth element is assigned Capital A
                v(6) = "I"
                v(7) = "O"
                v(8) = "U"
                v(9) = "E"
                strLen = Len(InputString)

                For flag = 1 To strLen 'It will get every letter of entered string and loop
                    'will terminate when all letters have been examined

                    For i = 0 To 9 'Takes every element of v(9) one by one
                        'Check if current letter is a vowel
                        If Mid(InputString, flag, 1) = v(i) Then
                            vcount = vcount + 1 ' If letter is equal to vowel
                            'then increment vcount by 1
                        End If
                    Next i 'Consider next value of v(i)
                Next flag 'Consider next letter of the enterd string

                VowelCount = vcount

            End Function

            <Runtime.CompilerServices.Extension()>
            Public Function CountVowels(ByVal InputString As String) As Integer
                Dim v(9) As String 'Declare an array  of 10 elements 0 to 9
                Dim vcount As Short 'This variable will contain number of vowels
                Dim flag As Integer
                Dim strLen As Integer
                Dim i As Integer
                v(0) = "a" 'First element of array is assigned small a
                v(1) = "i"
                v(2) = "o"
                v(3) = "u"
                v(4) = "e"
                v(5) = "A" 'Sixth element is assigned Capital A
                v(6) = "I"
                v(7) = "O"
                v(8) = "U"
                v(9) = "E"
                strLen = Len(InputString)

                For flag = 1 To strLen 'It will get every letter of entered string and loop
                    'will terminate when all letters have been examined

                    For i = 0 To 9 'Takes every element of v(9) one by one
                        'Check if current letter is a vowel
                        If Mid(InputString, flag, 1) = v(i) Then
                            vcount = vcount + 1 ' If letter is equal to vowel
                            'then increment vcount by 1
                        End If
                    Next i 'Consider next value of v(i)
                Next flag 'Consider next letter of the enterd string

                CountVowels = vcount

            End Function

            ''' <summary>
            ''' Counts tokens in string
            ''' </summary>
            ''' <param name="Str">string to be searched</param>
            ''' <param name="Delimiter">delimiter such as space comma etc</param>
            ''' <returns></returns>
            <System.Runtime.CompilerServices.Extension>
            Public Function CountTokensInString(ByRef Str As String, ByRef Delimiter As String) As Integer
                Dim Words() As String = Split(Str, Delimiter)
                Return Words.Count
            End Function

            ''' <summary>
            ''' Add full stop to end of String
            ''' </summary>
            ''' <param name="MESSAGE"></param>
            ''' <returns></returns>
            <System.Runtime.CompilerServices.Extension()>
            Public Function AddFullStop(ByRef MESSAGE As String) As String
                AddFullStop = MESSAGE
                If MESSAGE = "" Then Exit Function
                MESSAGE = Trim(MESSAGE)
                If MESSAGE Like "*." Then Exit Function
                AddFullStop = MESSAGE + "."
            End Function

            'Text
            <Runtime.CompilerServices.Extension()>
            Public Function Capitalise(ByRef MESSAGE As String) As String
                Dim FirstLetter As String
                Capitalise = ""
                If MESSAGE = "" Then Exit Function
                FirstLetter = Left(MESSAGE, 1)
                FirstLetter = UCase(FirstLetter)
                MESSAGE = Right(MESSAGE, Len(MESSAGE) - 1)
                Capitalise = (FirstLetter + MESSAGE)
            End Function

            ''' <summary>
            ''' Capitalise the first letter of each word / Tilte Case
            ''' </summary>
            ''' <param name="words">A string - paragraph or sentence</param>
            ''' <returns>String</returns>
            <Runtime.CompilerServices.Extension()>
            Public Function CapitalizeWords(ByVal words As String)
                Dim output As System.Text.StringBuilder = New System.Text.StringBuilder()
                Dim exploded = words.Split(" ")
                If (exploded IsNot Nothing) Then
                    For Each word As String In exploded
                        If word IsNot Nothing Then
                            output.Append(word.Substring(0, 1).ToUpper).Append(word.Substring(1, word.Length - 1)).Append(" ")
                        End If

                    Next
                End If

                Return output.ToString()

            End Function

            ''' <summary>
            ''' SPLITS THE GIVEN PHRASE UP INTO TWO PARTS by dividing word SplitPhrase(Userinput, "and",
            ''' Firstp, SecondP)
            ''' </summary>
            ''' <param name="PHRASE">String: Sentence to be divided</param>
            ''' <param name="DIVIDINGWORD">String: Word to divide sentence by</param>
            ''' <param name="FIRSTPART">String-Returned : firstpart of sentence to be populated</param>
            ''' <param name="SECONDPART">String-Returned : Secondpart of sentence to be populated</param>
            ''' <remarks></remarks>
            <System.Runtime.CompilerServices.Extension()>
            Public Sub SplitString(ByVal PHRASE As String, ByRef DIVIDINGWORD As String, ByRef FIRSTPART As String, ByRef SECONDPART As String)
                Dim POS As Short
                'Check Error
                If DIVIDINGWORD IsNot Nothing And PHRASE IsNot Nothing Then

                    POS = InStr(PHRASE, DIVIDINGWORD)
                    If (POS > 0) Then
                        FIRSTPART = Trim(Left(PHRASE, POS - 1))
                        SECONDPART = Trim(Right(PHRASE, Len(PHRASE) - POS - Len(DIVIDINGWORD) + 1))
                    Else
                        FIRSTPART = ""
                        SECONDPART = PHRASE
                    End If
                Else

                End If
            End Sub

            ''' <summary>
            ''' Advanced search String pattern Wildcard denotes which position 1st =1 or 2nd =2 Send
            ''' Original input &gt; Search pattern to be used &gt; Wildcard requred SPattern = "WHAT
            ''' COLOUR DO YOU LIKE * OR *" Textstr = "WHAT COLOUR DO YOU LIKE red OR black" ITEM_FOUND =
            ''' = SearchPattern(USERINPUT, SPattern, 1) ---- RETURNS RED ITEM_FOUND = =
            ''' SearchPattern(USERINPUT, SPattern, 2) ---- RETURNS black
            ''' </summary>
            ''' <param name="TextSTR">TextStr = "Pick Red OR Blue" . String.</param>
            ''' <param name="SPattern">Search String = ("Pick * OR *") String.</param>
            ''' <param name="Wildcard">Wildcard Required. Integer. = 1= Red / 2= Blue</param>
            ''' <returns></returns>
            ''' <remarks>finds the * in search pattern</remarks>
            <System.Runtime.CompilerServices.Extension()>
            Public Function SearchStringbyPattern(ByRef TextSTR As String, ByRef SPattern As String, ByRef Wildcard As Short) As String
                Dim SearchP2 As String
                Dim SearchP1 As String
                Dim TextStrp3 As String
                Dim TextStrp4 As String
                SearchStringbyPattern = ""
                SearchP2 = ""
                SearchP1 = ""
                TextStrp3 = ""
                TextStrp4 = ""
                If TextSTR Like SPattern = True Then
                    Select Case Wildcard
                        Case 1
                            Call SplitString(SPattern, "*", SearchP1, SearchP2)
                            TextSTR = Replace(TextSTR, SearchP1, "", 1, -1, CompareMethod.Text)

                            SearchP2 = Replace(SearchP2, "*", "", 1, -1, CompareMethod.Text)
                            Call SplitString(TextSTR, SearchP2, TextStrp3, TextStrp4)

                            TextSTR = TextStrp3

                        Case 2
                            Call SplitString(SPattern, "*", SearchP1, SearchP2)
                            SPattern = Replace(SPattern, SearchP1, " ", 1, -1, CompareMethod.Text)
                            TextSTR = Replace(TextSTR, SearchP1, " ", 1, -1, CompareMethod.Text)

                            Call SplitString(SearchP2, "*", SearchP1, SearchP2)
                            Call SplitString(TextSTR, SearchP1, TextStrp3, TextStrp4)

                            TextSTR = TextStrp4

                    End Select

                    SearchStringbyPattern = TextSTR
                    LTrim(SearchStringbyPattern)
                    RTrim(SearchStringbyPattern)
                Else
                End If

            End Function

            ''' <summary>
            ''' GO THROUGH EACH CHARACTER AND ' IF PUNCTUATION IE .!?,:'"; REPLACE WITH A SPACE ' IF ,
            ''' OR . THEN CHECK IF BETWEEN TWO NUMBERS, IF IT IS ' THEN LEAVE IT, ELSE REPLACE IT WITH A
            ''' SPACE '
            ''' </summary>
            ''' <param name="STRINPUT">String to be formatted</param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            <System.Runtime.CompilerServices.Extension()>
            Public Function AlphaNumericalOnly(ByRef STRINPUT As String) As String

                Dim A As Short
                For A = 1 To Len(STRINPUT)
                    If Mid(STRINPUT, A, 1) = "." Or
                Mid(STRINPUT, A, 1) = "!" Or
                Mid(STRINPUT, A, 1) = "?" Or
                Mid(STRINPUT, A, 1) = "," Or
                Mid(STRINPUT, A, 1) = ":" Or
                Mid(STRINPUT, A, 1) = "'" Or
                Mid(STRINPUT, A, 1) = "[" Or
                Mid(STRINPUT, A, 1) = """" Or
                Mid(STRINPUT, A, 1) = ";" Then

                        ' BEGIN CHECKING PERIODS AND COMMAS THAT ARE IN BETWEEN NUMBERS '
                        If Mid(STRINPUT, A, 1) = "." Or Mid(STRINPUT, A, 1) = "," Then
                            If Not (A - 1 = 0 Or A = Len(STRINPUT)) Then
                                If Not (IsNumeric(Mid(STRINPUT, A - 1, 1)) Or IsNumeric(Mid(STRINPUT, A + 1, 1))) Then
                                    STRINPUT = Mid(STRINPUT, 1, A - 1) & " " & Mid(STRINPUT, A + 1, Len(STRINPUT) - A)
                                End If
                            Else
                                STRINPUT = Mid(STRINPUT, 1, A - 1) & " " & Mid(STRINPUT, A + 1, Len(STRINPUT) - A)
                            End If
                        Else
                            STRINPUT = Mid(STRINPUT, 1, A - 1) & " " & Mid(STRINPUT, A + 1, Len(STRINPUT) - A)
                        End If

                        ' END CHECKING PERIODS AND COMMAS IN BETWEEN NUMBERS '
                    End If
                Next A
                ' RETURN PUNCTUATION STRIPPED STRING '
                AlphaNumericalOnly = STRINPUT.Replace("  ", " ")
            End Function

            ''' <summary>
            ''' Returns Propercase Sentence
            ''' </summary>
            ''' <param name="TheString">String to be formatted</param>
            ''' <returns></returns>
            <System.Runtime.CompilerServices.Extension()>
            Public Function ProperCase(ByRef TheString As String) As String
                ProperCase = UCase(Left(TheString, 1))

                For i = 2 To Len(TheString)

                    ProperCase = If(Mid(TheString, i - 1, 1) = " ", ProperCase & UCase(Mid(TheString, i, 1)), ProperCase & LCase(Mid(TheString, i, 1)))
                Next i
            End Function

            ''' <summary>
            ''' Capitalizes the text
            ''' </summary>
            ''' <param name="MESSAGE"></param>
            ''' <returns></returns>
            <System.Runtime.CompilerServices.Extension()>
            Public Function CapitaliseTEXT(ByVal MESSAGE As String) As String
                Dim FirstLetter As String = ""
                CapitaliseTEXT = ""
                If MESSAGE = "" Then Exit Function
                FirstLetter = Left(MESSAGE, 1)
                FirstLetter = UCase(FirstLetter)
                MESSAGE = Right(MESSAGE, Len(MESSAGE) - 1)
                CapitaliseTEXT = (FirstLetter + MESSAGE)
            End Function

            ''' <summary>
            ''' Checks if String Contains Letters
            ''' </summary>
            ''' <param name="str"></param>
            ''' <returns></returns>
            <Runtime.CompilerServices.Extension()>
            Public Function ContainsLetters(ByVal str As String) As Boolean

                For i = 0 To str.Length - 1
                    If Char.IsLetter(str.Chars(i)) Then
                        Return True
                    End If
                Next

                Return False

            End Function

            ''' <summary>
            ''' Adds string to end of string (no spaces)
            ''' </summary>
            ''' <param name="Str">base string</param>
            ''' <param name="Prefix">Add before (no spaces)</param>
            ''' <returns></returns>
            <System.Runtime.CompilerServices.Extension()>
            Public Function AddPrefix(ByRef Str As String, ByVal Prefix As String) As String
                Return Prefix & Str
            End Function

            ''' <summary>
            ''' Adds Suffix to String (No Spaces)
            ''' </summary>
            ''' <param name="Str">Base string</param>
            ''' <param name="Suffix">To be added After</param>
            ''' <returns></returns>
            <System.Runtime.CompilerServices.Extension()>
            Public Function AddSuffix(ByRef Str As String, ByVal Suffix As String) As String
                Return Str & Suffix
            End Function

            ''' <summary>
            ''' Returns random character from string given length of the string to choose from
            ''' </summary>
            ''' <param name="Source"></param>
            ''' <param name="Length"></param>
            ''' <returns></returns>
            <Runtime.CompilerServices.Extension()>
            Public Function GetRndChar(ByVal Source As String, ByVal Length As Integer) As String
                Dim rnd As New Random
                If Source Is Nothing Then Throw New ArgumentNullException(NameOf(Source), "Must contain a string,")
                If Length <= 0 Then Throw New ArgumentException("Length must be a least one.", NameOf(Length))
                Dim s As String = ""
                Dim builder As New System.Text.StringBuilder()
                builder.Append(s)
                For i = 1 To Length
                    builder.Append(Source(rnd.Next(0, Source.Length)))
                Next
                s = builder.ToString()
                Return s
            End Function

            ''' <summary>
            ''' Define the search terms. This list could also be dynamically populated at runtime Find
            ''' sentences that contain all the terms in the wordsToMatch array Note that the number of
            ''' terms to match is not specified at compile time
            ''' </summary>
            ''' <param name="TextStr1">String to be searched</param>
            ''' <param name="Words">List of Words to be detected</param>
            ''' <returns>Sentences containing words</returns>
            <Runtime.CompilerServices.Extension()>
            Public Function FindSentencesContaining(ByRef TextStr1 As String, ByRef Words As List(Of String)) As List(Of String)
                ' Split the text block into an array of sentences.
                Dim sentences As String() = TextStr1.Split(New Char() {".", "?", "!"})

                Dim wordsToMatch(Words.Count) As String
                Dim I As Integer = 0
                For Each item In Words
                    wordsToMatch(I) = item
                    I += 1
                Next

                Dim sentenceQuery = From sentence In sentences
                                    Let w = sentence.Split(New Char() {" ", ",", ".", ";", ":"},
                                                   StringSplitOptions.RemoveEmptyEntries)
                                    Where w.Distinct().Intersect(wordsToMatch).Count = wordsToMatch.Count()
                                    Select sentence

                ' Execute the query

                Dim StrList As New List(Of String)
                For Each str As String In sentenceQuery
                    StrList.Add(str)
                Next
                Return StrList
            End Function

            ''' <summary>
            ''' Counts the number of elements in the text, useful for declaring arrays when the element
            ''' length is unknown could be used to split sentence on full stop Find Sentences then again
            ''' on comma(conjunctions) "Find Clauses" NumberOfElements = CountElements(Userinput, delimiter)
            ''' </summary>
            ''' <param name="PHRASE"></param>
            ''' <param name="Delimiter"></param>
            ''' <returns>Integer : number of elements found</returns>
            ''' <remarks></remarks>
            <System.Runtime.CompilerServices.Extension()>
            Public Function CountElements(ByVal PHRASE As String, ByVal Delimiter As String) As Integer
                Dim elementcounter As Integer = 0
                Dim PhraseArray As String()
                PhraseArray = PHRASE.Split(Delimiter)
                elementcounter = UBound(PhraseArray)
                Return elementcounter
            End Function

            ''' <summary>
            ''' counts occurrences of a specific phoneme
            ''' </summary>
            ''' <param name="strIn"></param>
            ''' <param name="strFind"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            <Runtime.CompilerServices.Extension()>
            Public Function CountOccurrences(ByRef strIn As String, ByRef strFind As String) As Integer
                '**
                ' Returns: the number of times a string appears in a string
                '
                '@rem           Example code for CountOccurrences()
                '
                '  ' Counts the occurrences of "ow" in the supplied string.
                '
                '    strTmp = "How now, brown cow"
                '    Returns a value of 4
                '
                '
                'Debug.Print "CountOccurrences(): there are " &  CountOccurrences(strTmp, "ow") &
                '" occurrences of 'ow'" &    " in the string '" & strTmp & "'"
                '
                '@param        strIn Required. String.
                '@param        strFind Required. String.
                '@return       Long.

                Dim lngPos As Integer
                Dim lngWordCount As Integer

                On Error GoTo PROC_ERR

                lngWordCount = 1

                ' Find the first occurrence
                lngPos = InStr(strIn, strFind)

                Do While lngPos > 0
                    ' Find remaining occurrences
                    lngPos = InStr(lngPos + 1, strIn, strFind)
                    If lngPos > 0 Then
                        ' Increment the hit counter
                        lngWordCount = lngWordCount + 1
                    End If
                Loop

                ' Return the value
                CountOccurrences = lngWordCount

PROC_EXIT:
                Exit Function

PROC_ERR:
                MsgBox("Error: " & Err.Number & ". " & Err.Description, , NameOf(CountOccurrences))
                Resume PROC_EXIT

            End Function

            ''' <summary>
            ''' Counts the words in a given text
            ''' </summary>
            ''' <param name="NewText"></param>
            ''' <returns>integer: number of words</returns>
            ''' <remarks></remarks>
            <System.Runtime.CompilerServices.Extension()>
            Public Function CountWords(NewText As String) As Integer
                Dim TempArray() As String = NewText.Split(" ")
                CountWords = UBound(TempArray)
                Return CountWords
            End Function

            ''' <summary>
            ''' Expand a string such as a field name by inserting a space ahead of each capitalized
            ''' letter (where none exists).
            ''' </summary>
            ''' <param name="inputString"></param>
            ''' <returns>Expanded string</returns>
            ''' <remarks></remarks>
            <System.Runtime.CompilerServices.Extension()>
            Public Function ExpandToWords(ByVal inputString As String) As String
                If inputString Is Nothing Then Return Nothing
                Dim charArray = inputString.ToCharArray
                Dim outStringBuilder As New System.Text.StringBuilder(inputString.Length + 10)
                For index = 0 To charArray.GetUpperBound(0)
                    If Char.IsUpper(charArray(index)) Then
                        'If previous character is also uppercase, don't expand as this may be an acronym.
                        If (index > 0) AndAlso Char.IsUpper(charArray(index - 1)) Then
                            outStringBuilder.Append(charArray(index))
                        Else
                            outStringBuilder.Append(String.Concat(" ", charArray(index)))
                        End If
                    Else
                        outStringBuilder.Append(charArray(index))
                    End If
                Next

                Return outStringBuilder.ToString.Replace("_", " ").Trim

            End Function

            ''' <summary>
            ''' checks Str contains keyword regardless of case
            ''' </summary>
            ''' <param name="Userinput"></param>
            ''' <param name="Keyword"></param>
            ''' <returns></returns>
            <Runtime.CompilerServices.Extension()>
            Public Function DetectKeyWord(ByRef Userinput As String, ByRef Keyword As String) As Boolean
                Dim mfound As Boolean = False
                If UCase(Userinput).Contains(UCase(Keyword)) = True Or
                    InStr(Userinput, Keyword) > 1 Then
                    mfound = True
                End If

                Return mfound
            End Function

            ''' <summary>
            ''' Extract words Either side of Divider
            ''' </summary>
            ''' <param name="TextStr"></param>
            ''' <param name="Divider"></param>
            ''' <param name="Mode">Front = F Back =B</param>
            ''' <returns></returns>
            <System.Runtime.CompilerServices.Extension>
            Public Function ExtractWordsEitherSide(ByRef TextStr As String, ByRef Divider As String, ByRef Mode As String) As String
                ExtractWordsEitherSide = ""
                Select Case Mode
                    Case "F"
                        Return ExtractWordsEitherSide(TextStr, Divider, "F")
                    Case "B"
                        Return ExtractWordsEitherSide(TextStr, Divider, "B")
                End Select

            End Function

            <System.Runtime.CompilerServices.Extension()>
            Public Function ExtractFirstWord(ByRef Statement As String) As String
                Dim StrArr() As String = Split(Statement, " ")
                Return StrArr(0)
            End Function

            <System.Runtime.CompilerServices.Extension()>
            Public Function ExtractLastChar(ByRef InputStr As String) As String

                ExtractLastChar = Right(InputStr, 1)
            End Function

            <System.Runtime.CompilerServices.Extension()>
            Public Function ExtractFirstChar(ByRef InputStr As String) As String

                ExtractFirstChar = Left(InputStr, 1)
            End Function

            ''' <summary>
            ''' Returns The last word in String
            ''' NOTE: String ois converted to Array then the last element is extracted Count-1
            ''' </summary>
            ''' <param name="InputStr"></param>
            ''' <returns>String</returns>
            <System.Runtime.CompilerServices.Extension()>
            Public Function ExtractLastWord(ByRef InputStr As String) As String
                Dim TempArr() As String = Split(InputStr, " ")
                Dim Count As Integer = TempArr.Count - 1
                Return TempArr(Count)
            End Function

        End Module
    End Namespace

End Namespace