function onOpen() {
    var ui = SpreadsheetApp.getUi();
    // Or DocumentApp or FormApp.
    ui.createMenu('Custom Script Menu')
        .addItem('importData()', 'menuItem1')
        .addItem('nameValFilter()', 'menuItem2')
        .addToUi();
}

function menuItem1() {
    return importData();
}

function menuItem2() {
    return nameValFilter();
}

var Alpha = ['Seng Kit', 'Gladson','Shane','Adrian', 'Bryan Yon', 'Suren', 'Clement Chang', 'Jie Qi', 'Wei Keong', 'Jun ming','Weihian','Sara', 'Hua Zong', 'Bellamy', 'He Qun', 'Colin Ho', 'Jeff','Aloysius', 'Zheng Yuan', 'Bryant Lim', 'Hari', 'Owen Lee', 'Zinne','Marcus Lim','Vincze', 'Raj', 'Ashok', 'Manfred', 'Clement Lim', 'Riston', 'OC', 'MAJ JAMES', 'CPT DAVID', 'CPT JERRY', 'CPT LAURA', 'CPT MARC']; //this name order has to correspond with AlphaRanked!

var AlphaRanked = ['LTA SENG KITT', '2LT GLADSON','2LT SHANE','3WO ADRIAN', 'SSG BRYAN YON', '2SG SUREN', '3SG CLEMENT CHANG', '3SG JIE QI', 'LTA WEI KEONG', 'LTA JUN MING','2LT WEIHIAN','2WO SARA', '3WO HUA ZONG', 'SSG BELLAMY', '3SG HE QUN', '3SG COLIN', 'LTA JEFFERY','2LT ALOYSIUS','3WO ZHENG YUAN','2SG BRYANT', '2SG HARI', '3SG OWEN', 'LTA ZINN-E','2LT MARCUS','2SG VINCZE', '3WO RAJ', 'MSG ASHOK', 'MSG MANFRED', '3SG CLEMENT LIM', '3SG RISTON', 'OC', 'MAJ JAMES', 'CPT DAVID', 'CPT JERRY', 'CPT LAURA', 'CPT MARC'];

//personnel involved in SB, tech SB and duty
var crew = ['Seng Kit','Gladson','Shane','Jiajie','Adrian','Bryan Yon', 'Suren', 'Clement Chang', 'Jie Qi', 'Wei Keong', 'Jun ming', 'Sara', 'Hua Zong', 'Bellamy', 'He Qun', 'Colin Ho', 'Jeff','Aloysius','Jun Rong','Yashi','Zheng Yuan', 'Bryant Lim', 'Hari', 'Owen Lee', 'Zinne','Wong Wei','Vincze', 'Raj', 'Ashok', 'Clement Lim', 'Riston', 'Rik', 'Dave', 'Ivan', 'Bryan C', 'Daniel', 'Yao Ren', 'Zech', 'Kenneth', 'Cheok', 'Yewseng', 'Weng Yew', 'Yihao', 'Rongpo', 'Keith', 'Jason', 'Kai Feng', 'Marcus Chia','Marcus Lim','Andy', 'Kevin', 'Mathew', 'Qixian', 'Fredrick', 'Loheng', 'Javin','Jen','Kan Wu','Weihian','Denzel','Mike', 'Pravean', 'Kityin', 'Manfred', 'Jian hong', 'Ramanen', 'Elijah', 'Xavier'];

var crewRanked = ['LTA SENG KITT', '2LT GLADSON','2LT SHANE','2LT JIAJIE','3WO ADRIAN', 'SSG BRYAN YON', '2SG SUREN', '3SG CLEMENT CHANG', '3SG JIE QI', 'LTA WEI KEONG', 'LTA JUN MING', '2WO SARA', '3WO HUA ZONG', 'SSG BELLAMY', '3SG HE QUN', '3SG COLIN', 'LTA JEFFERY','2LT ALOYSIUS','2LT JUN RONG','2LT YASHI','3WO ZHENG YUAN', '2SG BRYANT', '2SG HARI', '3SG OWEN', 'LTA ZINN-E','2LT WONG WEI','2SG VINCZE', '3WO RAJ', 'MSG ASHOK', '3SG CLEMENT LIM', '3SG RISTON', 'LTA RIK', '3WO DAVE', 'MSG IVAN', '2SG BRYAN CHUA', '3WO DANIEL', 'MSG YAOREN', '3SG ZECH', '3SG KENNETH', '3WO CHEOK', 'SSG YEWSENG', '3WO WENG YEW', 'MSG YIHAO', '3WO RONGPO', '3SG KEITH', '3SG JASON', '3SG KAIFENG', 'LTA MARCUS','2LT MARCUS','2WO ANDY', '2WO KEVIN', '3WO MATHEW', '2SG QIXIAN', '1SG FREDRICK', '3WO LOHENG', '3SG JAVIN','LTA JENEVIEVE','LTA KAN WU','2LT WEIHIAN','2LT DENZEL','1WO MIKE', '3WO PRAVEAN', 'SSG KITYIN', 'MSG MANFRED', '1SG JIANHONG', '2SG RAMANEN', '2SG ELIJAH', '3SG XAVIER'];

var alphaTeamA = ['LTA SENG KITT','2LT GLADSON','3WO ADRIAN', 'SSG BRYAN YON', '2SG SUREN', '2SG HARI', '3SG JIE QI', 'LTA WEI KEONG', 'LTA JUN MING','2LT WEIHIAN','2WO SARA', '3WO HUA ZONG', 'SSG BELLAMY', '3SG HE QUN', '3SG COLIN', 'OC', 'CPT JERRY'];
var alphaTeamB = ['LTA JEFFERY','2LT ALOYSIUS','2LT SHANE','3WO ZHENG YUAN', '2SG BRYANT', '3SG CLEMENT CHANG', '3SG OWEN', 'LTA ZINN-E', '2SG VINCZE', '3WO RAJ', 'MSG ASHOK', 'MSG MANFRED', '3SG CLEMENT LIM', '3SG RISTON', 'MAJ JAMES', 'CPT DAVID', 'CPT LAURA', '2LT MARCUS'];

//Organising and outputting data
//takes an array of arrays and converts an array with strings only
function cleanArray(arr) {
    var string;
    var valuesF = []
    arr.forEach(function (e) {
        string = e.toString();
        valuesF.push(string);
    });
    return valuesF;
};

function orderByRank(arr) {
    var alphaRankedOrdered = [];
    var i;
    var x;
    var ranks = ['OC', 'MAJ', 'CPT', 'LTA', '2LT', '1WO', '2WO', '3WO', 'MSG', 'SSG', '1SG', '2SG', '3SG'];
    for (i = 0; i < ranks.length; i++) {
        for (x = 0; x < arr.length; x++) {
            if (arr[x].substring(0, 3) == ranks[i]) {
                alphaRankedOrdered.push(arr[x]);
            }
        }
    }
    return alphaRankedOrdered;
}

function addNewline(arr) {
    return String(arr).split(",").join("%0A");
}

var PRESENT = [];
var DYME = [];
var OFF = [];
var FFI = [];
var CCL = [];
var PCL = [];
var RSO = [];
var MC = [];
var CL = [];
var LL = [];
var WFH = [];
var CSE = [];
var MA = [];
var OS = ['CPT LAURA (BMTC) (280921-041221)'];
var CO = [];
var HL = [];
var UL = [];
var UNKNOWN = ['OC', 'MAJ JAMES', 'CPT DAVID', 'CPT JERRY','CPT MARC'];
var crewX = [];
var crewSB = [];
var crewTECHSB = [];

function nameValFilter(telinput, inputNum) {
    var i;
    var x;

    //Update sheet and name of sheet here
    var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/19H0CbJsgs4fovu1Ys-WyxrzTUdKfddJZksiyc4znpaY/edit#gid=1370368481"); //imported sheet
    var sheet = ss.getSheetByName("Oct 21 (BCP)");

    var values = sheet.getSheetValues(1, 1, 100, 1);
    var valuesFixed = cleanArray(values);

    for (i = 0; i < valuesFixed.length; i++) {
        for (x = 0; x < Alpha.length; x++) {
            if (Alpha[x] === valuesFixed[i]) {
                var status = sheet.getSheetValues(i + 1, telinput + 1, 1, 1);
                status = cleanArray(status);
                var exp = String(status[0]);

                switch (exp) {
                    case 'X':
                        DYME.push(AlphaRanked[x]);
                        break;
                    case 'OFF':
                    case 'OFF ':
                    case 'OIL':
                        OFF.push(AlphaRanked[x]);
                        break;
                    case 'LL':
                        LL.push(AlphaRanked[x]);
                        break;
                    case 'WFH':
                        WFH.push(AlphaRanked[x]);
                        break;
                    case 'CSE':
                    case 'DLS CSE':
                    case 'MSV CSE':
                        CSE.push(AlphaRanked[x] + " (" + exp + ")");
                        break;
                    case 'MA':
                        MA.push(AlphaRanked[x]);
                        break;
                    case 'ESCORT':
                        OS.push(AlphaRanked[x] + " (" + exp + ")");//SPECIAL CASE TO BE SETTLED
                        break;
                    case '\\': //"\" is a special character that needs bypassing!!
                    case '/':
                        CO.push(AlphaRanked[x]);
                        break;
                    case 'FFI':
                        FFI.push(AlphaRanked[x]);
                        break;
                    case 'CCL':
                        CCL.push(AlphaRanked[x]);
                        break;
                    case 'PCL':
                        PCL.push(AlphaRanked[x]);
                        break;
                    case 'RSO':
                        RSO.push(AlphaRanked[x]);
                        break;
                    case 'MC':
                        MC.push(AlphaRanked[x]);
                        break;
                    case 'CL':
                        CL.push(AlphaRanked[x]);
                        break;
                    case 'HL':
                        HL.push(AlphaRanked[x]);
                        break;
                    case 'UL':
                        UL.push(AlphaRanked[x]);
                        break;
                    case '':
                        //exact same code as the one directly below for WFH, just without additional brackets
                        if (telinput % 2 == 1 && alphaTeamA.includes(AlphaRanked[x])) {
                            WFH.push(AlphaRanked[x]);
                        } else if (telinput % 2 == 0 && alphaTeamB.includes(AlphaRanked[x])) {
                            WFH.push(AlphaRanked[x]);
                        } else {
                            PRESENT.push(AlphaRanked[x]);
                        }
                        break;
                    case 'SB':
                    case 'TECH SB':
                        //modify below code according to WFH schedule
                        if (telinput % 2 == 1 && alphaTeamA.includes(AlphaRanked[x])) {
                            WFH.push(AlphaRanked[x] + " (" + exp + ")");
                        } else if (telinput % 2 == 0 && alphaTeamB.includes(AlphaRanked[x])) {
                            WFH.push(AlphaRanked[x] + " (" + exp + ")");
                        } else {
                            PRESENT.push(AlphaRanked[x] + " (" + exp + ")");
                        }
                        break;
                    default:
                        exp = exp.replace(/\\/g, "/");
                        UNKNOWN.push(AlphaRanked[x] + " (" + exp + ")");
                }

            }
        }
        var crewstatus = sheet.getSheetValues(i + 1, telinput + 1, 1, 1);
        crewstatus = cleanArray(crewstatus);
        var crewexp = String(crewstatus[0]);
        if ((crewexp.includes('X') && !(crewexp.includes('MX'))) || crewexp.includes('SB') || crewexp.includes('TECH SB')) {
            for (var a = 0; a < crew.length; a++) {
                if (crew[a] == valuesFixed[i]) {
                    if (crewexp.includes('X')) {
                        crewX.push(crewRanked[a]);
                    } else if (crewexp.includes('TECH SB')) {
                        crewTECHSB.push(crewRanked[a]);
                    } else if (crewexp.includes('SB')) {
                        crewSB.push(crewRanked[a]);
                    }
                }
            }
        }
    }


    var PRESENTordered = orderByRank(PRESENT);
    var DYMEordered = orderByRank(DYME);
    var OFFordered = orderByRank(OFF);
    var FFIordered = orderByRank(FFI);
    var CCLordered = orderByRank(CCL);
    var PCLordered = orderByRank(PCL);
    var RSOordered = orderByRank(RSO);
    var MCordered = orderByRank(MC);
    var CLordered = orderByRank(CL);
    var LLordered = orderByRank(LL);
    var WFHordered = orderByRank(WFH);
    var CSEordered = orderByRank(CSE);
    var MAordered = orderByRank(MA);
    var OSordered = orderByRank(OS);
    var COordered = orderByRank(CO);
    var HLordered = orderByRank(HL);
    var ULordered = orderByRank(UL);
    var UNKNOWNordered = orderByRank(UNKNOWN);
    var crewXOrdered = orderByRank(crewX);
    var crewSBOrdered = orderByRank(crewSB);
    var crewTECHSBOrdered = orderByRank(crewTECHSB);

    var PRESENTsplit = addNewline(PRESENTordered);
    var DYMEsplit = addNewline(DYMEordered);
    var OFFsplit = addNewline(OFFordered);
    var FFIsplit = addNewline(FFIordered);
    var CCLsplit = addNewline(CCLordered);
    var PCLsplit = addNewline(PCLordered);
    var RSOsplit = addNewline(RSOordered);
    var MCsplit = addNewline(MCordered);
    var CLsplit = addNewline(CLordered);
    var LLsplit = addNewline(LLordered);
    var WFHsplit = addNewline(WFHordered);
    var CSEsplit = addNewline(CSEordered);
    var MAsplit = addNewline(MAordered);
    var OSsplit = addNewline(OSordered);
    var COsplit = addNewline(COordered);
    var HLsplit = addNewline(HLordered);
    var ULsplit = addNewline(ULordered);
    var UNKNOWNsplit = addNewline(UNKNOWNordered);
    var longLine = "---------------------------------------------------";

    //Final text formatting

    return "TOTAL STRENGTH " + "(" + Alpha.length + ")" + "%0A%0A" +
        "PRESENT: " + "(" + PRESENT.length + ")" + "%0A" + String(PRESENTsplit) + "%0A%0A" +
        "DYME: " + "(" + DYME.length + ")" + "%0A" + String(DYMEsplit) + "%0A%0A" +
        "OFF: " + "(" + OFF.length + ")" + "%0A" + String(OFFsplit) + "%0A%0A" +
        "C/O: " + "(" + CO.length + ")" + "%0A" + String(COsplit) + "%0A%0A" +
        "O/S: " + "(" + OS.length + ")" + "%0A" + String(OSsplit) + "%0A%0A" +
        "CSE: " + "(" + CSE.length + ")" + "%0A" + String(CSEsplit) + "%0A%0A" +
        "WFH: " + "(" + WFH.length + ")" + "%0A" + String(WFHsplit) + "%0A%0A" +
        "LL: " + "(" + LL.length + ")" + "%0A" + String(LLsplit) + "%0A%0A" +
        "MA: " + "(" + MA.length + ")" + "%0A" + String(MAsplit) + "%0A%0A" +
        "MC: " + "(" + MC.length + ")" + "%0A" + String(MCsplit) + "%0A%0A" +
        "RSO: " + "(" + RSO.length + ")" + "%0A" + String(RSOsplit) + "%0A%0A" +
        "CCL: " + "(" + CCL.length + ")" + "%0A" + String(CCLsplit) + "%0A%0A" +
        "PCL: " + "(" + PCL.length + ")" + "%0A" + String(PCLsplit) + "%0A%0A" +
        "HL: " + "(" + HL.length + ")" + "%0A" + String(HLsplit) + "%0A%0A" +
        "UL: " + "(" + UL.length + ")" + "%0A" + String(ULsplit) + "%0A%0A" +
        "CL: " + "(" + CL.length + ")" + "%0A" + String(CLsplit) + "%0A%0A" +
        "FFI: " + "(" + FFI.length + ")" + "%0A" + String(FFIsplit) + "%0A%0A" +
        "UNKNOWN: " + "(" + UNKNOWN.length + ")" + "%0A" + String(UNKNOWNsplit) + "%0A%0A" + longLine + "%0A%0A" + longLine + "%0A" +
        "[DUTY CREW FOR " + String(inputNum) + "]" + "%0A" + "OSC: " + String(crewXOrdered[0]) + "%0A" + "DYOSC: " + String(crewXOrdered[1]) + "%0A" + "ADSS: " + String(crewXOrdered[2]) + "%0A" + "ADSS: " + String(crewXOrdered[3]) + "%0A" + "ADWS: " + String(crewXOrdered[4]) + "%0A" + "ADWS: " + String(crewXOrdered[5]) + "%0A%0A" +
        "[STANDBY CREW FOR " + String(inputNum) + "]" + "%0A" + "AWO: " + String(crewSBOrdered[0]) + "%0A" + "ADWS: " + String(crewSBOrdered[1]) + "%0A%0A" +
        "[TECH SB CREW FOR " + String(inputNum) + "]" + "%0A" + "AWO: " + String(crewTECHSBOrdered[0]) + "%0A" + "ADSS: " + String(crewTECHSBOrdered[1]) + "%0A" + "ADWS: " + String(crewTECHSBOrdered[2]);

}
