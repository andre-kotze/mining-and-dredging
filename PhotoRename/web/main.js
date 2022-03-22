// === 0. META
// --- Functions for navigation of program
// using JavaScript to move between pages instead of html, to allow exit function to work
function movehome(){
    window.location.replace("home.html");
}
function moveabout(){
    window.location.replace("about.html");
}
function movesettings(){
    window.location.replace("settings.html");
}

// === 1. DETECT
// get last used concession and path

// list concession options in selection box
async function list_concessions(){
    let conc_list = await eel.conc_options()();
    let path = await eel.lastpath()();
    document.getElementById("currentconc").innerHTML = conc_list;
    document.getElementById("currentpath").innerHTML = path;
}


// --- Detect the number of photos in ToBeRenamed folder
async function detect(){
    document.getElementById("detect").innerHTML = "Detecting...";
    document.getElementById("detect").style.backgroundColor = "grey";
    let selectedconcession = document.getElementById("currentconc").value;
    var data = await eel.detect(selectedconcession)();
    document.getElementById("num_detected").innerHTML = data.num;
    document.getElementById("currentpath").innerHTML = data.path;
    document.getElementById("currentsl").innerHTML = data.sl_file;
    document.getElementById("detect").innerHTML = "Detect Files";
    document.getElementById("detect").style.backgroundColor = "darkblue";
}

// -- called from python to update path == # NOT WORK
//eel.expose(jupdatepath);
//function jupdatepath(path){
//    document.getElementById("currentpath").innerHTML = path;
//}
function sorttable(sldata) {
    var data = JSON.parse(sldata);
    var table = "<table border=1>";
    var item = "";
    var count = 0;
    for (row in data) {
        if ((data[row].Count == 0 && data[row].Smpl_photos == 'Y') || (data[row].Count > 0 && data[row].Smpl_photos == 'N')) {
            item = '<tr style="color:red">';
        } else {
            item = "<tr>";
        }
        item += '<td style="width: 10%">' + data[row].Seq + "</td>";
        item += '<td style="width: 15%">' + data[row].Sample_ID + "</td>";
        item += "<td>" + data[row].Start_Datetime.slice(0, 16).replace("T", "_") + "</td>";
        if (data[row].End_Datetime != null) {
            item += "<td>" + data[row].End_Datetime.slice(0, 16).replace("T", "_") + "</td>";
        } else {
            item += "<td>" + data[row].End_Datetime + "</td>";
        }
        item += '<td style="width: 10%">' + data[row].Count + "</td>";
        count += data[row].Count
        item += '<td style="width: 10%">' + data[row].Smpl_photos + "</td>";
        item += "</tr>";
        table += item;
    }
    table += "</table>";
    document.getElementById("sorttable").innerHTML = table;
    document.getElementById("numbersorted").innerHTML = count + " photos sorted";
    document.getElementById("sortbutton").innerHTML = "Sort Photos";
    document.getElementById("sortbutton").style.backgroundColor = "green";
}

// === 2. SORT
async function sort(){
    document.getElementById("sortbutton").innerHTML = "Sorting...";
    document.getElementById("sortbutton").style.backgroundColor = "grey";
    let sldata = await eel.read_screenlog()();
    sorttable(sldata);
}
async function rename(){
    document.getElementById("renamebutton").innerHTML = "Renaming...";
    document.getElementById("renamebutton").style.backgroundColor = "grey";
    var result = await eel.rename()();
    document.getElementById("renamecard").style.display = "block";
    document.getElementById("renamedlog").innerHTML = result.log;
    document.getElementById("numberrenamed").innerHTML = result.renamed;
    document.getElementById("numbergenned").innerHTML = result.genned;
    document.getElementById("renamebutton").innerHTML = "FINISHED";
    document.getElementById("renamebutton").style.backgroundColor = "black";
}
// --- Function for killing program
// --- JavaScript kills user interface while Python kills program in terminal
function quitprogram(){
    eel.killprogram();
    window.close();
}


// setting functions for add/remove concession paths:
// ALSO ADD "UPDATE" FEATURE
async function gettable(){
    var table = await eel.gettable()();
    document.getElementById("ctable").innerHTML = table;
}

async function markremove(){
    document.getElementById("toremove").style.display = "inline";
    let conc_list = await eel.conc_options()();
    document.getElementById("toremove").innerHTML = conc_list;
    document.getElementById("cremove").setAttribute("onclick", "remove_c()");
}
function get_new_c(){
    document.getElementById("to_add_input").style.display = "inline";
    document.getElementById("cadd").setAttribute("onclick", "add_c()");
}

async function remove_c(){
    cname = document.getElementById("toremove").value;
    var report = await eel.remove_conc(cname)();
    document.getElementById("ctable").innerHTML = report.table;
    window.alert(report.msg);
}
async function add_c(){
    cname = document.getElementById("to_add").value;
    var report = await eel.add_conc(cname)();
    document.getElementById("ctable").innerHTML = report.table;
    window.alert(report.msg);
}




