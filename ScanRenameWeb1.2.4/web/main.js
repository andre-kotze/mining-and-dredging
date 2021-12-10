// Nav
function movehome(){
    window.location.replace("home.html")
}
function moveabout(){
    window.location.replace("about.html")
}
function movesettings(){
    window.location.replace("settings.html");
}
function movebrutal(){
    window.location.replace("bmode.html");
}
function quitprogram(){
    eel.killprogram();
    window.close();
}

// .... 0. Setting functions for concession/path edit
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
    document.getElementById("toremove").style.display = "none";
}
async function add_c(){
    cname = document.getElementById("to_add").value;
    var report = await eel.add_conc(cname)();
    document.getElementById("ctable").innerHTML = report.table;
    window.alert(report.msg);
    document.getElementById("to_add_input").style.display = "none";
}

// --- 1. Detect the number of scans
async function detect(butcolour){
    document.getElementById("detect").innerHTML = "Detecting...";
    document.getElementById("detect").style.backgroundColor = "grey";
    var data = await eel.detect()();
    document.getElementById("numfound").innerHTML = data.count;
    document.getElementById("foundfiles").innerHTML = data.finds;
    document.getElementById("detect").innerHTML = "Detect Files";
    document.getElementById("detect").style.backgroundColor = butcolour;
}

// === 2. Rename the scans
async function rename(){
    document.getElementById("rename").style.display = "none";
    document.getElementById("proglabel").style.display = "inline-block";
    document.getElementById("progbg").style.display = "inline-block";
    document.getElementById("progbar").style.display = "inline-block";
    var data = await eel.rename()();
    document.getElementById("numrenamed").innerHTML = data.count;
    document.getElementById("renamedfiles").innerHTML = data.renames;
    document.getElementById("progbar").style.display = "none"; // new
    document.getElementById("progbg").style.display = "none";
    document.getElementById("proglabel").style.display = "none"; // new
    document.getElementById("rename").style.display = "inline-block";
}
// === 2.1 Brutally rename scans
async function brutalrename(){
    document.getElementById("rename").style.display = "none";
    document.getElementById("proglabel").style.display = "inline-block";
    document.getElementById("progbg").style.display = "inline-block";
    document.getElementById("progbar").style.display = "inline-block";
    var data = await eel.bruteforce()();
    document.getElementById("numrenamed").innerHTML = data.count;
    document.getElementById("renamedfiles").innerHTML = data.renames;
    document.getElementById("progbar").style.display = "none"; // new
    document.getElementById("progbg").style.display = "none";
    document.getElementById("proglabel").style.display = "none"; // new
    document.getElementById("rename").style.display = "inline-block";
}

// :::: 3. Move the files
async function move(butcolour){
    document.getElementById("move").innerHTML = "Moving...";
    document.getElementById("move").style.backgroundColor = "grey";
    var data = await eel.movefiles()();
    document.getElementById("nummoved").innerHTML = data.count;
    document.getElementById("movedfiles").innerHTML = data.moves;
    document.getElementById("move").innerHTML = "Move Files";
    document.getElementById("move").style.backgroundColor = butcolour;
}

// executed on homescreen load to ensure updated concessions table
function refresh_concessions(){
    eel.refresh_concs()();
}

// python signal to update progbar
//eel.expose(jupdateprog);
//function jupdateprog(percent){
//    document.getElementById("progbar").style.width = percent;
//    console.log("progress at " + percent)
//}