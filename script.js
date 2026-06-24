let exportData = [];

const fileInput = document.getElementById("fileInput");
const processBtn = document.getElementById("processBtn");
const downloadBtn = document.getElementById("downloadBtn");
const downloadExcelBtn = document.getElementById("downloadExcelBtn");


// PROCESS BUTTON

processBtn.addEventListener("click",()=>{

const file = fileInput.files[0];

if(!file){

alert("Please select CSV file");

return;

}


const reader = new FileReader();

reader.onload = function(e){

const data = new Uint8Array(e.target.result);

const workbook = XLSX.read(data,{type:"array"});

const sheetName = workbook.SheetNames[0];

const worksheet = workbook.Sheets[sheetName];

const json = XLSX.utils.sheet_to_json(

worksheet,

{header:1}

);


exportData=[];

let currentState="";


json.forEach((row,index)=>{


if(index===0) return;

if(!row || !row[0]) return;


const fullName = String(row[0]).trim();

let firstName="";
let lastName="";


// FORMAT:

// Delgado, Maria

if(fullName.includes(",")){

const parts = fullName.split(",");

lastName = parts[0].trim();

firstName = parts[1]
? parts[1].trim()
: "";

}


// FORMAT:

// Maria Delgado

else{

const parts = fullName.split(" ");

lastName = parts.pop() || "";

firstName = parts.join(" ");

}



const position = row[1] || "";

const email = row[2] || "";

const school = row[3] || "";

const district = row[4] || "";



// AUTO STATE

if(row[5] && row[5].toString().trim()!==""){

currentState = row[5].toString().trim();

}

const state = currentState;



exportData.push({

firstName,

lastName,

position,

email,

school,

district,

state

});

});



// PREVIEW

const preview = exportData.slice(0,10);

let html="";

preview.forEach(item=>{

html += `

<tr>

<td>${item.firstName}</td>

<td>${item.lastName}</td>

<td>${item.email}</td>

<td>${item.school}</td>

<td>${item.district}</td>

<td>${item.state}</td>

</tr>

`;

});


document.getElementById("resultBody").innerHTML = html;



// STATUS

document.getElementById("status").innerHTML=`

<h3>✅ EBE Data Cleaner Complete</h3>

<p>

Total Contacts :

<strong>${exportData.length}</strong>

</p>

<p>

CSV Ready For Download

</p>

`;



// DASHBOARD

document.getElementById("totalContacts").innerText=

exportData.length;



const schools = new Set(

exportData.map(x=>x.school)

);

document.getElementById("totalSchools").innerText=

schools.size;



const districts = new Set(

exportData.map(x=>x.district)

);

document.getElementById("totalDistricts").innerText=

districts.size;



};


reader.readAsArrayBuffer(file);


});




// DOWNLOAD CSV

downloadBtn.addEventListener("click",()=>{


let csv =

"First Name,Last Name,Position,Email,School,District,State\n";


exportData.forEach(item=>{


csv +=

`"${item.firstName}","${item.lastName}","${item.position}","${item.email}","${item.school}","${item.district}","${item.state}"\n`;



});


const blob = new Blob([csv],{

type:"text/csv"

});


const url = URL.createObjectURL(blob);


const a = document.createElement("a");


a.href=url;

a.download="EBE_Cleaned_Data.csv";

a.click();


URL.revokeObjectURL(url);


});




// DOWNLOAD EXCEL


downloadExcelBtn.addEventListener("click",()=>{


const worksheet = XLSX.utils.json_to_sheet(exportData);


const workbook = XLSX.utils.book_new();


XLSX.utils.book_append_sheet(

workbook,

worksheet,

"Cleaned Data"

);


XLSX.writeFile(

workbook,

"EBE_Cleaned_Data.xlsx"

);


});




// DRAG DROP


const dropArea = document.getElementById("dropArea");

const browseBtn = document.getElementById("browseBtn");


browseBtn.addEventListener("click",()=>{

fileInput.click();

});


fileInput.addEventListener("change",()=>{

document.getElementById("fileName").innerText=

fileInput.files[0].name;

});



dropArea.addEventListener("dragover",(e)=>{

e.preventDefault();

dropArea.classList.add("dragover");

});


dropArea.addEventListener("dragleave",()=>{

dropArea.classList.remove("dragover");

});


dropArea.addEventListener("drop",(e)=>{

e.preventDefault();

dropArea.classList.remove("dragover");


fileInput.files = e.dataTransfer.files;


document.getElementById("fileName").innerText=

fileInput.files[0].name;


});