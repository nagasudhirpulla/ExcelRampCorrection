//var datebody = document.getElementById("htxt_Day_ID");
//datebody.selectedIndex = 4;
//datebody.onchange();
var schtable = document.getElementById("table4").getElementsByTagName("tbody")[0];
var str = "";
var len = schtable.rows[2].cells.length
for(var i = 0; i < 98; i++) {
	if(i!=1){
		for(var j = 0; j < len; j++) {
			temp = schtable.rows[i].cells[j].innerHTML.trim();
			temp = temp.replace(/ /g,'');
        		str += temp + " ";
			
		}
		str += "\n";
	}
}
if(!document.getElementById("CONSOLE")){
	var newDiv = document.createElement("textarea"); 
	newDiv.setAttribute("id", "CONSOLE");
	newDiv.style.width = "600px";
	newDiv.style.height = "600px"
	var newContent = document.createTextNode(str); 
	newDiv.appendChild(newContent); //add the text node to the newly created div. 
	document.body.appendChild(newDiv);
} else{
	document.getElementById("CONSOLE").innerHTML = str;
}
