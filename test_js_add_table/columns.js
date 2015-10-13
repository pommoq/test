window.onload = function () {
	var columnAreas = document.getElementsByClassName('columns');
	
	if(columnAreas.length > 0) {
		columnAreas.each( function (area) {
			columnify(area, 2);
		});
	}
}

function columnify(area, num) {
	var children = area.childNodes;
	var length = area.innerHTML.length;
	Element.hide(area);
	
	new Insertion.After(area, '<table class="columned"><tr id="columns_temp"></tr></table>');
	for(var i=0;i<num;i++) {
		new Insertion.Bottom($('columns_temp'), '<td id="column_'+i+'" valign="top"></td>');
	}
	
	var column = 0;
	
	for(i=0;i<children.length;i++) {
		if(children[i].nodeType == 3) {
			var text = children[i].nodeValue;
			while(text.length > length/num) {
				pos = Math.round(text.length * (column + 1) / num);
				while(text.substr(pos, 1) != " ") { 
					pos++;
				}
				
				var insert = text.substring(0, pos);
				new Insertion.Top($('column_'+column), insert);
				column++;
				text = text.substring(pos);
			}
			
			new Insertion.Top($('column_'+column), text);
		}
		else {
			new Insertion.Bottom($('column_'+column), children[i]);
			if($('column_'+column).nodeValue.length > length/num) {
				column++;
			}
		}
	}
	
	$('columns_temp').id = "";
	for(i=0;i<num;i++) {
		$('column_'+i).id = "";
	}
}

