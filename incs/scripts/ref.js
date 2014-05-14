window.onload = function(){

	var questionFieldset = document.getElementById('questions'),
		questionParas = questionFieldset.getElementsByTagName('p'),
		//refTable = document.getElementsByClassName('.ref')
		row  = document.getElementById('templateRow'),//refTable.getElementsByTagName('tbody')[0].getElementsByTagName('tr')[0];
		parasLength = questionParas.length;

	for(i=0; i<parasLength; i++){ //} in questionParas){
		var para = questionParas[i],
			rowClone = row,//row.cloneNode(true),
			inputs = rowClone.getElementsByTagName('input'),
			label = rowClone.getElementsByTagName('label')[0],
			number = rowClone.getElementsByTagName('td')[0],
			questionName = 'question'+i,
			questionLabel= para.innerText || para.innerHTML;	//Firefox doesn't support innerText


		for(j=0; j<inputs.length; j++){
			inputs[j].setAttribute('name', questionName);
			inputs[j].setAttribute('data-index', i);
		}
			label.setAttribute('for', questionName);
			label.innerHTML = questionLabel;
			number.innerHTML = i+1;

		rowClone = rowClone.cloneNode(true);
		insertBeforeElement(row, rowClone);

		//row.insertAdjacentElement('beforeBegin', rowClone);
		 //(para.remove || para.removeNode).call(para);

	}

	for (i=parasLength-1; i>=0; i--){
		removeElement(questionParas[i]);
	}

	removeElement(row);


	questionFieldset.addEventListener('change',function(e){

			//alert('input changed');
			var input = e.target,
				row = input.parentNode.parentNode,	// tr -> td -> input
				value = input.value,//.getAttribute('value')
				index = parseInt(input.getAttribute('data-index')),
				correctValue = questionAnswers[index], //e.target.getAttribute('data-index'),
				gotItRight = value == correctValue.answer;

				//console.log(value);
				//console.log(correctValue);
				if (gotItRight){
					addClass(row, 'right');
					removeClass(row, 'wrong');

				}else{
					addClass(row, 'wrong');
					removeClass(row, 'right');

				}
	});

	//row.remove();


};

function addClass(e, c){
	if (e.className.indexOf(c)>-1) return;	//Already has it
	else e.className = e.className + ' ' + c;

};
function removeClass(e, c){
	//Leaves multiple spaces, but...
	e.className = e.className.replace(new RegExp('\\b' + c + '\\b'), '');
	
};


function removeElement(e){
	if (e.remove) e.remove();
	else e.parentNode.removeChild(e);		//IE8 http://www.quirksmode.org/dom/core/
	//else if (e.removeNode) e.removeNode();
	return;

	//(e.remove || e.removeNode).call(e);
}

function insertBeforeElement(b,a){
	//Firefox doesn't support insertAdjacentElement;
	//IE doesn't support insertBefore
	if (b.insertAdjacentElement) b.insertAdjacentElement('beforeBegin', a); //(e.remove) e.remove();
	else (b.parentNode.insertBefore(a,b));		//IE8 http://www.quirksmode.org/dom/core/
	//else if (e.removeNode) e.removeNode();
	return;

	//(e.remove || e.removeNode).call(e);
}

var no = 'no',
	yes = 'yes',
	questionAnswers = [
		 {answer:no},
		 {answer:yes},
		 {answer:no},
		 {answer:no},
		 {answer:yes},
		 {answer:no},
		 {answer:yes},
		 {answer:yes},
		 {answer:yes},
		 {answer:no},
		 {answer:yes},
		 {answer:yes},
		 {answer:no},
		 {answer:yes},
		 {answer:no},
		 {answer:yes},
		 {answer:yes},
		 {answer:yes},
		 {answer:no},
		 {answer:no},
		 {answer:no},
		 {answer:yes}
	];

//if (document.readyState == "complete") 
//http://stackoverflow.com/questions/799981/document-ready-equivalent-without-jquery/800010#800010
