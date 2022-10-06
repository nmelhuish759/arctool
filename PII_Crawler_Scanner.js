var currDate = new Date()
var badWords = {
	//BIRTH
			//BIRTH
			"dob":"Detected the word DOB (Date of Birth). This is PII and must be redacted before uploading to a PII container.",
			"date of birth":"Detected the words Date of Birth. This is PII and must be redacted before uploading to a PII container.",
			"place of birth":"Detected the words Place of Birth. This is PII and must be redacted before uploading to a PII container.",
			//ADDRESS
			"address":"Detected the word Address. Home addresses are PII and must be redacted before uploading to a PII container.",
			"ship to":"Detected the word Ship to. Home addresses are PII and must be redacted before uploading to a PII container.",
			"bill to":"Detected the word Bill to. Home addresses are PII and must be redacted before uploading to a PII container.",
			//"email":"Detected the word Email. Personal Email is PII",
			//CLEARANCE
			"secret":"Detected the word Secret. Clearance Levels are PII and must be uploaded to a PII container.",
			"top secret":"Detected the word Top Secret. Clearance Levels are PII and must be uploaded to a PII container.",
			"security clearance":"Detected the words Security Clearance. Clearance Levels are PII and must be uploaded to a PII container.",
			"poly":"Detected the word Poly. Clearance Levels are PII and must be uploaded to a PII container.",
			"classified":"Detected the word Classified. Classifided data is PII and must be uploaded to a PII container.",
			//ARREST RECORDS
			"police":"Detected the word Police. Arrest Records are PII and must be redacted before uploading to a PII container.",
			"arrest":"Detected the word Arrest. Arrest Records are PII and must be redacted before uploading to a PII container.",
			"dui":"Detected the word DUI. Arrest Records are PII and must be redacted before uploading to a PII container.",
			"dwi":"Detected the word DWI. Arrest Records are PII and must be redacted before uploading to a PII container.",
			"felony":"Detected the word Felony. Arrest Records are PII and must be redacted before uploading to a PII container.",
			"misdemeanor":"Detected the word Misdemeanor. Arrest Records are PII and must be redacted before uploading to a PII container.",
			"detained":"Detected the word Detained. Arrest Records are PII and must be redacted before uploading to a PII container.",
			//LICENSE/STATE
			"driver's license number":"Detected the word Driver's License Number. This is considered PII and must be redacted before uploading to a PII container.",
			"license number":"Detected the word License Number. This is considered PII and must be redacted before uploading to a PII container.",
			"license":"Detected the word License. This is considered PII and must be redacted before uploading to a PII container.",
			"state identification number":"Detected the words State Identification Number. This is considered PII and must be redacted before uploading to a PII container.",
			"state identification":"Detected the words State Identification. This is considered PII and must be redacted before uploading to a PII container.",
			"state id number":"Detected the words State ID Number. This is considered PII and must be redacted before uploading to a PII container.",
			"state id":"Detected the words State ID. This is considered PII and must be redacted before uploading to a PII container.",
			//PERSONAL COMPUTER DATA
			"password":"Detected the word Password. Username/Passwords are PII and must be uploaded to a PII container.",
			"username":"Detected the word Username. Username/Passwords are PII and must be uploaded to a PII container.",
			"login name":"Detected the word Login Name. Username/Passwords are PII and must be uploaded to a PII container.",
			//"ip address": "Detected the word IP Address. IP Addresses are PII",
			//SSN
			"social security":"Detected the word Social Security. Social Security numbers are PII and must be redacted before uploading to a PII container",
			"ssn":"Detected the word SSN. Social Security numbers are PII and must be redacted before uploading to a PII container",
			//RACE/GENDER/REGLIGION
			"race":"Detected the word Race. Religion, race, nationality, and sexual orientation are PII and must be redacted before uploading to a PII container",
			"nationality":"Detected the word Nationality. Religion, race, nationality, and sexual orientation are PII and must be redacted before uploading to a PII container",
			"ethnicity":"Detected the word Ethnicity. Religion, race, nationality, and sexual orientation are PII and must be redacted before uploading to a PII container",
			"sexual orientation":"Detected the word Sexual Orientation. Religion, race, nationality, and sexual orientation are PII and must be redacted before uploading to a PII container",
			"female":"Detected the word Race. Religion, race, nationality, and sexual orientation are PII and must be redacted before uploading to a PII container",
			"male":"Detected the word Race. Religion, race, nationality, and sexual orientation are PII and must be redacted before uploading to a PII container",
			"sex":"Detected the word Sex. Religion, race, nationality, and sexual orientation are PII and must be redacted before uploading to a PII container",
			"gender":"Detected the word Gender. Religion, race, nationality, and sexual orientation are PII and must be redacted before uploading to a PII container",
			"religion": "Detected the word Religion. This is considered PII and must be redacted before uploading to a PII container.",
			//BANKING/PAY
			"aba": "Detected the word ABA. Routing numbers/Bank numbers are PII and must be redacted before uploading to a PII container",
			"pay": "Detected the word Pay. Pay Grade with a name is considered PII and must be uploaded to a PII container.",
			"pay grade": "Detected the word Pay Grade. Pay Grade with a name is considered PII and must be uploaded to a PII container.",
			"base pay": "Detected the word Base Pay. Pay Grade with a name is considered PII and must be uploaded to a PII container.",
			"basic pay": "Detected the word Basic Pay. Pay Grade with a name is considered PII and must be uploaded to a PII container.",
			"name": "Detected the word Name. Pay Grade with a name is considered PII and must be uploaded to a PII container.",
			"allowance": "Detected the word Allowance. Pay Grade with a name is considered PII and must be uploaded to a PII container.",
			"entitlement": "Detected the word Entitlement. This is considered PII and must be uploaded to a PII container.",
			"account": "Detected the word Account. Financial data is considered PII and must be redacted before uploading to a PII container.",
			"account number": "Detected the words Account Number. Financial data is considered PII and must be redacted before uploading to a PII container.",
			"salary": "Detected the word Salary. Financial data is considered PII and must be redacted before uploading to a PII container.",
			"internal revenue service": "Detected the words internal revenue service. Financial data is considered PII and must be redacted before uploading to a PII container.",
			"settlement amount": "Detected the words Settlement Amount. Financial data is considered PII and must be redacted before uploading to a PII container.",
			//FOREIGN
			"alien registration number": "Detected the word Alien Registration Number. This is considered PII and must be redacted before uploading to a PII container.",
			"alien registration": "Detected the word Alien Registration. This is considered PII and must be redacted before uploading to a PII container.",
			"alien number": "Detected the word Alien Number. This is considered PII and must be redacted before uploading to a PII container.",
			"alien": "Detected the word Alien. This is considered PII and must be redacted before uploading to a PII container.",
			"citizenship": "Detected the word Citizenship. This is considered PII and must be redacted before uploading to a PII container.",
			//PASSPORT
			"passport number": "Detected the word Passport Number. This is considered PII and must be redacted before uploading to a PII container.",
			"passport": "Detected the word Passport. This is considered PII and must be redacted before uploading to a PII container.",
			//FAMILY
			"maiden name": "Detected the word Maiden Name. Family data is PII and must be redacted before uploading to a PII container.",
			"maiden": "Detected the word Maiden. Family data is PII and must be redacted before uploading to a PII container.",
			"son": "Detected the word Son. Family data is PII and must be redacted before uploading to a PII container.",
			"daughter": "Detected the word Daughter. Family data is PII and must be redacted before uploading to a PII container.",
			"wife": "Detected the word Wife. Family data is PII and must be redacted before uploading to a PII container.",
			"husband": "Detected the word Husband. Family data is PII and must be redacted before uploading to a PII container.",
			"partner": "Detected the word Partner. Family data is PII and must be redacted before uploading to a PII container.",
			"significant other": "Detected the word Significant Other. Family data is PII and must be redacted before uploading to a PII container.",
			"death record" : "Detected the word Death Record. This is PII and must be redacted before uploading to a PII container.",		
			//BIOMETRIC DATA
			"hospital": "Detected the word Hospital. Biometric data is PII and must be redacted before uploading to a PII container.",
			"fingerprint": "Detected the word Fingerprint. Biometric data is PII and must be redacted before uploading to a PII container.",
			"pharmacy": "Detected the word Pharmacy. Biometric data is PII and must be redacted before uploading to a PII container.",
			"pharmaceutical": "Detected the word Pharmaceutical. Biometric data is PII and must be redacted before uploading to a PII container.",
			"prescription": "Detected the word Prescription. Biometric data is PII and must be redacted before uploading to a PII container.",
			"prescribe": "Detected the word Prescribe. Biometric data is PII and must be redacted before uploading to a PII container.",
			"walter reed": "Detected the word Walter Reed. Biometric data is PII and must be redacted before uploading to a PII container.",
			"rehabilitation": "Detected the word Rehabilitation. Health information is considered PII and must be redacted before uploading to a PII container.",
			"rehab": "Detected the word Rehab. Health information is considered PII and must be redacted before uploading to a PII container.",
			"blood type": "Detected the word Blood Type. Health information is considered PII and must be redacted before uploading to a PII container.",
			"blood": "Detected the word Blood. Health information is considered PII and must be redacted before uploading to a PII container.",
			//EDUCATION DATA
			"school": "Detected the word School. Education level is considered PII and must be uploaded to a PII container.",
			"university": "Detected the word University. Education level is considered PII and must be uploaded to a PII container.",
			"college": "Detected the word College. Education level is considered PII and must be uploaded to a PII container.",
			"technical training": "Detected the words Technical Training. Education level is considered PII and must be uploaded to a PII container.",
			"academic" : "Detected the word Academic. Education level is considered PII and must be uploaded to a PII container.",
			//OFFICIAL USE
			"for official use only": "Detected the words 'For Official Use Only'. Please review the document for CUI.",
			"fouo": "Detected the words 'FOUO'. Please review the document for CUI.",	
			//Personnel Records
			"inspector general" : "Detected the words Inspector General. Inspector General Protected Information is considered PII and must be uploaded to a PII container.",
			"priig" : "Detected the word PRIIG. Inspector General Protected Information is considered PII and must be uploaded to a PII container.",
			"military personnel records" : "Detected the words Military Personnel Records. Military Personnel Records are considered PII and must be uploaded to a PII container.",
			"mpr" : "Detected the word MPR. Military Personnel Records are considered PII and must be uploaded to a PII container.",
			"civilian personnel records" : "Detected the words Civilian Personnel Records. Civilian Personnel Records are considered PII and must be uploaded to a PII container.",
			"cpr" : "Detected the word CPR. Civilian Personnel Records are considered PII and must be uploaded to a PII container."
	
}
var badWordKeys = Object.keys(badWords)
var badWordLen = badWordKeys.length


function scan(text,link){
	var detected = []
	var SSN,Alien,Credit,Phone,PDFRandomNumbers,Email,date;
	var removeDashes = text.replace(/\-/g,"") //removing the dashes 
	var SNNRegex = (/\d{3}(\s|-)\d{2}(\s|-)\d{4}|\XXX(\s|-)XX(\s|-)\d{4}|\xxx(\s|-)xx(\s|-)\d{4}/g)
	//var AlienRegex = (/\d{3}(\s|)-(\s|)\d{3}(\s|)-(\s|)\d{3}|\XXX(\s|)-(\s|)XXX(\s|)-(\s|)\d{3}|\xxx(\s|)-(\s|)xxx(\s|)-(\s|)\d{3}/g)
	//var CreditRegex = (/\d{4}((\s|)(\s|-)(\s|))\d{4}((\s|)(\s|-)(\s|))\d{4}((\s|)(\s|-)(\s|))\d{4}/g)
	// Matches Visa, MasterCard, American Express, Diners Club, Discover, and JCB cards 
	var CreditRegex = (/\b(?:4[0-9]{12}(?:[0-9]{3})?|[25][1-7][0-9]{14}|6(?:011|5[0-9][0-9])[0-9]{12}|3[47][0-9]{13}|3(?:0[0-5]|[68][0-9])[0-9]{11}|(?:2131|1800|35\d{3})\d{11})\b/g)
	var PhoneRegex = (/^[\+]?[(]?[0-9]{3}[)]?[-\s\.]?[0-9]{3}[-\s\.]?[0-9]{4,6}$/g)
	var PDFRandomNumbersRegex = (/\b\d{9}\b/g) //get 9 or 16 digits in a row.  have the CC locked down so removing it
	var EmailRegex = (/[a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}/g);
	//var dateRegex = (/\d{1,2}(\/|-|\.)\d{1,2}(\/|-|\.)([0-9]{2}$|[0-9]{4})|([0-9]{2}$|[0-9]{4})(\/|-|\.)\d{1,2}(\/|-|\.)\d{1,2}/g)
	var dateRegex = (/(0[1-9]|1[0-2])\D(0[1-9]|1\d|2\d|3[01])\D(19|20)\d{2}/g)
	
	
	if(SNNRegex.test(text))
		var SSN = text.match(SNNRegex) // matches XXX-XX-1234 or 123-45-6789 also if there are any spaces between the "-"
	//if(AlienRegex.test(StringData))
		//var Alien = StringData.match(AlienRegex) // matches XXX-XXX-123 or 123-456-789 also if there are any spaces between the "-"			
	if(CreditRegex.test(removeDashes))
		var Credit = removeDashes.match(CreditRegex) // matches all major credit cards with no spaces or dashes
	if(PhoneRegex.test(text))
		var Phone = text.match(PhoneRegex)
	if(PDFRandomNumbersRegex.test(text))
		var PDFRandomNumbers = text.match(PDFRandomNumbersRegex) //checking other patterns
	//var Email = PDFText.match(/\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})/g); //matches test@blah.com this just kills my cpu and just causes my browser to crash?!?!?!
	//var Email = PDFText.match(/(([^<>()\[\]\.,;:\s@\"]+(\.[^<>()\[\]\.,;:\s@\"]+)*)|(\".+\"))@(([^<>()[\]\.,;:\s@\"]+\.)+[^<>()[\]\.,;:\s@\"]{2,})/g); //for some reason the pdf parser is trying to parse some pictures and messing this up
	if(EmailRegex.test(text))
		var Email = text.match(EmailRegex);
	if(dateRegex.test(text))
		var date = text.match(dateRegex)


	var detectedPay = false;
	var detectedName = false;
	
	for(; i < badWordLen; i++){
		var word = badWordKeys[i]
		if(lowerCaseStr.indexOf(" "+word+" ") > -1){ //space in front and back so it doesnt detect nate in dominate
			switch(word){
				case "name":
					detectedName = true
				break;
				case "pay":
				case "pay grade":
				case "base pay":
				case "basic pay":
					detectedPay = true
				break;
				
				default:
					detected.push("• "+badWords[word]+"<br>")
				break;
			}
			if(detectedPay && detectedName){
				detected.push("• Detected the word Name and Pay. Pay grade along with a name is considered PII<br>")
			}
		}				
	}
	
	if(SSN){
		for(var i=0;i<SSN.length;i++){
			detected.push("• SSN has been detected. SSN: "+SSN[i]+"<br>")
		}
	}
	/*if(Alien){
		for(var i=0;i<Alien.length;i++){
			detected.push("• Alien Registration Number has been detected. Alien number: "+Alien[i]+"<br>")
		}
	}
	if(IP){
		for(var i=0;i<IP.length;i++){
			detected.push("• IP Address detected. IP Address: "+IP[i]+"<br>")
		}
	}*/
	if(Credit){
		for(var i=0;i<Credit.length;i++){
			detected.push("• Credit Card number detected. Credit Card Number: "+Credit[i]+"<br>")
		}
	}
	if(Phone){
		for(var i=0;i<Phone.length;i++){
			detected.push("• Phone number has been detected. Personal phone numbers are considered PII. Phone: "+Phone[i]+"<br>")
		}
	}
	if(Email){
		for(var i=0;i<Email.length;i++){
			var lowerCaseEmail = Email[i].toLowerCase()
			if(lowerCaseEmail.indexOf(".mil") == -1 && lowerCaseEmail.indexOf("@ey.com") == -1 && lowerCaseEmail.indexOf("@delloite.com") == -1 && lowerCaseEmail.indexOf("@bah.com") == -1 && lowerCaseEmail.indexOf("@kpmg.com") == -1) //if not a .mil account
				detected.push("• Email Address has been detected. Personal Email addresses are considered PII. Email: "+Email[i]+"<br>")
		}
	}
	if(date){
		for(var i=0;i<date.length;i++){
			var selectedDate = new Date(date[i]);
			
			if(!isNaN(Date.parse(selectedDate))){
				//selectedDate.setFullYear(selectedDate.getFullYear()+18); //if date is over 18 years old then im assuming it is a DOB
				//var isDOB = selectedDate < new Date()
				var selectedDateYear = selectedDate.getFullYear();
				var isOver18 = currDate.getFullYear()-18
				var isOver120 = currDate.getFullYear()-120
				var isDOB = isOver18 >= selectedDateYear && isOver120 <= selectedDateYear
				
				if(isDOB)
					detected.push("• Date of birth has been detected. Date of birth is considered PII. DOB: "+date[i]+"<br>")
			}
		}
	}
	postMessage({"detected":detected,"link":link})
}


onmessage = function(e){
	if(e.origin.indexOf('chrome-extension') == -1)
		scan(e.data.text, e.data.link)
}

PIICrawlerScanner = {} //for loading