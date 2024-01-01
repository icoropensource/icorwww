/*********************************************************************
JavaScript 1.2 Validation Script (IE and Netscape)
	version 3.1.0
	by matthew frank

There are no warranties expressed or implied.  This script may be
re-used and distrubted freely provided this header remains intact
and all supporting files are included (unaltered) in the distribution:

		validation.js   - this file
		validation.htm  - example form
		readme.htm      - directions on using this script
		(validation-3.1.0.zip contains all files)

If you are interested in keeping up with the latest releases of this
script or asking questions about its implementation, think about joining
the Yahoo! Groups discussion forum dedicated to validation:

	http://groups.yahoo.com/group/validation

*********************************************************************/
/*====================================================================
Function: Validation
Purpose:  Custom object constructor.
Inputs:   None
Returns:  undefined
====================================================================*/
function Validation(){
	var validationFunctions = new Array;
	var isDOMCompliant = !!document.getElementById;
	var isInitialFocusSet = false;
	/*====================================================================
	Function: Err
	Purpose:  Custom object constructor
	Inputs:   None
	Returns:  undefined
	====================================================================*/
	this.Err=function(){
		/*********************************************************************
		Method:  Err.raise
		Purpose: Gives visual warning to user about all errors contained in
		         the Error object
		Inputs:  source  - element object
		         message - error message to user
		         stem    - identifier which defines the proper MESSAGE attribute
		         event   - event which triggered validation (e.g. onchange)
		Returns: undefined
		*********************************************************************/
		this.raise=function(source,message,stem,event){
			if(typeof event!="object") event=new Object;
			if(event.type!="change" || event.type=="change" && !suppressOnchange(source)){
	 			var sName=getProperty(source,"DISPLAY-NAME");
				var sMsg=choose((stem?getProperty(source,stem.toUpperCase()+"-MESSAGE"):null),getProperty(source,"MESSAGE"),message.replace(/\.$/,"") + (sName?"\nin the <"+sName+"> field.":"."));
				alert(sMsg);
				// Execute onvalidatefocus processing
				toFunction(getProperty(source,"onvalidatefocus"))();
				if(source.focus) source.focus();
				if(source.select) source.select();
			}
			paintElement(source);
		}
	}
	/*********************************************************************
	Function: setProperty
	Purpose:  Set property value of element
	Inputs:   object   - object reference
	          property - string
	          value    - value to give property
	Returns:  undefined
	*********************************************************************/
	var setProperty=function(object,property,value){
		if(object.setAttribute)
			object.setAttribute(property,value);
		else
			object[property]=value;
	}
	/*********************************************************************
	Function: getProperty
	Purpose:  Retrieve property value of element
	Inputs:   object   - object reference
	          property - string
	Returns:  variant
	NOTE: IE supports case-sensitivity with the getAttribute method.
	*********************************************************************/
	var getProperty=function(object, property){
		var returnValue;
		if(object){
			returnValue = object[property];
			if(!propertyOn(returnValue) && object.getAttribute)
				returnValue=object.getAttribute(property);
		}
		return returnValue;
	}
	/*********************************************************************
	Function: propertyOn
	Purpose:  Check that at least one property is turned on (defined, non-null, non-false)
	Inputs:   attribute list - values of attributes (e.g. results of getProperty call)
	Returns:  boolean
	*********************************************************************/
	var propertyOn=function(/* attribute list */){
		var isOn,attribute,length=arguments.length;
		for(var i=0;i<length;i++){
			attribute = arguments[i];
			if(typeof attribute=="string")attribute=attribute.toLowerCase();
			isOn = typeof attribute!="undefined" &&
			       attribute!=null &&
				   attribute!="false" &&
				   attribute!="off" &&
			       attribute!=false;
			if(isDOMCompliant) isOn &= (attribute!="");
			if(isOn) break;
		}
		return !!isOn;
	}
	/*********************************************************************
	Function: trim
	Purpose:  Remove leading and trailing spaces
	Inputs:   string
	Returns:  string
	*********************************************************************/
	var trim=function(string){
		// NN4 does not handle /^s+|\s+$/g correctly
		return string.replace(/^\s+/,"").replace(/\s+$/,"");
	}
	/*********************************************************************
	Function: pad
	Purpose:  Left pad a number with zeros to a given width
	Inputs:   value, width (defaults to 2)
	Returns:  numeric string of width characters
	*********************************************************************/
	var pad=function(value, width){
		width=choose(width,2);
		var returnValue=value.toString();
		for(var i=width-value.length;i>0;i--)
			returnValue="0"+returnValue;
		return returnValue;
	}
	/*********************************************************************
	Function: toFunction
	Purpose:  Transforms value into function if it is not
	Inputs:   value - function pointer
	Returns:  function
	*********************************************************************/
	var toFunction=function(value){
		if(value){
			if(typeof value!="function")
				value = new Function(value);
		}else
			value = function(){return;}
		return value;
	}
	/*********************************************************************
	Function: minMaxRange
	Purpose:  Helps to build an appropriate error message
	Inputs:   min, max
	Returns:  string
	*********************************************************************/
	var minMaxRange=function(min,max){
		if(propertyOn(min))	min=""+min;
		if(propertyOn(max)) max=""+max;
		if(!!min&&!!max)
			return " between "+min+" and "+max;
		else if(!!min)
			return " greater than or equal to "+min;
		else if(!!max)
			return " less than or equal to "+max;
		else return "";
	}
	/*********************************************************************
	Function: dateOrTime
	Purpose:  Determines whether a data type is date, time or datetime
	          Accepted tokens: d dd m mm yyyy hh hh24 nn ss ap
	Inputs:   format
	Returns:  string
	*********************************************************************/
	var dateOrTime=function(format){
		var date=false,time=false;
		with(format){
			date=search(/mm?/i)>-1||search(/dd?/i)>-1||search(/yyyy/i)>-1;
			time=search(/hh?/i)>-1||search(/nn/i)>-1||search(/ss/i)>-1||search(/ap/i)>-1;
		}
		return (date?"date":"")+(time?"time":"");
	}
	/*********************************************************************
	Function: toDate
	Purpose:  Converts string to date based on format
	          Accepted tokens: d dd m mm yyyy hh hh24 nn ss ap
	Inputs:   date, format
	Returns:  string (YYYYMMDDHH24NNSS), undefined indicates failure to match format
	*********************************************************************/
	var toDate=function(date,format){
		var i, regex, index=new Array;
		var day, month, year, hour, minute, second, ampm;
		// Determine order of datetime tokens
		with(format){
			index[search(/dd?/i)]="day";
			index[search(/mm?/i)]="month";
			index[search(/yyyy/i)]="year";
			index[search(/hh/i)]="hour";
			index[search(/nn/i)]="minute";
			index[search(/ss/i)]="second";
			index[search(/ap/i)]="ampm";

			// timing of replaces is quite important!
			regex=format.replace(/(\$|\^|\*|\(|\)|\+|\.|\?|\\|\{|\}|\||\[|\])/g,"\\$1");
			// only allow one pass for day and month
			if(search(/dd/i)>-1)
				regex=regex.replace(/dd/i,"(0[1-9]|[1-2]\\d|3[0-1])");
			else
				regex=regex.replace(/d/i,"(0?[1-9]|[1-2]\\d|3[0-1])");
			if(search(/mm/i)>-1)
				regex=regex.replace(/mm/i,"(0[1-9]|1[0-2])");
			else
				regex=regex.replace(/m/i,"(0?[1-9]|1[0-2])");
			regex=regex.replace(/nn/i,"([0-5]\\d)")
				.replace(/ss/i,"([0-5]\\d)")
				.replace(/yyyy/i,"(\\d{4})")
				.replace(/\s+/g,"\\s*");
			if(search(/hh24/i)>-1)
				regex=regex.replace(/hh24/i,"([0-1]\\d|2[0-3])");
			else
				regex=regex.replace(/hh/i,"(0\\d|1[0-2])").replace(/ap/i,"([ap]m?)");
		}
		// test date against format
		if(!new RegExp("^"+regex+"$","i").test(date))
			return;
		// set values from user input
		year=month=day=hour=minute=second=0,ampm="";
		for(var key=0,i=0;key<index.length;key++)
			if(index[key]) eval(index[key]+"=RegExp.$"+ ++i);
		// set default values
		if(hour<12&&/^pm?$/i.test(ampm))
			hour+=12;
		else if(hour==12&&/^am?$/i.test(ampm))
			hour=0;
		if(year==0) year=1;
		if(month==0) month=1;
		if(day==0) day=1;
		// check days in month
		if(month==2 && day>((year%4==0&&year%100!=0||year%400==0)?29:28) || day>((month-1)%7+1)%2+30)
			return;
		// build result value
		return ""+pad(year,4)+pad(month)+pad(day)+pad(hour)+pad(minute)+pad(second);
	}
	/*********************************************************************
	Function: formatNumber
	Purpose:  Format an integer for display
	Inputs:   i - integer value as string
	Returns:  string
	*********************************************************************/
	var formatNumber=function(i){
		if(typeof i=="undefined"||i==null) return null;
		var sEnd=(/\./.test(i=i.toString()))?"\\.":"$";
		var re=new RegExp("(\\d)(\\d{3})(,|"+sEnd+")");
		if(re.test(i)) i=formatNumber(i.replace(re, "$1,$2$3"));
		return i;
	}
	/*********************************************************************
	Function: choose
	Purpose:  Return the first non-null, non-empty parameter value
	Inputs:   variant list
	Returns:  variant
	*********************************************************************/
	var choose=function(/* list */){
		var i, value, iArguments=arguments.length;
		for(i=0;i<iArguments;i++){
			value = arguments[i];
			if(typeof value!="undefined" && value!=null && value!="")
				return value;
		}
	}
	/*********************************************************************
	Function: suppressOnchange
	Purpose:  Determine whether the onchange validation message should appear
	Inputs:   element - the validated element
	Returns:  boolean
	*********************************************************************/
	var suppressOnchange=function(element){
		return propertyOn(getProperty(element,"SUPPRESS-ONCHANGE-MESSAGE"),getProperty(element.form,"SUPPRESS-ONCHANGE-MESSAGE"))
	}
	/*********************************************************************
	Function: getValueOf
	Purpose:  Return the value of a form field as seen at server
	Inputs:   element - form field
	Returns:  variant
	*********************************************************************/
	var getValueOf=function(element){
		var returnValue=null;
		switch (element.type){
			case "text" : case "textarea" : case "file" : case "password" : case "hidden" :
				returnValue=element.value;
				break;
			case "select-one" :
				if(element.selectedIndex>=0)
					returnValue=element.options[element.selectedIndex].value;
				break;
			case "select-multiple" :
				for(var i=0,iOptions=element.options.length; i<iOptions; i++)
					if(element.options[i].selected && trim(element.options[i].value.toString())){
						returnValue=true;
						break;
					}
				break;
			case "radio" : case "checkbox" :
				returnValue=element.checked;
				break;
			default:
				returnValue = null;
		}
		return returnValue;
	}
	/******************************************************************************
	Function: paintElement
	Purpose:  Give visual cue to user that field has failed validation
	Inputs:   element - form field
	Returns:  undefined
	******************************************************************************/
	var paintElement=function(element){
		if(element.style){
			var backgroundColor = choose(getProperty(element,"INVALID-COLOR"),getProperty(element.form,"INVALID-COLOR"));
			if (backgroundColor){
				// Paint element by changing color directly
				setProperty(element,"OLD-BG-COLOR",element.style.backgroundColor);
				element.style.backgroundColor = backgroundColor;
			}else{
				// Paint element by changing class
				setProperty(element,"OLD-CLASS-NAME",element.className);
				if(isDOMCompliant){
					element.className = trim(element.className+" invalid");
				}else{
					element.className = "invalid";
				}
			}
		}
	}
	/******************************************************************************
	Function: restoreForm
	Purpose:  Restore all elements in FORM to original style
	Inputs:   form   - HTML form object
	          bReset - boolean flag: call the onreset handler for the form's elements
	Returns:  undefined
	******************************************************************************/
	var restoreForm=function(form, bReset){
		for(var i=0,elements=form.elements,iElements=elements.length;i<iElements;i++){
			restoreElement(elements[i]);
			if(bReset)elements[i].onreset();
		}
	}
	/******************************************************************************
	Function: restoreElement
	Purpose:  Restore element to original style
	Inputs:   element - form field
	Returns:  undefined
	******************************************************************************/
	var restoreElement=function(element){
		element.__=null;
		if(element.style){
			var backgroundColor=getProperty(element,"OLD-BG-COLOR");
			if (backgroundColor!=null && typeof backgroundColor!="undefined") {
				// Revert to previous background color
				element.style.backgroundColor = backgroundColor;
				setProperty(element,"OLD-BG-COLOR",null);
			}else{
				var oldClass=getProperty(element,"OLD-CLASS-NAME");
				if (oldClass!=null && typeof oldClass!="undefined"){
					// Revert to previous class
					element.className=oldClass;
					setProperty(element,"OLD-CLASS-NAME",null);
				}
			}
		}
	}
	/******************************************************************************
	Function: isValidForm
	Purpose:  Validate a form's fields based on the attributes provided in the HTML text
	Inputs:   form - HTML form object
	Returns:  boolean
	******************************************************************************/
	var isValidForm=function(form,event){
		var i,iElements,orderBy,position;
		var element,elementList = form.elements;
		restoreForm(form);
		// Execute onbeforevalidate processing
		if(form.onbeforevalidate()==false)
			return false;
		// Determine proper order of validation
		orderBy = getProperty(form,"ORDERED-VALIDATION");
		if(propertyOn(orderBy)){
			orderBy = /^tabindex$/i.test(orderBy)?"tabIndex":"VALIDATION-ORDER";
			elementList = new Array();
			for(i=0,iElements=form.elements.length;i<iElements;i++){
				element=form.elements[i];
				position=parseInt(getProperty(element,orderBy));
				if(propertyOn(position) && !isNaN(position))
					elementList=elementList.slice(0,position).concat(element,elementList.slice(position));
				else
					elementList[elementList.length]=element;
			}
		}
      // Validate individual elements
      for(i=0,iElements=elementList.length;i<iElements;i++) {
         if (elementList[i].tagName!='FIELDSET') {
            if (!elementList[i].validate(event)) return false;
         }
      }
		// Execute onaftervalidate processing
		if(form.onaftervalidate()==false)
			return false;
		return true;
	}
	/******************************************************************************
	Function: isValidElement
	Purpose:  Validate an element based on the attributes provided in the HTML text
	Inputs:   element - form field
	Returns:  boolean
	******************************************************************************/
	var isValidElement=function(element,event){
		// Do not validate label or fieldset elements
		if(typeof element.type == "undefined"||element.__||propertyOn(getProperty(element,"disabled"))||propertyOn(getProperty(element,"readOnly")))
			return true;
		element.__=true;
		var iLength,vAnd,or,oRegexp,iMin,iMax,format,sMask,date,minDate,maxDate,bFirst;
		iMin=getProperty(element,"MIN");
		iMax=getProperty(element,"MAX");
		var pass=true;
		// Trim leading and trailing spaces
		if(element.value) element.value = trim(element.value);
		var sValue=getValueOf(element);
		var bSigned=getProperty(element,"SIGNED");
		// Execute onbeforevalidate processing
		if(element.onbeforevalidate()==false)
			return false;
		// REQUIRED
		if(propertyOn(getProperty(element,"REQUIRED")) && !sValue){
			Validation.Err.raise(element, "Please enter a value", "REQUIRED",event);
	 		return false;
		}
		// FLOAT, NUMBER
		if(((bFirst=propertyOn(getProperty(element,"FLOAT")))||propertyOn(getProperty(element,"NUMBER"))) && sValue){
			if(!new RegExp("^"+(propertyOn(bSigned)?"-?":"")+"(\\d*(,?\\d{3})*\\.?\\d+|\\d+(,?\\d{3})*\\.?\\d*)$").test(sValue))
				pass=false;
			else if(iMin==parseFloat(iMin) && parseFloat(sValue.replace(/,/g,""))<parseFloat(iMin))
				pass=false;
			else if(iMax==parseFloat(iMax) && parseFloat(sValue.replace(/,/g,""))>parseFloat(iMax))
				pass=false;
			if(!pass){
				Validation.Err.raise(element, "Please enter a "+(bSigned?"":"positive ")+"number"+minMaxRange(formatNumber(iMin),formatNumber(iMax)),bFirst?"FLOAT":"NUMBER",event);
				return false;
			}
		}
		// AMOUNT, CURRENCY
		else if (((bFirst=propertyOn(getProperty(element,"AMOUNT")))||propertyOn(getProperty(element,"CURRENCY"))) && sValue){
			// NN4 does not recognize \d{1,3} correctly
			var reMain = "((\\d{1,3})(,?\\d{3})*(\\.\\d{2})?|((\\d{1,3})(,?\\d{3})*)?\\.\\d{2})";
			if(!new RegExp("^("+
			              "("+ (bSigned?"(\\$?\\-?|\\-?\\$?)":"\\$?")+reMain +")"+
			              (bSigned?"|(\\(\\$?"+reMain+"\\))":"")
			              +")$").test(sValue))
				pass=false;
			else if(iMin==parseFloat(iMin) && parseFloat(sValue.replace(/[\$,]/g,"").replace(/^\(\$(.*)\)$/,"-$1"))<parseFloat(iMin))
				pass=false;
			else if(iMax==parseFloat(iMax) && parseFloat(sValue.replace(/[\$,]/g,"").replace(/^\(\$(.*)\)$/,"-$1"))>parseFloat(iMax))
				pass=false;
			if(!pass){
				Validation.Err.raise(element, "Please enter a "+(bSigned?"":"positive ")+"dollar amount"+minMaxRange(formatNumber(iMin),formatNumber(iMax)),bFirst?"AMOUNT":"CURRENCY",event);
				return false;
			}
		}
		// INTEGER
		else if (propertyOn(getProperty(element,"INTEGER")) && sValue){
			// NN4 does not recognize \d{1,3} correctly
			if (!new RegExp("^"+((bSigned)?"-?":"")+"(\\d|\\d{2}|\\d{3})(,?\\d{3})*$").test(sValue))
				pass=false;
			else if(iMin==parseInt(iMin) && parseInt(sValue.replace(/,/g,""))<parseInt(iMin))
				pass=false;
			else if(iMax==parseInt(iMax) && parseInt(sValue.replace(/,/g,""))>parseInt(iMax))
				pass=false;
			if(!pass){
				Validation.Err.raise(element, "Please enter a "+(bSigned?"n ":"positive ")+"integer"+minMaxRange(formatNumber(iMin),formatNumber(iMax)),"INTEGER",event);
				return false;
			}
		}
		// DATE, DATETIME
		else if(((bFirst=propertyOn((format=getProperty(element,"DATE"))))||propertyOn((format=getProperty(element,"DATETIME"))))&&sValue){
			// Set default date format
			if(format==""||typeof format!="string") format="M/D/YYYY";
			minDate=toDate(iMin,format);
			maxDate=toDate(iMax,format);
			if(typeof (date=toDate(sValue,format))=="undefined")
				pass=false;
			else if(propertyOn(iMin)&&typeof minDate!="undefined"&&date<minDate)
				pass=false;
			else if(propertyOn(iMax)&&typeof maxDate!="undefined"&&date>maxDate)
				pass=false;
			if(!pass){
				Validation.Err.raise(element,"Please enter a "+dateOrTime(format)+" value"+minMaxRange(iMin,iMax)+" in the proper format:\n\t"+format.replace(/ap/i,"AM/PM").toUpperCase(),bFirst?"DATE":"DATETIME",event);
				return false;
			}
		}
		// PHONE
		else if(propertyOn(getProperty(element,"PHONE"))&&sValue){
			var sPhone=sValue.replace(/\D/g,"");
			var iDigits=sPhone.length;
			if(!(iDigits==10||iDigits==11&&/^1/.test(sPhone))){
				Validation.Err.raise(element,"Please enter a valid phone number","PHONE",event);
				return false;
			}
		}
		// EMAIL
		else if(propertyOn(getProperty(element,"EMAIL"))&&sValue){
			if(!/^[\w_-]+(\.[\w_-]+)*@[\w_-]+(\.[\w_-]+)*\.[a-z]{2,4}$/i.test(sValue)){
				Validation.Err.raise(element,"Please enter a valid email address","EMAIL",event);
				return false;
			}
		}
		// ZIP Code
		else if(propertyOn(getProperty(element,"ZIP"))&&sValue){
			if(!/^\d{5}(-?\d{4})?$/.test(sValue)){
				Validation.Err.raise(element, "Please enter a valid ZIP code","ZIP",event);
				return false;
			}
		}
		// REGEXP
		if(propertyOn(oRegexp=getProperty(element,"REGEXP"))&&sValue){
			if(oRegexp.constructor!=RegExp)
				oRegexp = new RegExp(oRegexp, "i");
			if(!oRegexp.test(sValue)){
				Validation.Err.raise(element, "Please enter a valid value", "REGEXP",event);
				return false;
			}
		}
		// MAXLENGTH
		if(sValue&&(iLength=getProperty(element,"maxLength"))&&!/\D/.test(iLength)&&sValue.length>iLength){
			Validation.Err.raise(element,"Please enter a value having no more than " + formatNumber(iLength) + " characters", "MAXLENGTH",event);
			return false;
		}
		// Check field against run-time validation functions
		for(var i=0;validationFunctions[i];i++){
			if(validationFunctions[i](element, sValue)==false)
				return false;
		}
		// AND
		if(propertyOn(vAnd=getProperty(element,"AND"))&&!sValue){
			// If not an array, create one
			if(typeof vAnd == "string") vAnd=vAnd.toString().split(/,/);
			// Require each element in the list
			for(var oNewElement,i=0,iFields=vAnd.length; i<iFields; i++){
				if((oNewElement=(typeof vAnd[i].form=="object")?vAnd[i]:element.form.elements[trim(vAnd[i])])){
					if(!!getValueOf(oNewElement)){
						Validation.Err.raise(element, "Please enter a value", "AND",event);
						return false;
					}
				}
			}
		}
		// OR
		if((or=getProperty(element,"OR"))&&!sValue){
			var fields,message;
			if(or.constructor==Array||or.constructor==String)
				fields=or;
			else
				fields=or["fields"];
			if(fields){
				// If not an array, create one
				if(fields.constructor!=Array)
					fields=fields.toString().split(/,/);
				// Require each element in the list
				for(var oNewElement,i=0,iFields=fields.length,bValue; !bValue&&i<iFields; i++){
					oNewElement=(fields[i].form)?fields[i]:element.form.elements[trim(fields[i])];
					if(oNewElement) bValue = !!getValueOf(oNewElement);
				}
				if(!bValue){
					Validation.Err.raise(element, message?message:"Please enter a value", "OR",event);
					return false;
				}
			}
		}
		// Perform onvalidate event handler
		if(element.onvalidate()==false)
			return false;
		// Execute onaftervalidate processing
		if(element.onaftervalidate()==false)
			return false;
		// element is valid
		return true;
	}
	/*********************************************************************
	Method:  Validation.add
	Purpose: Adds global validation functions at run-time
	Inputs:  code    - function to be used in validation for all fields
	         message - default error message for a given function
	         stem    - identifier for validation check used in finding error message
	Returns: undefined
	*********************************************************************/
	this.add=function(code){
		if(typeof code=="function"){
			eval("code="+code.toString());
			validationFunctions[validationFunctions.length]=code;
		}
	}
	/*********************************************************************
	Method:  Validation.setup
	Purpose: Set up methods and event handlers for all forms and elements
	Inputs:  none
	Returns: undefined
	*********************************************************************/
	this.setup=function(){
		var initialFocus;
		// Fan through forms on page to perform initializations
		for(var i=0,oForm,iForms=document.forms.length; i<iForms; i++){
			oForm=document.forms[i];
			if(!oForm._){
				// Capture current event handlers
				oForm._onsubmit_=toFunction(getProperty(oForm,"onsubmit"));
				oForm._onreset_=toFunction(getProperty(oForm,"onreset"));
				// Create methods on form
				oForm.validate=function(oEvent){return isValidForm(this,oEvent);}
				oForm.onbeforevalidate = toFunction(getProperty(oForm,"onbeforevalidate"));
				oForm.onaftervalidate = toFunction(getProperty(oForm,"onaftervalidate"));
				oForm.onautosubmit = toFunction(getProperty(oForm,"onautosubmit"));
				// Replace onsubmit event handler
				oForm.onsubmit=function(oEvent){ //NN passed event
					if(!this.validate(choose(oEvent,window.event))) return false;
					// Perform original onsubmit event handler
					if (this._onsubmit_ && this._onsubmit_(oEvent)==false){
						return false;
					}
					return true;
				}
				// Replace onreset event handler
				oForm.onreset=function(oEvent){
					restoreForm(this, true);
					// Perform original onreset event handler
					if (this._onreset_ && this._onreset_(oEvent)==false){
						return false;
					}
				}
				// Create markRequired method
				oForm.markRequired=function(){
					var oldClassName, element;
					for(var k=0;(element=this.elements[k]);k++){
						oldClassName = getProperty(element, "OLD-CLASS-NAME");
						if(propertyOn(getProperty(element,"REQUIRED"))){
							if(!/\brequired\b/i.test(element.className)){
								if(propertyOn(oldClassName))
									setProperty(element,"OLD-CLASS-NAME",isDOMCompliant?trim(oldClassName+" required"):"required");
								else{
									if(isDOMCompliant)
										element.className = trim(element.className+" required");
									else
										element.className = "required";
								}
							}
						}else{
							if(propertyOn(oldClassName))
								setProperty(element,"OLD-CLASS-NAME",trim(oldClassName.replace(/\brequired\b/gi,"")));
							else if(element.className)
								element.className = trim(element.className.replace(/\brequired\b/gi,""));
						}
					}
				}
				oForm._=true;
			}
			// Set up all form input elements
			for(var j=0,element,iElements=oForm.elements.length;j<iElements;j++){
				element=oForm.elements[j];
				if(!element._){
					if(!isInitialFocusSet && propertyOn((initialFocus=getProperty(element,"INITIAL-FOCUS")))){
						element.focus();
						if(/^select$/i.test(initialFocus)) element.select();
						isInitialFocusSet = true;
					}
					// Capture existing event handlers
					element._onkeypress_=toFunction(getProperty(element,"onkeypress"));
					element._onchange_=toFunction(getProperty(element,"onchange"));
					element._onpropertychange_=toFunction(getProperty(element,"onpropertychange"));

					// Create methods on element
					element.validate=function(oEvent){return isValidElement(this,oEvent);}
					element.onbeforevalidate=toFunction(getProperty(element,"onbeforevalidate"));
					element.onvalidate=toFunction(getProperty(element,"onvalidate"));
					element.onaftervalidate=toFunction(getProperty(element,"onaftervalidate"));
					element.onautosubmit=toFunction(getProperty(element,"onautosubmit"));
					element.onreset=toFunction(getProperty(element,"onreset"));
					element.onpropertychange=function(oEvent){
						oEvent=choose(oEvent, window.event);
						if(oEvent.propertyName && oEvent.propertyName.toUpperCase()=="REQUIRED")
							this.form.markRequired();
						if(this._onpropertychange_)this._onpropertychange_(oEvent);
					}
					// Replace onkeypress event handler
					element.onkeypress=function(oEvent){ //NN passes event object
						if(this._onkeypress_ && this._onkeypress_(oEvent)==false) return false;
						var filter=getProperty(this,"FILTER");
						if(propertyOn(filter)){
							if(filter.constructor!=RegExp)filter=new RegExp(filter);
							oEvent=choose(oEvent, window.event);
							var keyCode=choose(oEvent.which,oEvent.keyCode);
							var keyChar=String.fromCharCode(keyCode);
							if(keyCode!=0 && keyCode!=9 && keyCode!=13 && keyCode!=10 && !filter.test(keyChar))
								return false;
						}
						return true;
					}
					// Replace onchange event handler
					if(element.type!="radio" && element.type!="checkbox"){
						element.onchange=function(oEvent){ //NN4
							oEvent=choose(oEvent, window.event);
							restoreElement(this);
							if(propertyOn(getProperty(this, "VALIDATE-ONCHANGE"),getProperty(this.form, "VALIDATE-ONCHANGE"))){
								if(!this.validate(oEvent) && !suppressOnchange(this))
									return false;
							}
							if(this._onchange_ && this._onchange_()==false)
								return false;
							//auto submit
							var autoSubmit = getProperty(this, "AUTO-SUBMIT");
							if(propertyOn(autoSubmit) && this.onautosubmit()!=false && this.form.onautosubmit()!=false)
								this.form.submit();
						}
					}
					element._=true;
				}
			}
			oForm.markRequired();
		}
	}
	this.Err=new this.Err;
	this.setup();
}
// Limit use of script to valid environments
if(!!window.RegExp && !!"".replace && "ab".replace(/a/,"")=="b" && !!document.forms)
	Validation=new Validation;