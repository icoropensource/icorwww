/*----------------------------------------------------------------------------\
|                             String Builder 1.02                             |
\----------------------------------------------------------------------------*/
 
function StringBuilder(sString) {
	
	// public
	this.length = 0;
	
	this.append = function (sString) {
		// append argument
		this.length += (this._parts[this._current++] = String(sString)).length;
		
		// reset cache
		this._string = null;
		return this;
	};
	
	this.toString = function () {
		if (this._string != null)
			return this._string;
		
		var s = this._parts.join("");
		this._parts = [s];
		this._current = 1;
		this.length = s.length;
		
		return this._string = s;
	};

	// private
	this._current	= 0;
	this._parts		= [];
	this._string	= null;	// used to cache the string
	
	// init
	if (sString != null)
		this.append(sString);
}
